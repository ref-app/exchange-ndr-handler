#!/usr/bin/env -S yarn exec ts-node
import * as ews from "ews-javascript-api";
import * as _ from "lodash";
import axios from "axios";

import {
  collectionToArray,
  findOrCreateContactGroup,
  getConfigFromEnvironmentVariable,
  withEwsConnection,
  writeError,
  writeProgress,
} from "./ews-connect";
import { createMailjetEvent } from "./mailjet-event";
import { z } from "zod";

type FieldType = "string" | "number" | "date";

function notEmpty<TValue>(value: TValue | null | undefined): value is TValue {
  return value !== null && value !== undefined;
}

function isHardBounce(errorCode: string) {
  return errorCode.startsWith("5.");
}

/**
 * Identity function to get a narrow field name type
 */
const toFieldList = <T extends string>(
  ...items: ReadonlyArray<[T, number, FieldType]>
) => items;

// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagoriginalmessageclass-canonical-property
// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/62366ac9-8c81-45f5-baa9-8b7bfd4db755
const extraFields = toFieldList(
  ["PidTagOriginalMessageClass", 0x004b, "string"],
  ["PidTagOriginalSubject", 0x0049, "string"],
  ["PidTagOriginalSubmitTime", 0x004e, "date"],
  ["PidTagOriginalMessageId", 0x1046, "string"],
  // Would have been nice if the properties below actually existed on NDR items but nope
  ["PidTagNonDeliveryReportStatusCode", 0x0c20, "number"],
  ["PidTagNonDeliveryReportReasonCode", 0x0c04, "number"],
  ["PidTagNonDeliveryReportDiagCode", 0x0c05, "number"]
);

const mapiPropTypes: { [k in FieldType]: ews.MapiPropertyType } = {
  date: ews.MapiPropertyType.SystemTime,
  number: ews.MapiPropertyType.Integer,
  string: ews.MapiPropertyType.String,
};

interface MapiTypes {
  date: number;
  number: number;
  string: string;
}

const mapiProps = extraFields.map(([name, tag, type]) => ({
  name,
  type,
  prop: new ews.ExtendedPropertyDefinition(tag, mapiPropTypes[type]),
}));

const extraProps = [
  ews.ItemSchema.Subject,
  ews.ItemSchema.DateTimeReceived,
  ews.ItemSchema.TextBody,
  ews.ItemSchema.Body,
  ews.ItemSchema.ItemClass,
  ews.EmailMessageSchema.ToRecipients,
  ...mapiProps.map(({ prop }) => prop),
];
const props = new ews.PropertySet(
  ews.BasePropertySet.FirstClassProperties,
  extraProps
);

type OutValue<T extends FieldType> = ews.IOutParam<MapiTypes[T]>;

const initOutValue = <T extends FieldType>(t: T): OutValue<T> => {
  return { outValue: null } as any; // Typings are a bit weird here but null seems to be expected
};

interface FieldValue {
  name: string;
  type: FieldType;
  value: OutValue<FieldType>;
}

function getItemExtendedProperties(item: ews.Item): ReadonlyArray<FieldValue> {
  const values = mapiProps
    .map(({ name, type, prop }): FieldValue | undefined => {
      const value = initOutValue(type);
      if (item.ExtendedProperties.TryGetValue(prop, value)) {
        return { name, type, value };
      }
    })
    .filter(notEmpty);
  return values;
}

async function fetchItemByMessageId(
  service: ews.ExchangeService,
  messageId: string
): Promise<ews.EmailMessage | undefined> {
  // Seems we can only use a (potentially slow) search filter to find an item based on its message id
  const messageIdFilter = new ews.SearchFilter.IsEqualTo(
    ews.EmailMessageSchema.InternetMessageId,
    _.escape(messageId)
  );
  const findResult = await service.FindItems(
    ews.WellKnownFolderName.SentItems,
    messageIdFilter,
    new ews.ItemView(10)
  );
  if (findResult.Items.length === 1) {
    const result = await ews.EmailMessage.Bind(service, findResult.Items[0].Id);
    return result;
  }
  return undefined;
}

// It’s pretty stupid that the error code is not a property of the message itself but
// has to be retrieved from the body text
function extractNdrErrorCode(body: string): string | undefined {
  const matches = /Remote Server returned '([^']+)'/i.exec(body);
  if (matches) {
    const remoteResponse = matches[1];
    // Sometimes the hash is missing, e.g. 550 5.1.17 SMTPSEND.Utf8RecipientAddress; UTF-8 recipient address not supported.
    const codeMatches = /[#\s]([45]\.\d\.\d{1,3})\s/.exec(remoteResponse);
    if (codeMatches) {
      return codeMatches[1];
    }
  }
  return undefined;
}

async function invokeWebhook(
  ndrItem: ews.Item,
  originalMessage: ews.Item,
  originalInternetMessageId: string,
  errorCode: string,
  { webhookUrl, dryRun = false }: Readonly<NdrProcessorConfig>
): Promise<"success" | "failure"> {
  // This is where we could create different kinds of payloads
  const content = createMailjetEvent(
    ndrItem,
    originalMessage,
    originalInternetMessageId,
    errorCode
  );
  try {
    if (dryRun) {
      writeProgress(`Would have invoked webhook for ${originalInternetMessageId}`);
    }
    else {
      const _result = await axios.post(webhookUrl, content);
    }
    return "success";
  } catch (e) {
    return "failure";
  }
}

async function blockRecipients(
  service: ews.ExchangeService,
  recipients: ReadonlyArray<ews.EmailAddress>,
  config: Readonly<NdrProcessorConfig>
) {
  const blockedSendersList = await findOrCreateContactGroup(
    service,
    config.blockedRecipientsListName ?? "Blocked Recipients"
  );
  if (blockedSendersList) {
    const blockedSenders = collectionToArray(blockedSendersList.Members);
    const foundEmailAddresses: ReadonlyArray<string> = blockedSenders.map(
      (member) => member.AddressInformation.Address
    );
    let changed = false;
    for (const recipient of recipients) {
      if (recipient.Address && !foundEmailAddresses.includes(recipient.Address)) {
        if (config.dryRun) {
          writeProgress(
            `Would have added blocked contact ${recipient.Address} to contact group`
          );      
        }
        else {
          writeProgress(
            `Saving new blocked contact ${recipient.Address} to contact group`
          );
          blockedSendersList.Members.AddOneOff(recipient.Name, recipient.Address);
          changed = true;
        }
      }
    }
    if (changed) {
      await blockedSendersList.Update(ews.ConflictResolutionMode.AutoResolve);
    }
  }
}

/**
 * Returns "processed" if the ndr item should be move to the Processed folder
 * because we are ready with it
 */
async function processOneNdrItem(
  service: ews.ExchangeService,
  item: ews.Item,
  values: ReadonlyArray<FieldValue>,
  config: Readonly<NdrProcessorConfig>
): Promise<"processed" | "unprocessed"> {
  if (
    !item.ItemClass.localeCompare("Report.IPM.Note.NDR", undefined, {
      sensitivity: "base",
    })
  ) {
    writeProgress(`NDR item found with Subject: ${item.Subject}`);
    // Load as an email message to get the recipient list for the NDR because only those recipients
    // failed
    const ndrAsEmail = await ews.EmailMessage.Bind(service, item.Id);
    /**
     * NDR RFC 3463 code
     */
    const errorCode = extractNdrErrorCode(item.TextBody.Text);
    if (!errorCode) {
      writeError("Could not extract error code");
    }
    else {
      const messageId = values.find(
        ({ name }) => name === "PidTagOriginalMessageId"
        )?.value.outValue;
      if (messageId && typeof messageId === "string") {
        const originalMessage = await fetchItemByMessageId(service, messageId);
        if (originalMessage) {
          const webhookResult = await invokeWebhook(
            item,
            originalMessage,
            messageId,
            errorCode,
            config
          );
          if (isHardBounce(errorCode)) {
            await blockRecipients(
              service,
              collectionToArray(ndrAsEmail.ToRecipients),
              config
            );
          }

          if (webhookResult === "success") {
            return "processed";
          }
        }
      }
      else {
        writeError("Could not find PidTagOriginalMessageId for message - will move it anyway");
        return "processed";
      }
    }
  }
  return "unprocessed";
}

async function findOrCreateFolder(service: ews.ExchangeService, name: string) {
  const rootFolder = ews.WellKnownFolderName.MsgFolderRoot;
  const filter = new ews.SearchFilter.IsEqualTo(
    ews.FolderSchema.DisplayName,
    name
  );
  const foundFolders = await service.FindFolders(
    rootFolder,
    filter,
    new ews.FolderView(2)
  );
  if (foundFolders.Folders.length > 1) {
    writeError(`Found more than one folder named ${name}`);
    return undefined;
  } else if (foundFolders.Folders.length === 1) {
    return foundFolders.Folders[0];
  }
  const createdFolder = new ews.Folder(service);
  createdFolder.DisplayName = name;
  writeProgress(`Creating folder ${name}`);

  await createdFolder.Save(rootFolder);
  return createdFolder;
}

const ndrProcessorConfigSchema = z.object({
  processedFolderName: z.string().optional(),
  blockedRecipientsListName: z.string().optional(),
  webhookUrl: z.string(),
  dryRun: z.boolean().optional(),
})
.strict()
.readonly();

type NdrProcessorConfig = z.infer<typeof ndrProcessorConfigSchema>

/**
 * Searching by query may be faster but there seems to be a huge delay (more than 10 minutes, then I gave up) between e.g. moving items between folders in Outlook Web Access
 * and when the items actually show up in a search so we use a filter because it seems reliable
 */
async function findItemsByQueryOrFilter(
  service: ews.ExchangeService,
  query: string,
  filter: ews.SearchFilter,
  view: ews.ItemView
) {
  // Not: const result = service.FindItems(ews.WellKnownFolderName.Inbox, query, view);
  const result = await service.FindItems(
    ews.WellKnownFolderName.Inbox,
    filter,
    view
  );
  return result;
}

async function processNdrMessages(service: ews.ExchangeService) {
  // References:
  // https://stackoverflow.com/questions/12176360/get-original-message-headers-using-ews-for-bounced-emails
  // https://docs.microsoft.com/en-us/previous-versions/office/developer/exchange-server-2010/ee693615%28v%3dexchg.140%29
  // kind:Report is not documented - found through trial and error
  const query = `kind:report`;
  const filter = new ews.SearchFilter.SearchFilterCollection(ews.LogicalOperator.And,
    [
      new ews.SearchFilter.IsEqualTo(
        ews.ItemSchema.ItemClass,
        "REPORT.IPM.Note.NDR"),
    ]);

  let offset = 0;

  const configRaw = getConfigFromEnvironmentVariable<unknown>(
    "NDR_PROCESSOR_CONFIG"
  );
  if (!configRaw) {
    writeError("Error: NDR_PROCESSOR_CONFIG environment variable must be set");
    process.exit(4);
  }
  const parsed = ndrProcessorConfigSchema.safeParse(configRaw);
  if (!parsed.success)  {
    writeError("Error: configuration did not validate");
    for (const error of parsed.error.issues) {
      writeError(`${error.message} ${error.path}`);
    }
    process.exit(4);
  }

  const processorConfig = parsed.data;
  const processedFolder = await findOrCreateFolder(
    service,
    processorConfig.processedFolderName ?? "NDR Processed"
  );
  if (!processedFolder) {
    writeError(
      "Could not find or create folder for processed items - aborting"
    );
    process.exit(2);
  }
  do {
    const view = new ews.ItemView(10, offset);
    const found = await findItemsByQueryOrFilter(service, query, filter, view);
    if (found.Items.length > 0) {
      await service.LoadPropertiesForItems(found.Items, props);
      for (const item of found.Items) {
        const processResult = await processOneNdrItem(
          service,
          item,
          getItemExtendedProperties(item),
          processorConfig
        );
        if (processResult === "processed") {
          if (processorConfig.dryRun) {
            writeProgress("Would have moved item to processed documents")
          }
          else {
            await item.Move(processedFolder.Id);
          }
        }
      }
    }
    if (!found.MoreAvailable) {
      break;
    }
    offset = found.NextPageOffset;
  } while (true);
}

writeProgress("process-ndr-messages");
withEwsConnection(processNdrMessages);
