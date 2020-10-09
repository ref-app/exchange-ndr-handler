#!/usr/bin/env ts-node-script
import * as ews from "ews-javascript-api";
import * as _ from "lodash";
import { withEwsConnection, writeProgress } from "./ews-connect";

type FieldType = "string" | "number" | "date";

// https://docs.microsoft.com/en-us/office/client-developer/outlook/mapi/pidtagoriginalmessageclass-canonical-property
// https://docs.microsoft.com/en-us/openspecs/exchange_server_protocols/ms-oxomsg/62366ac9-8c81-45f5-baa9-8b7bfd4db755

const extraFields: [string, number, FieldType][] = [
  ["PidTagOriginalMessageClass", 0x004b, "string"],
  ["PidTagOriginalSubject", 0x0049, "string"],
  ["PidTagOriginalSubmitTime", 0x004e, "date"],
  ["PidTagOriginalMessageId", 0x1046, "string"],
  // Would have been nice if the properties below actually existed on NDR items but no
  ["PidTagNonDeliveryReportStatusCode", 0x0c20, "number"],
  ["PidTagNonDeliveryReportReasonCode", 0x0c04, "number"],
  ["PidTagNonDeliveryReportDiagCode", 0x0c05, "number"],
];

const mapiPropTypes: { [k in FieldType]: ews.MapiPropertyType } = {
  date: ews.MapiPropertyType.SystemTime,
  number: ews.MapiPropertyType.Integer,
  string: ews.MapiPropertyType.String,
};

interface mapiTypes {
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
  ...mapiProps.map(({ prop }) => prop),
];
const props = new ews.PropertySet(
  ews.BasePropertySet.FirstClassProperties,
  extraProps
);

type OutValue<T extends FieldType> = ews.IOutParam<mapiTypes[T]>;

const initOutValue = <T extends FieldType>(t: T): OutValue<T> => {
  return { outValue: null } as OutValue<T>;
};

interface FieldValue {
  name: string;
  type: FieldType;
  value: OutValue<FieldType>;
}

/**
 * Loads item properties and invokes an action
 */
async function processFoundItems(
  service: ews.ExchangeService,
  items: ews.Item[],
  itemAction?: (
    service: ews.ExchangeService,
    item: ews.Item,
    values: FieldValue[]
  ) => Promise<void>
) {
  await service.LoadPropertiesForItems(items, props);

  for (const item of items) {
    writeProgress("");
    writeProgress(`subject: ${item.Subject}`);
    const allValues = mapiProps.map(({ name, type, prop }) => {
      const value = initOutValue(type);
      item.ExtendedProperties.TryGetValue(prop, value);
      return { name, type, value };
    });
    if (itemAction) {
      await itemAction(service, item, allValues);
    }
  }
}

async function processOriginalMessageForNDR(
  service: ews.ExchangeService,
  item: ews.Item,
  values: FieldValue[]
) {
  writeProgress("Found the original in SentItems!");
  writeProgress(`Its unique id is ${item.Id.UniqueId}`);
}

async function fetchItemByMessageId(
  service: ews.ExchangeService,
  messageId: string
) {
  writeProgress("");
  writeProgress(`Searching for ${messageId}`);
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
    await processFoundItems(
      service,
      findResult.Items,
      processOriginalMessageForNDR
    );
  }
}

async function processOneNdrItem(
  service: ews.ExchangeService,
  item: ews.Item,
  values: FieldValue[]
) {
  if (
    !item.ItemClass.localeCompare("Report.IPM.Note.NDR", undefined, {
      sensitivity: "base",
    })
  ) {
    const matches = /Remote Server returned '([^']+)'/i.exec(
      item.TextBody.Text
    );
    if (matches) {
      const remoteResponse = matches[1];
      const codeMatches = /#([45]\.[0-9]\.[0-9]{1,3})/.exec(remoteResponse);
      if (codeMatches) {
        writeProgress(`NDR RFC 3463 code: ${codeMatches[1]}`);
      }
    }
  }
  const messageId = values.find(
    ({ name }) => name === "PidTagOriginalMessageId"
  )?.value.outValue;
  if (messageId && typeof messageId === "string") {
    await fetchItemByMessageId(service, messageId);
  }
}

async function processNdrMessages(service: ews.ExchangeService) {
  // https://stackoverflow.com/questions/12176360/get-original-message-headers-using-ews-for-bounced-emails
  const sf = new ews.SearchFilter.IsEqualTo(
    ews.ItemSchema.ItemClass,
    "REPORT.IPM.Note.NDR"
  );
  let offset = 0;
  do {
    const view = new ews.ItemView(10, offset);
    const found = await service.FindItems(
      ews.WellKnownFolderName.Inbox,
      sf,
      view
    );
    await processFoundItems(service, found.Items, processOneNdrItem);
    if (!found.MoreAvailable) {
      break;
    }
    offset = found.NextPageOffset;
  } while (true);
}

withEwsConnection(processNdrMessages);
