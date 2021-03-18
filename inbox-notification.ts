#!/usr/bin/env ts-node-script
import axios from "axios";
import * as ews from "ews-javascript-api";
import { URL } from "url";
import {
  getConfigFromEnvironmentVariable,
  withEwsConnection,
  writeError,
} from "./ews-connect";

async function invokeWebhook(
  numNewMessages: number,
  webhookUrl: string,
  mailboxName: string,
  mailboxLink: string
): Promise<"success" | "failure"> {
  // This is where we could create different kinds of payloads
  const content = {
    message: "You have new support mail",
    numNewMessages,
    text: `You have ${numNewMessages} new message${
      numNewMessages > 1 ? "s" : ""
    } in the ${mailboxName} mailbox - click here to access: ${mailboxLink}`,
  };
  console.info(content);
  try {
    await axios.post(webhookUrl, content);
    return "success";
  } catch (e) {
    return "failure";
  }
}

interface InboxNotificationConfig {
  processedTag: string;
  webhookUrl: string;
  mailboxName: string;
}
/**
 * Searching by query may be faster but there seems to be a huge delay (more than 10 minutes, then I gave up) between e.g. moving items between folders in Outlook Web Access
 * and when the items actually show up in a search so we use a filter because it seems reliable
 */
async function findItemsByFilter(
  service: ews.ExchangeService,
  filter: ews.SearchFilter,
  view: ews.ItemView
) {
  //const result = service.FindItems(ews.WellKnownFolderName.Inbox, query, view);
  const result = await service.FindItems(
    ews.WellKnownFolderName.Inbox,
    filter,
    view
  );
  return result;
}

async function processInbox(service: ews.ExchangeService) {
  let offset = 0;

  const processorConfig = getConfigFromEnvironmentVariable<InboxNotificationConfig>(
    "INBOX_NOTIFICATION_CONFIG"
  );
  if (!processorConfig) {
    writeError(
      "Error: INBOX_NOTIFICATION_CONFIG environment variable must be set"
    );
    process.exit(4);
  }
  const filter = new ews.SearchFilter.Not(
    new ews.SearchFilter.IsEqualTo(
      ews.ItemSchema.Categories,
      processorConfig.processedTag
    )
  );

  let numNew = 0;
  do {
    const view = new ews.ItemView(100, offset);
    const found = await findItemsByFilter(service, filter, view);
    if (found.Items.length > 0) {
      for (const item of found.Items) {
        item.Categories.Add(processorConfig.processedTag);
        await item.Update(ews.ConflictResolutionMode.AutoResolve);
        numNew++;
      }
    }

    if (!found.MoreAvailable) {
      break;
    }
    offset = found.NextPageOffset;
  } while (true);
  if (numNew > 0) {
    const mailboxLink = new URL(service.Url.AbsoluteUri);
    mailboxLink.pathname = "/owa";
    await invokeWebhook(
      numNew,
      processorConfig.webhookUrl,
      processorConfig.mailboxName,
      mailboxLink.toString()
    );
  }
}

console.info("inbox-notification");
withEwsConnection(processInbox);
