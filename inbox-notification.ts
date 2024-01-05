#!/usr/bin/env -S yarn exec ts-node
import axios from "axios";
import * as ews from "ews-javascript-api";
import { URL } from "url";
import {
  getConfigFromEnvironmentVariable,
  withEwsConnection,
  writeError,
} from "./ews-connect";

interface InboxNotificationConfig {
  processedTag: string;
  /**
   * Webhook to invoke, e.g. for a Slack channel
   */
  webhookUrl: string;
  /**
   * Display name for the notification
   */
  mailboxName: string;
  /**
   * Folder to search in or default to Inbox
   */
  folderName?: string;
  customMessage?: string;
  customText?: string;
}

async function invokeWebhook({
  numNewMessages,
  mailboxLink,
  config: { mailboxName, webhookUrl, customMessage, customText },
}: {
  numNewMessages: number;
  mailboxLink: string;
  config: InboxNotificationConfig;
}): Promise<"success" | "failure"> {
  // This is where we could create different kinds of payloads
  const content = {
    message: customMessage || "You have new support mail",
    numNewMessages,
    text:
      customText ||
      `You have ${numNewMessages} new message${
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

/**
 * Searching by query may be faster but there seems to be a huge delay (more than 10 minutes, then I gave up) between e.g. moving items between folders in Outlook Web Access
 * and when the items actually show up in a search so we use a filter because it seems reliable
 */
async function findItemsByFilter(
  service: ews.ExchangeService,
  folderId: ews.FolderId,
  filter: ews.SearchFilter,
  view: ews.ItemView
) {
  const result = await service.FindItems(folderId, filter, view);
  return result;
}

/**
 * Falls back to inbox if no name given
 */
async function getFolderIdFromName(
  service: ews.ExchangeService,
  folderName: string | undefined
): Promise<ews.FolderId> {
  if (folderName) {
    const view = new ews.FolderView(1);
    view.PropertySet = new ews.PropertySet(
      ews.BasePropertySet.IdOnly,
      ews.FolderSchema.DisplayName
    );
    view.Traversal = ews.FolderTraversal.Deep;

    const filter = new ews.SearchFilter.IsEqualTo(
      ews.FolderSchema.DisplayName,
      folderName
    );

    const found = await service.FindFolders(
      ews.WellKnownFolderName.Root,
      filter,
      view
    );
    if (found.TotalCount > 0) {
      return found.Folders[0].Id;
    }
    // Throw?
  }
  return new ews.FolderId(ews.WellKnownFolderName.Inbox);
}

async function processInbox(service: ews.ExchangeService) {
  let offset = 0;

  const processorConfig =
    getConfigFromEnvironmentVariable<InboxNotificationConfig>(
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

  const parentFolderId = await getFolderIdFromName(
    service,
    processorConfig.folderName
  );

  let numNew = 0;
  do {
    const view = new ews.ItemView(100, offset);
    const found = await findItemsByFilter(
      service,
      parentFolderId,
      filter,
      view
    );
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
    await invokeWebhook({
      config: processorConfig,
      numNewMessages: numNew,
      mailboxLink: mailboxLink.toString(),
    });
  }
}

console.info("inbox-notification");
withEwsConnection(processInbox);
