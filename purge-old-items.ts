#!/usr/bin/env ts-node-script
import * as ews from "ews-javascript-api";
import { withEwsConnection, writeError } from "./ews-connect";

/**
 * Searching by query may be faster but there seems to be a huge delay (more than 10 minutes, then I gave up) between e.g. moving items between folders in Outlook Web Access
 * and when the items actually show up in a search so we use a filter because it seems reliable
 */
async function findItemsByFilter(
  service: ews.ExchangeService,
  folder: ews.WellKnownFolderName,
  filter: ews.SearchFilter,
  view: ews.ItemView
) {
  const result = await service.FindItems(folder, filter, view);
  return result;
}

const resultsMap: Record<ews.ServiceResult, string> = {
  0: ".",
  1: "W",
  2: "E",
}

async function purgeItems(
  service: ews.ExchangeService,
  folderName: ews.WellKnownFolderName,
  filter: ews.SearchFilter
) {
  let offset = 0;
  let numDeleted = 0;
  do {
    const view = new ews.ItemView(1000, offset);
    const found = await findItemsByFilter(service, folderName, filter, view);
    if (found.Items.length > 0) {
      console.info(
        `Found ${found.Items.length} items in ${folderName}, moving to Deleted Items`
      );
      const results = await service.MoveItems(
        found.Items.map(item => item.Id),
        new ews.FolderId(ews.WellKnownFolderName.DeletedItems)
      );
      numDeleted += results.Count;
      process.stderr.write(results.Responses.map(result => resultsMap[result.Result]).join("") + "\n");
    }

    if (!found.MoreAvailable) {
      break;
    }
    offset = found.NextPageOffset;
  } while (true);
  console.info(`Deleted ${numDeleted} items total from ${folderName}`);
}

async function processItems(service: ews.ExchangeService) {
  const dateArg = process.argv[2];
  const cutoffDate = Date.parse(dateArg);
  if (isNaN(cutoffDate)) {
    writeError("Invalid date argument");
    return -1;
  }
  const ewsDateTime = new ews.DateTime(cutoffDate);
  console.info(`Cutoff date is ${ewsDateTime.ToString()}`);

  const filter = new ews.SearchFilter.IsLessThan(
    ews.ItemSchema.DateTimeCreated,
    ewsDateTime
  );

  for (const folderName of [
    ews.WellKnownFolderName.Inbox,
    ews.WellKnownFolderName.SentItems,
    ews.WellKnownFolderName.Calendar,
  ]) {
    await purgeItems(service, folderName, filter);
  }
}

console.info("purge-old-items");
withEwsConnection(processItems);
