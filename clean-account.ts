#!/usr/bin/env -S yarn exec ts-node
import * as ews from "ews-javascript-api";
import { isNumber } from "lodash";
import { argv, exit, stderr, stdout } from "process";
import {
  type Identifier,
  identifiersFromNames,
  withEwsConnection,
  sleep,
} from "./ews-connect";

/**
 * Maps different {@link ews.ServiceResult} values to a progress symbol,
 * inspired by test runners: period for pass, W for warning, and E for error.
 */
const resultsMap: Record<ews.ServiceResult, string> = {
  0: ".",
  1: "W",
  2: "E",
};

type LoopItemsOptions = {
  service: ews.ExchangeService;
  folder?: ews.WellKnownFolderName | ews.FolderId;
  before?: ews.DateTime;
  paging?: number;
  sleepSeconds?: number;
};

const itemsInFolder = async function* ({
  service,
  folder = ews.WellKnownFolderName.Inbox,
  before = ews.DateTime.UtcNow,
  paging = 500,
  sleepSeconds = 1,
}: LoopItemsOptions) {
  const filter = new ews.SearchFilter.IsLessThan(
    ews.ItemSchema.DateTimeCreated,
    before
  );
  let offset = 0;
  let more = false;
  do {
    const view = new ews.ItemView(paging, offset);
    /* TypeScript messing with me. `FindItems` has separate declarations
     * with `WellKnownFolderName` and `FolderId` as first arguments, but is
     * never typed to take in `WellKnownFolderName | FolderId`.
     */
    const search = isNumber(folder)
      ? await service.FindItems(folder, filter, view)
      : await service.FindItems(folder, filter, view);
    for (const item of search.Items) {
      yield item;
    }
    offset += search.Items.length;
    more = search.MoreAvailable;
    // Before we request another page of items, give it a sleep.
    await sleep({ ms: sleepSeconds * 1000 });
  } while (more);
};

type PurgeItemsOptions = {
  service: ews.ExchangeService;
  folderIdentifier: ews.WellKnownFolderName | Identifier;
  before: ews.DateTime;
  deleteMode: ews.DeleteMode.HardDelete | ews.DeleteMode.MoveToDeletedItems;
  sleepSeconds?: number;
};

const purgeItems = async ({
  service,
  folderIdentifier,
  before,
  deleteMode,
  sleepSeconds = 1,
}: PurgeItemsOptions) => {
  const displayName = isNumber(folderIdentifier)
    ? ews.WellKnownFolderName[folderIdentifier]
    : folderIdentifier.displayName;
  const folder = isNumber(folderIdentifier)
    ? folderIdentifier
    : folderIdentifier.id;
  stdout.write(
    `${
      deleteMode === ews.DeleteMode.HardDelete
        ? "Deleting items"
        : "Moving items to DeletedItems"
    } from ${displayName} where the creation date is before ${before.ToString()}.\n`
  );
  let purgedItems = 0;
  try {
    let itemIdsForPurge: ews.ItemId[] = [];
    for await (const item of itemsInFolder({
      service,
      folder,
      before,
      sleepSeconds,
    })) {
      itemIdsForPurge.push(item.Id);
      if (itemIdsForPurge.length === 100) {
        const results = await service.DeleteItems(
          itemIdsForPurge,
          deleteMode,
          ews.SendCancellationsMode.SendToNone,
          ews.AffectedTaskOccurrence.SpecifiedOccurrenceOnly
        );
        // After deleting a bunch, give it a sleep.
        await sleep({ ms: sleepSeconds * 1000 });
        itemIdsForPurge = [];
        purgedItems += results.Responses.length;
        stderr.write(
          results.Responses.map((response) => resultsMap[response.Result]).join(
            ""
          )
        );
      }
    }
    if (itemIdsForPurge.length > 0) {
      const results = await service.DeleteItems(
        itemIdsForPurge,
        deleteMode,
        ews.SendCancellationsMode.SendToNone,
        ews.AffectedTaskOccurrence.SpecifiedOccurrenceOnly
      );
      // After deleting a bunch, give it a sleep.
      await sleep({ ms: sleepSeconds * 1000 });
      purgedItems += results.Responses.length;
      stderr.write(
        results.Responses.map((response) => resultsMap[response.Result]).join(
          ""
        ) + "\n"
      );
    } else if (purgedItems > 0) {
      stderr.write("\n"); // Close out the progress line.
    }
  } catch (error) {
    stderr.write("\n");
    throw error;
  } finally {
    stdout.write(
      `${
        deleteMode === ews.DeleteMode.HardDelete ? "Deleted" : "Moved"
      } ${purgedItems} items.\n`
    );
  }
};

let purgeBefore: number;
let keepTrashMonths = 3;

/**
 * When running in production, do not take the date from the CLI. Instead:
 * 1. Move all data older than 1 month into the trash,
 * 2. Only keep 2 months of data in the trash.
 */
if (process.env.NODE_ENV === "production") {
  const date = new Date();
  date.setMonth(date.getMonth() - 1);
  purgeBefore = date.getTime();
  keepTrashMonths = 2;
} else {
  purgeBefore = Date.parse(argv[2]);
}
if (isNaN(purgeBefore)) {
  stderr.write("Invalid date argument\n");
  exit(1);
}

withEwsConnection(async (service) => {
  // Move items from the following folders to DeletedItems if older than the provided purgeBefore date.
  const before = new ews.DateTime(purgeBefore);
  for (const folderIdentifier of [
    ews.WellKnownFolderName.Inbox,
    ews.WellKnownFolderName.SentItems,
    ews.WellKnownFolderName.Calendar,
    ...(await identifiersFromNames({
      service,
      displayNames: ["NDR Processed", "Flood warnings"],
    })),
  ]) {
    await purgeItems({
      service,
      folderIdentifier,
      before,
      deleteMode: ews.DeleteMode.MoveToDeletedItems,
      sleepSeconds: 2,
    });
  }
  // Then remove everything from the trash older than 3 months before the provided purgeBefore date.
  await purgeItems({
    service,
    folderIdentifier: ews.WellKnownFolderName.DeletedItems,
    before: before.AddMonths(-1 * keepTrashMonths),
    deleteMode: ews.DeleteMode.HardDelete,
    sleepSeconds: 2,
  });
});
