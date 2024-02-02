#!/usr/bin/env -S yarn exec ts-node

// Before running, make sure to source the EXCHANGE_CONFIG environtment variable dependency.
// E.g. `source ~/.gns/credentials`.

import { createHash, getHashes } from "crypto";
import * as ews from "ews-javascript-api";
import { filter, isNumber, size, values } from "lodash";
import { withEwsConnection } from "./ews-connect";

const hash = ((algorithm: string) => {
  if (getHashes().includes(algorithm)) {
    return (text: string) => {
      const hash = createHash(algorithm);
      hash.update(text);
      return hash.digest("hex");
    };
  }
  throw new Error(`${algorithm} is not supported by the current Node.js`);
})("md5");

type LoopItemsOptions = {
  service: ews.ExchangeService;
  folder?: ews.WellKnownFolderName | ews.FolderId;
  after?: ews.DateTime;
  before?: ews.DateTime;
  paging?: number;
  // Try `ews.PropertySet.FirstClassProperties` if you do not know what you need to load
  propertiesToLoad?: ews.PropertySet;
};

const emailsInFolder = async function* ({
  service,
  folder = ews.WellKnownFolderName.Inbox,
  after = ews.DateTime.UtcNow.AddDays(-14),
  before = ews.DateTime.UtcNow,
  paging = 500,
  propertiesToLoad,
}: LoopItemsOptions) {
  const filter = new ews.SearchFilter.SearchFilterCollection(
    ews.LogicalOperator.And,
    [
      new ews.SearchFilter.IsGreaterThan(ews.ItemSchema.DateTimeCreated, after),
      new ews.SearchFilter.IsLessThan(ews.ItemSchema.DateTimeCreated, before),
      new ews.SearchFilter.IsEqualTo(ews.ItemSchema.ItemClass, "IPM.Note"),
    ]
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
    if (propertiesToLoad && search.Items.length > 0) {
      const loadState = await service.LoadPropertiesForItems(
        search.Items,
        propertiesToLoad
      );
      if (loadState.OverallResult !== ews.ServiceResult.Success) {
        throw new Error("Could not load properties for all items.");
      }
    }
    for (const item of search.Items) {
      // As all results have an ItemClass IPM.Note it should be fine to cast:
      yield item as ews.EmailMessage;
    }
    offset += search.Items.length;
    more = search.MoreAvailable;
  } while (more);
};

withEwsConnection(async (service) => {
  const today = new Date();
  const currentDay =
    (Date.UTC(today.getFullYear(), today.getMonth(), today.getDate()) -
      Date.UTC(today.getFullYear(), 0, 0)) /
    24 /
    60 /
    60 /
    1000;
  // console.error(`Today: ${currentDay}`);
  let day = 100;
  do {
    const start = Date.UTC(2022, 0, day, 0, 0, 0);
    const end = Date.UTC(2022, 0, ++day, 0, 0, 0);
    const collection: Record<string, any[]> = {};
    let count = 0;
    for await (const email of emailsInFolder({
      service,
      folder: ews.WellKnownFolderName.SentItems,
      paging: 250,
      propertiesToLoad: new ews.PropertySet([
        ews.EmailMessageSchema.Body,
        ews.EmailMessageSchema.Subject,
        ews.EmailMessageSchema.DisplayTo,
        ews.EmailMessageSchema.DateTimeSent,
        ews.EmailMessageSchema.InternetMessageId,
      ]),
      after: new ews.DateTime(start),
      before: new ews.DateTime(end),
    })) {
      count++;
      if (count % 50 === 0) console.error(count);
      const emailInfo = {
        body: hash(email.Body.Text),
        subject: email.Subject,
        to: email.DisplayTo,
      };
      const identifier = JSON.stringify(values(emailInfo));
      if (collection[identifier] === undefined) {
        collection[identifier] = [];
      }
      collection[identifier].push({
        sent: email.DateTimeSent.ToISOString(),
        imid: email.InternetMessageId,
        ...emailInfo,
      });
    }

    const filteredLength = filter(collection, (item) => item.length > 1).length;
    const unfilteredLength = size(collection);
    console.log(
      [
        new Date(start).toLocaleDateString("sv-SE"),
        filteredLength,
        unfilteredLength,
        filteredLength / unfilteredLength,
      ].join("\t")
    );
  } while (day <= currentDay);
});
