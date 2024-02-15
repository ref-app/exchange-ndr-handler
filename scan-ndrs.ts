#!/usr/bin/env -S yarn exec ts-node
import * as ews from "ews-javascript-api";
import { isNumber } from "lodash";
import { identifiersFromNames, sleep, withEwsConnection } from "./ews-connect";

/**
 * Define the email addresses to look for amongst NDR notes
 */
const lookFor: ReadonlyArray<string> = [
  //   "martijn.vanderven@refapp.com",
];

type LoopItemsOptions = {
  service: ews.ExchangeService;
  folders?: ReadonlyArray<ews.WellKnownFolderName | ews.FolderId>;
  before?: ews.DateTime;
  paging?: number;
  sleepSeconds?: number;
};

const itemsInFolder = async function* ({
  service,
  folders = [ews.WellKnownFolderName.Inbox],
  before = ews.DateTime.UtcNow,
  paging = 500,
  sleepSeconds = 1,
}: LoopItemsOptions) {
  const filter = new ews.SearchFilter.SearchFilterCollection(
    ews.LogicalOperator.And,
    [
      new ews.SearchFilter.ContainsSubstring(
        ews.ContactGroupSchema.ItemClass,
        "NDR"
      ),
      new ews.SearchFilter.IsLessThan(ews.ItemSchema.DateTimeCreated, before),
    ]
  );
  for (const folder of folders) {
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
  }
};

withEwsConnection(async (service) => {
  const ndrFolder = (
    await identifiersFromNames({
      service,
      displayNames: ["NDR Processed"],
    })
  ).pop();
  if (ndrFolder === undefined) {
    throw new Error("Could not find NDR Processed folder.");
  }
  const loadProps = new ews.PropertySet(
    ews.EmailMessageSchema.ToRecipients,
    ews.EmailMessageSchema.DateTimeCreated,
    ews.EmailMessageSchema.TextBody
  );
  for await (const item of itemsInFolder({
    service,
    folders: [ndrFolder.id, ews.WellKnownFolderName.DeletedItems],
  })) {
    const email = await ews.EmailMessage.Bind(service, item.Id, loadProps);
    const hasLookedFor = email.ToRecipients.GetEnumerator().some((address) =>
      lookFor.includes(address.Address)
    );
    if (hasLookedFor) {
      console.log(email.DateTimeCreated.ToISOString());
      console.log(email.TextBody.Text);
      console.log("\n\n---\n\n");
    }
  }
});
