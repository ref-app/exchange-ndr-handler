#!/usr/bin/env -S yarn exec -s ts-node

// Before running, make sure to source the EXCHANGE_CONFIG environtment variable dependency.
// E.g. `source ~/.gns/credentials`.

import * as ews from "ews-javascript-api";
import { withEwsConnection } from "./ews-connect";

const thisSystem = "exchange";

const findContactGroup = async (
  service: ews.ExchangeService,
  name: string
): Promise<ews.ContactGroup | undefined> => {
  const filter = new ews.SearchFilter.SearchFilterCollection(
    ews.LogicalOperator.And,
    [
      new ews.SearchFilter.IsEqualTo(ews.ContactGroupSchema.DisplayName, name),
      new ews.SearchFilter.IsEqualTo(
        ews.ContactGroupSchema.ItemClass,
        "IPM.DistList"
      ),
    ]
  );
  const found = await service.FindItems(
    ews.WellKnownFolderName.Contacts,
    filter,
    new ews.ItemView(2)
  );
  if (found.Items.length > 1) {
    console.error({
      system: thisSystem,
      message: `Found more than one contact group named ${name}`,
    });
    return undefined;
  } else if (found.Items.length === 1) {
    return ews.ContactGroup.Bind(service, found.Items[0].Id);
  }
  return undefined;
};

function collectionToArray<T extends ews.ComplexProperty>(
  collection: ews.ComplexPropertyCollection<T>
): ReadonlyArray<T> {
  const result = Array.from({ length: collection.Count }).map((n, i) =>
    collection._getItem(i)
  );
  return result;
}

const getBlockedRecipients = async (
  service: ews.ExchangeService
): Promise<Set<string>> => {
  const contactGroup = await findContactGroup(service, "Blocked Recipients");
  const blockedEmails = contactGroup
    ? collectionToArray(contactGroup.Members).map(
        (m) => m.AddressInformation.Address
      )
    : [];
  const blockedEmailsSet = new Set(blockedEmails);
  return blockedEmailsSet;
};

withEwsConnection(async (service) => {
  console.log(await getBlockedRecipients(service));
});
