#!/usr/bin/env ts-node-script
import { SIGUSR1 } from "constants";
import * as ews from "ews-javascript-api";
import _ = require("lodash");
import { withEwsConnection, writeProgress } from "./ews-connect";

async function countSentMail(service: ews.ExchangeService) {
  const now = new Date();
  const daysBack = 100;
  const cutoffDate = new Date(now).setDate(now.getDate() - daysBack);
  const ewsDateTime = new ews.DateTime(cutoffDate);
  const filter = new ews.SearchFilter.IsGreaterThan(
    ews.ItemSchema.DateTimeCreated,
    ewsDateTime
  );
  console.info(`Looking for sent emails after ${ewsDateTime.ToString()}`);
  let offset = 0;
  const dates: string[] = [];
  do {
    const view = new ews.ItemView(1000, offset);
    view.PropertySet = new ews.PropertySet(ews.ItemSchema.DateTimeSent);
    const found = await service.FindItems(
      ews.WellKnownFolderName.SentItems,
      filter,
      view
    );
    dates.push(
      ...found.Items.map((item) => {
        const sentOn = item.DateTimeSent;
        return sentOn.Format("yy-MM-DD");
      })
    );

    if (!found.MoreAvailable) {
      break;
    }
    offset = found.NextPageOffset;
  } while (true);
  const perDay = _.countBy(dates);
  const sortedPairs = _.toPairs(perDay).sort(([k1], [k2]) =>
    k1.localeCompare(k2)
  );
  for (const [k, v] of sortedPairs) {
    console.info(`${k}\t${v}`);
  }
}

writeProgress("count-sent-mail");
withEwsConnection(countSentMail);
