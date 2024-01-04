#!/usr/bin/env -S yarn exec ts-node
import * as ews from "ews-javascript-api";
import {
  collectionToArray,
  findOrCreateContactGroup,
  getConfigFromEnvironmentVariable,
  withEwsConnection,
  writeError,
  writeProgress,
} from "./ews-connect";

/**
 * Subset
 */
export interface NdrProcessorConfig {
  blockedRecipientsListName?: string;
}

async function getPreblockedAddresses(service: ews.ExchangeService) {
  const processorConfig = getConfigFromEnvironmentVariable<NdrProcessorConfig>(
    "NDR_PROCESSOR_CONFIG"
  );
  if (!processorConfig) {
    writeError("Error: NDR_PROCESSOR_CONFIG environment variable must be set");
    process.exit(4);
  }

  const blockedSendersList = await findOrCreateContactGroup(
    service,
    processorConfig.blockedRecipientsListName ?? "Blocked Recipients"
  );

  if (blockedSendersList) {
    const members = collectionToArray(blockedSendersList.Members);
    for (const member of members) {
      process.stdout.write(member.AddressInformation.Address + "\n");
    }
  }
}

writeProgress("get-blocked-recipients");
withEwsConnection(getPreblockedAddresses);
