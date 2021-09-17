import * as ews from "ews-javascript-api";

export function writeError(error: string) {
  process.stderr.write(`${error}\n`);
}

export function writeProgress(message: string) {
  writeError(message);
}

export function getConfigFromEnvironmentVariable<T>(
  name: string
): Readonly<T> | undefined {
  const configRaw = process.env[name];
  if (configRaw) {
    try {
      // Schema validation left to receiver
      const config: T = JSON.parse(configRaw);
      return config;
    } catch (e) {}
  }
  return undefined;
}

export function withEwsConnection(
  worker: (service: ews.ExchangeService) => Promise<void|number>
) {
  interface ExchangeConfig {
    username: string;
    password: string;
    serviceUrl: string;
  }
  const service = new ews.ExchangeService(ews.ExchangeVersion.Exchange2016);

  const config = getConfigFromEnvironmentVariable<ExchangeConfig>(
    "EXCHANGE_CONFIG"
  );
  if (!config) {
    writeError("Error: EXCHANGE_CONFIG environment variable must be set");
    process.exit(4);
  }

  const username = config.username;
  const password = config.password;

  service.Credentials = new ews.WebCredentials(username, password);
  service.TraceEnabled = true;
  service.TraceFlags = ews.TraceFlags.All;

  service.Url = new ews.Uri(config.serviceUrl);

  worker(service).then(
    (result) => {
      writeProgress("Success!");
      return result;
    },
    (e) => {
      writeError(`Error: ${e.message}`);
      if (e.faultString) {
        writeError(`Fault: ${e.faultString.faultstring}`);
      }
      if (e.stack) {
        writeError(e.stack.toString());
      }
    }
  );
}

export async function findOrCreateContactGroup(
  service: ews.ExchangeService,
  name: string
) {
  const rootFolder = ews.WellKnownFolderName.Contacts;
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
    rootFolder,
    filter,
    new ews.ItemView(2)
  );
  if (found.Items.length > 1) {
    writeError(`Found more than one contact group named ${name}`);
    return undefined;
  } else if (found.Items.length === 1) {
    return await ews.ContactGroup.Bind(service, found.Items[0].Id);
  }
  const createdItem = new ews.ContactGroup(service);
  createdItem.DisplayName = name;
  writeProgress(`Creating contact group ${name}`);

  await createdItem.Save(rootFolder);
  return createdItem;
}

export function collectionToArray<T extends ews.ComplexProperty>(
  collection: ews.ComplexPropertyCollection<T>
): ReadonlyArray<T> {
  const result = Array.from({ length: collection.Count }).map((n, i) =>
    collection._getItem(i)
  );
  return result;
}
