import * as ews from "ews-javascript-api";

export function writeError(error: string) {
  process.stderr.write(`${error}\n`);
}

export function writeProgress(message: string) {
  writeError(message);
}

export function getConfigFromEnvironmentVariable<T>(
  name: string
): T | undefined {
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
  worker: (service: ews.ExchangeService) => Promise<void>
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
    () => writeProgress("Success!"),
    (e) => {
      writeError(`${e.message}`);
      if (e.faultString) {
        writeError(`${e.faultString.faultstring}`);
      }
    }
  );
}
