import * as ews from "ews-javascript-api";

export function writeError(error: string) {
  process.stderr.write(`${error}\n`);
}

export function writeProgress(message: string) {
  writeError(message);
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
  const configRaw = process.env["EXCHANGE_CONFIG"];
  if (!configRaw) {
    writeError("Error: EXCHANGE_CONFIG environment variable must be set");
    process.exit(-2);
  }
  const config: ExchangeConfig = JSON.parse(configRaw);

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
