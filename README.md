# exchange-ndr-handler

The core of this project is a script (written in Typescropt) that handles non-deliverable messages found in an Exchange account inbox, e.g. "Recipient cannot be found".

For each such item, the script performs a number of subtasks

* It searches for the original item in the "Sent Items" folder and if it is found, it calls a configurable webhook with a payload that includes the original message UniqueId and the error code.
* For permanent errors (error codes 5.x.x), it adds the To address to the Exchange address book with a Category BLOCKED (if it doesn’t already exist).
* Finally, it moves the NDR message from Inbox to a folder "NDR Processed".

## Installation
The project is packaged as a Docker image ready to run, e.g. as a Cron job in a Kubernetes cluster.
The following environment variables must be set for the script to run properly:

###EXCHANGE_CONFIG
Json document string
```
{
    "username": "username",
    "password": "password",
    "serviceUrl": "https://your.exchange-server.com/EWS/Exchange.asmx"
}
```


## Run locally

Prerequisites: nodejs, yarn and ts-node with typescript

```
yarn global add ts-node typescript
```

```
yarn install
./process-ndr-messages.ts
```

## Build

docker build .