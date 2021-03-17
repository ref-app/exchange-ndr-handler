# exchange-ndr-handler

## Why did we build this?
Because of privacy concerns and specifically concerns about third party governments being able to issue secret court orders to extract personal data, we were looking for a transactional mail service hosted in the EU by a company headquartered in the EU.

Unfortunately, we couldn’t find any. All services seemed to be run by US companies (Mailjet seemed to be the last European supplier but it was gobbled up by US Mailgun in 2019).

So then we couldn’t come up with anything better than just buying a subscription to a normal mailbox, and found a supplier that had a hosted Microsoft Exchange service.

Compared to the capabilities of normal Transactional Email API vendors, these services lack some critical pieces in handling of bounced mail. We wanted both a callback when a mail message bounces and also be able to block addresses for which we get a permanent sending failure (hard bounce).

That’s where this component steps in to complete the picture.

## What does it do?
The core of this project is a script (written in Typescript) that handles non-deliverable messages found in an Exchange account inbox, e.g. "Recipient cannot be found".

For each such item, the script performs a number of subtasks

* It searches for the original item in the "Sent Items" folder and if it is found, it calls a configurable webhook with a payload that includes the original message UniqueId and the error code. The payload is formatted as a Mailjet event.
* For permanent errors (error codes 5.x.x), it adds the To address to a named distribution list so that your code that sends mail can do a lookup there and fail early instead of resending to the same incorrect email address and damaging your reputation
* Finally, it moves the NDR message from Inbox to a folder named "NDR Processed" (configurable).

## Installation and configuration
The project is packaged as a Docker image ready to run, e.g. as a Cron job in a Kubernetes cluster.

Download from docker hub:

twrefapp/exchange-ndr-handler:latest

The following environment variables must be set for the script to run properly:

### EXCHANGE_CONFIG
Json document string
```
{
    "username": "username",
    "password": "password",
    "serviceUrl": "https://your.exchange-server.com/EWS/Exchange.asmx"
}
```

### NDR_PROCESSOR_CONFIG
Json document string
```
{
    "processedFolderName": "NDR Processed",
    "blockedRecipientsListName": "Blocked Recipients",
    "webhookUrl": "https://webhook-receiver"
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

build/build-image.sh