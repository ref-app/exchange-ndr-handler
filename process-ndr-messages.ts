#!/usr/bin/env ts-node-script
import * as ews from "ews-javascript-api";
import * as _ from "lodash";
import { withEwsConnection } from "./ews-connect";

async function processNdrMessages(service: ews.ExchangeService) {}

withEwsConnection(processNdrMessages);
