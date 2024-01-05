#!/usr/bin/env bash
DIRECTORY=$(cd `dirname $0` && pwd)
VERSION=$(cat "${DIRECTORY}/../VERSION")

docker buildx build --load -t exchange-ndr-handler:latest -t exchange-ndr-handler:${VERSION} .
docker buildx build --load -t inbox-notifier:latest -t inbox-notifier:${VERSION} --build-arg SCRIPT="inbox-notification.ts" .