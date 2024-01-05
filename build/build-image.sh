#!/usr/bin/env bash
DIRECTORY=$(cd `dirname $0` && pwd)
VERSION=$(cat "${DIRECTORY}/../VERSION")

docker buildx build --platform linux/amd64,linux/arm64 -t exchange-ndr-handler:latest -t exchange-ndr-handler:${VERSION} .
docker buildx build --platform linux/amd64,linux/arm64 -t inbox-notifier:latest -t inbox-notifier:${VERSION} --build-arg SCRIPT="inbox-notification.ts" .