#!/usr/bin/env bash
DIRECTORY=$(cd `dirname $0` && pwd)
VERSION=$(cat "${DIRECTORY}/../VERSION")

# you may need to run these first:
# docker buildx create --name mybuilder
# docker buildx use mybuilder
# https://cloudolife.com/2022/03/05/Infrastructure-as-Code-IaC/Container/Docker/Docker-buildx-support-multiple-architectures-images/

docker buildx build --push --platform linux/amd64,linux/arm64 -t twrefapp/exchange-ndr-handler:${VERSION} .
docker buildx build --push --platform linux/amd64,linux/arm64 -t twrefapp/inbox-notifier:${VERSION} --build-arg SCRIPT="inbox-notification.ts" .