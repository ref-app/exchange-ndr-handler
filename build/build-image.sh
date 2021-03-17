#!/usr/bin/env bash
DIRECTORY=$(cd `dirname $0` && pwd)
VERSION=$(cat "${DIRECTORY}/../VERSION")

docker build -t exchange-ndr-handler:latest -t exchange-ndr-handler:${VERSION} . 