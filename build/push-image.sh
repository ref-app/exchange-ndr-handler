#!/usr/bin/env bash
DIRECTORY=$(cd `dirname $0` && pwd)
VERSION=$(cat "${DIRECTORY}/../VERSION")

for tag in latest ${VERSION}; do 
    docker tag exchange-ndr-handler:${tag} twrefapp/exchange-ndr-handler:${tag}
    docker push twrefapp/exchange-ndr-handler:${tag}
done
