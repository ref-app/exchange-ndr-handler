#!/usr/bin/env bash
DIRECTORY=$(cd `dirname $0` && pwd)
VERSION=$(cat "${DIRECTORY}/../VERSION")

for image in exchange-ndr-handler inbox-notifier; do
    for tag in latest ${VERSION}; do 
        docker tag ${image}:${tag} twrefapp/${image}:${tag}
        docker push twrefapp/${image}:${tag}
    done
done
