FROM node:14.13.1

LABEL Maintainer "Refapp - https://github.com/ref-app"

RUN yarn global add ts-node typescript

WORKDIR /usr/src

COPY package.json yarn.lock README.md ./

RUN yarn

ENTRYPOINT [ "/usr/local/bin/ts-node" ]

CMD ["process-ndr-messages.ts"]

# These files is most likely to change often so put it last in the Dockerfile for caching reasons
COPY *.ts ./
