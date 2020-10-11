FROM node:14.13.1

LABEL Maintainer "Refapp - https://github.com/ref-app"

RUN yarn global add ts-node typescript

WORKDIR /usr/src

COPY package.json yarn.lock README.md ./

RUN yarn

COPY *.ts ./

ENTRYPOINT [ "/usr/local/bin/ts-node" ]

CMD ["process-ndr-messages.ts"]