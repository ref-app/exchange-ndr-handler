FROM node:20-alpine3.18

LABEL Maintainer "Refapp - https://github.com/ref-app"

WORKDIR /usr/src

COPY .yarn/ ./.yarn/
COPY .yarnrc.yml package.json yarn.lock README.md ./

RUN yarn

# So we can override it when building the image
ARG SCRIPT=process-ndr-messages.ts
ENV SCRIPT $SCRIPT

CMD ["sh","-c","./${SCRIPT}"]

# These files is most likely to change often so put it last in the Dockerfile for caching reasons
COPY *.ts ./
