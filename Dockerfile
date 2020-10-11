FROM node:14.13.1

RUN yarn global add ts-node typescript

COPY package.json yarn.lock README.md SRC/

RUN cd SRC && yarn

COPY *.ts SRC/
