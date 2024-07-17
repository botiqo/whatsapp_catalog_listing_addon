FROM node:14

WORKDIR /app

COPY package*.json ./

RUN npm install

COPY . .

RUN npm install -g @google/clasp

CMD ["clasp", "push"]