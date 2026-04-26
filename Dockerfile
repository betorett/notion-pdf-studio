FROM node:22-alpine

WORKDIR /app
ENV NODE_ENV=production

COPY package.json ./
COPY server.js ./
COPY public ./public
COPY .env.example ./

RUN mkdir -p data

EXPOSE 4173
CMD ["node", "server.js"]
