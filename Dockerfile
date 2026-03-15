FROM node:20-slim

RUN apt-get update && apt-get install -y --no-install-recommends \
    python3 python3-pip python3-dev build-essential \
        libfreetype6-dev libpng-dev pkg-config \
            && rm -rf /var/lib/apt/lists/*

            RUN pip3 install --no-cache-dir --break-system-packages \
                numpy==1.26.4 matplotlib==3.8.4 scipy==1.13.1

                WORKDIR /app
                COPY package.json ./
                RUN npm install --omit=dev
                COPY . .
                RUN mkdir -p /tmp/ama-jobs

                EXPOSE 3000
                ENV PORT=3000
                ENV NODE_ENV=production
                CMD ["node", "server.js"]
