FROM node:20-bullseye

# Ruby + toolchain để cài gem (nhanh gọn)
RUN apt-get update && apt-get install -y --no-install-recommends \
    ruby-full build-essential \
  && rm -rf /var/lib/apt/lists/*

# 👇 Thêm pry để gem mathtype_to_mathml require được
RUN gem install --no-document pry mathtype_to_mathml

WORKDIR /app
COPY package.json package-lock.json* ./
# dùng npm ci nếu có lockfile, fallback npm install
RUN npm ci || npm install

COPY mt2mml.rb server.js ./

EXPOSE 8080
ENV NODE_ENV=production
CMD ["npm", "start"]
