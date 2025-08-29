FROM node:20-bullseye

# Ruby + toolchain
RUN apt-get update && apt-get install -y --no-install-recommends \
    ruby-full build-essential \
  && rm -rf /var/lib/apt/lists/*

# Gem: cần cho converter
RUN gem install --no-document pry mathtype_to_mathml

WORKDIR /app

# Cài npm deps: dùng ci nếu có lockfile, ngược lại dùng install
COPY package.json package-lock.json* ./
RUN if [ -f package-lock.json ]; then \
      npm ci --omit=dev; \
    else \
      npm install --omit=dev; \
    fi

# Source
COPY mt2mml.rb server.js ./

EXPOSE 8080
ENV NODE_ENV=production
CMD ["npm", "start"]
