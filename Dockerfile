FROM node:20-bullseye

# Cài Ruby để chạy gem chuyển MTEF → MathML
RUN apt-get update && apt-get install -y --no-install-recommends ruby-full \
    && rm -rf /var/lib/apt/lists/*

# Cài gem chuyển đổi MathType MTEF → MathML
RUN gem install mathtype_to_mathml

WORKDIR /app
COPY package.json package-lock.json* ./
RUN npm install

COPY mt2mml.rb server.js ./

EXPOSE 8080
ENV NODE_ENV=production
CMD ["npm", "start"]
