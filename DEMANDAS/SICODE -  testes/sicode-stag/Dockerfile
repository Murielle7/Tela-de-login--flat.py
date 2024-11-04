# Dockerfile development version

FROM ruby:3.1.0

# Set the Node.js version you want to install
ENV NODE_VERSION 18

# Install dependencies
RUN apt-get update -qq && apt-get install -y build-essential curl libpq-dev

RUN curl -fsSL https://deb.nodesource.com/setup_${NODE_VERSION}.x | bash -

RUN apt-get install -y nodejs

RUN npm install -g yarn

# Set the working directory
WORKDIR /app

# Copy the Gemfile and Gemfile.lock
COPY Gemfile /app/Gemfile
COPY Gemfile.lock /app/Gemfile.lock
COPY package.json /app/package.json
COPY yarn.lock /app/yarn.lock

RUN gem install bundler -v 2.5
RUN bundle install
# RUN yarn install --ignore-engines

COPY . ./app

# Entrypoint prepares the database.
#ENTRYPOINT ["/rails/bin/docker-entrypoint"]

# Start the server by default, this can be overwritten at runtime
EXPOSE 3000
# CMD ["./bin/rails", "server", '-b', '0.0.0.0']
