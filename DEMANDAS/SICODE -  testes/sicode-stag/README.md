# README

This README would normally document whatever steps are necessary to get the
application up and running.

Things you may want to cover:

* Ruby version
  3.1.0

* System dependencies
  - `docker`
  - `docker-compose`

* Configuration

* Database creation
  - `docker-compose run web rails db:create`

* Database initialization
  - `docker-compose run web rails db:migrate`

* Database seed
  - `docker-compose run web rails db:seed`

* How to run the test suite
  - `docker-compose run web bundle exec rspec`

* Assets compilation
  - `docker-compose run web yarn install`
  - `docker-compose run web rails yarn build`
  - `docker-compose run web rails assets:precompile`

* Services (job queues, cache servers, search engines, etc.)
  - `docker-compose up`

* Deployment instructions
  - `docker-compose run web rails db:migrate`
  - `docker-compose run web rails db:seed`
  - `docker-compose up`


* ...

Please feel free to use a different markup language if you do not plan to run
`rake doc:app`.
