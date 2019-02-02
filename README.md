# openstig-api-read
This is the openSTIG Read API for reading a checklist and all its metadata we store. It has two calls.

GET to / to list all records
GET to /{id} to get a record
GET to /download/{id} to download the CKL file to use in the STIG viewer

/swagger/ gives you the API structure.

## Making your local Docker image
docker build --rm -t openstig-api-read:0.1 .

## creating the user
* ~/mongodb/bin/mongo 'mongodb://root:myp2ssw0rd@localhost'
* use admin
* db.createUser({ user: "openstig" , pwd: "openstig1234!", roles: ["readWriteAnyDatabase"]});
* use openstig
* db.createCollection("Artifacts");

## connecting to the database collection straight
~/mongodb/bin/mongo 'mongodb://openstig:openstig1234!@localhost/openstig?authSource=admin'

## Messaging Platform
Using NATS from Synadia to have a messaging backbone and eventual consistency. Currently publishing to these known items:
* openstig.save.new with payload (new Guid Id)
* openstig.save.update with payload (new Guid Id)
* openstig.upload.new with payload (new Guid Id)
* openstig.upload.update with payload (new Guid Id)

More will follow as this expands for auditing, logging, etc.

### How to run NATS
* docker run --rm --name nats-main -p 4222:4222 -p 6222:6222 -p 8222:8222 nats
* this is the default and lets you run a NATS server version 1.2.0 (as of 8/2018)
* just runs in memory and no streaming (that is separate)