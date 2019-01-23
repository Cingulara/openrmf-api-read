# openstig-api-save
This is the openSTIG Save API for scoring a checklist. It has two calls.

POST to /api/save/ to save a new document
PUT to /api/save/{id} to update a document

/swagger/ gives you the API structure.


## creating the user
* ~/mongodb/bin/mongo 'mongodb://root:myp2ssw0rd@192.168.11.46'
* use admin
* db.createUser({ user: "openstig" , pwd: "openstig1234!", roles: ["readWriteAnyDatabase"]});
* use openstig
* db.createCollection("Artifacts");

## connecting to the database collection straight
~/mongodb/bin/mongo 'mongodb://openstig:openstig1234!@192.168.11.46/openstig?authSource=admin'
