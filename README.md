# openstig-api-save
This is the openSTIG Read API for reading a checklist and all its metadata we store. It has two calls.

GET to / to list all records
GET to /{id} to get a record
GET to /download/{id} to download the CKL file to use in the STIG viewer

/swagger/ gives you the API structure.


## creating the user
* ~/mongodb/bin/mongo 'mongodb://root:myp2ssw0rd@localhost'
* use admin
* db.createUser({ user: "openstig" , pwd: "openstig1234!", roles: ["readWriteAnyDatabase"]});
* use openstig
* db.createCollection("Artifacts");

## connecting to the database collection straight
~/mongodb/bin/mongo 'mongodb://openstig:openstig1234!@localhost/openstig?authSource=admin'
