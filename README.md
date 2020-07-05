![.NET Core Build and Test](https://github.com/Cingulara/openrmf-api-read/workflows/.NET%20Core%20Build%20and%20Test/badge.svg)

# openrmf-api-read
This is the openRMF Read API for reading a checklist and all its metadata we store. It has two calls.

* GET to / to list all records
* GET to /artifact/{id} to get a checklist record
* GET to /download/{id} to download a checklist record CKL file to use in the STIG viewer
* GET to /export/{id} to export a checklist to XLSX format color coded
* GET to /{id}/control/{control} to get all vulnerability IDs for a particular control type for this checklist (used for compliance generation view)
* GET to /system/export/{systemGroupId} to get a list of all checklists for a system in XLSX
* GET to /systems to get a list of all systems
* GET to /systems/{systemGroupId} to get all checklists for a system
* GET to /system/{systemGroupId} to get the system record
* GET to /system/{systemGroupId}/downloadnessus to download any NESSUS ACAS results file
* GET to /system/{systemGroupId}/nessuspatchsummary to download a Nessus ACAS Patch Summary XLSX
* GET to /system/{systemGroupId}/exportnessus to export the NESSUS ACAS results file
* GET to /system/{systemGroupId}/testplanexport to create a test plan in XLSX format
* GET to /system/{systemGroupId}/poamexport to export a POAM
* GET to /system/{systemGroupId}/rarexport to export a Risk Assessment Report
* GET to /count/artifacts to get a count of all checklists across all systems total
* GET to /count/systems to get a count of all systems
* /swagger/ gives you the API structure.

## Making your local Docker image
* make build
* make latest

## creating the user
* ~/mongodb/bin/mongo 'mongodb://root:myp2ssw0rd@localhost'
* use admin
* db.createUser({ user: "openrmf" , pwd: "openrmf1234!", roles: ["readWriteAnyDatabase"]});
* use openrmf
* db.createCollection("Artifacts");

## connecting to the database collection straight
~/mongodb/bin/mongo 'mongodb://openrmf:openrmf1234!@localhost/openrmf?authSource=admin'

## Messaging Platform
Using NATS from Synadia to have a messaging backbone and eventual consistency. Currently publishing to these known items:
* openrmf.score.read to get a score for a checklist
* openrmf.compliance.cci.control return a list of vulnerability IDs based on the control
* openrmf.scores.system get scores for all checklists in a system
* openrmf.compliance.cci get a list of all CCI items

More will follow as this expands for auditing, logging, etc.

### How to run NATS
* docker run --rm --name nats-main -p 4222:4222 -p 6222:6222 -p 8222:8222 nats
* this is the default and lets you run a NATS server version 1.2.0 (as of 8/2018)
* just runs in memory and no streaming (that is separate)

## Using Jaeger

The Jaeger Client is https://github.com/jaegertracing/jaeger-client-csharp. We use defaults but you can specify ENV for configuration.

## .NET Core 3.1
Need to at a minimum use the `mcr.microsoft.com/dotnet/core/aspnet:3.1` image to run this.