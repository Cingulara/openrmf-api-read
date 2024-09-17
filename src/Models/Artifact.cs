// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class Artifact
    {
        public Artifact () {
            CHECKLIST = new CHECKLIST();
            isWebDatabase = false;
            webDatabaseSite = "";
            webDatabaseInstance = "";
        }

        public DateTime created { get; set; }
        public CHECKLIST CHECKLIST { get; set; }
        public string rawChecklist { get; set; }
        
        // if this is part of a system, list that system.
        // if empty this is just a standalone checklist
        public string systemGroupId { get; set; }
        public string systemTitle { get; set; }
        public string hostName { get; set;}
        public string stigType { get; set; }
        public string stigRelease { get; set; }
        public string version {get; set;}
        public string title { get {
            string finalTitle = "";
            string validHostname = !string.IsNullOrEmpty(hostName)? hostName.Trim() : "Unknown";
            string validStigType = "";
            if (!string.IsNullOrWhiteSpace(stigType)) validStigType = stigType.Trim();
            string validStigRelease = "";
            if (!string.IsNullOrWhiteSpace(stigRelease)) validStigRelease = stigRelease.Trim();

            finalTitle = validHostname + "-" + validStigType + "-V" + version + "-" + validStigRelease;
            // v1.12 added web or database uniqueness
            if (isWebDatabase) { // must have one of the others filled out to show extra information
                if (!string.IsNullOrWhiteSpace(webDatabaseSite) && 
                    !string.IsNullOrWhiteSpace(webDatabaseInstance)) {
                    finalTitle += " (" + webDatabaseSite + ", " + webDatabaseInstance + ")";
                } else if (!string.IsNullOrWhiteSpace(webDatabaseSite)) {
                    finalTitle += " (" + webDatabaseSite + ")";
                } else if (!string.IsNullOrWhiteSpace(webDatabaseInstance)) {
                    finalTitle += " (" + webDatabaseInstance + ")";
                }
            }

            return finalTitle;
        }}

        public string typeFullTitle { get {
                string finalTitle = stigType;
                if (isWebDatabase) { // must have one of the others filled out to show extra information
                    if (!string.IsNullOrWhiteSpace(webDatabaseSite) && 
                        !string.IsNullOrWhiteSpace(webDatabaseInstance)) {
                        finalTitle += " (" + webDatabaseSite + ", " + webDatabaseInstance + ")";
                    } else if (!string.IsNullOrWhiteSpace(webDatabaseSite)) {
                        finalTitle += " (" + webDatabaseSite + ")";
                    } else if (!string.IsNullOrWhiteSpace(webDatabaseInstance)) {
                        finalTitle += " (" + webDatabaseInstance + ")";
                    }
                }
                return finalTitle;
            }
        }
        
        [BsonId]
        // standard BSonId generated by MongoDb
        public ObjectId InternalId { get; set; }
        public string InternalIdString { get { return InternalId.ToString();}}

        [BsonDateTimeOptions]
        // attribute to gain control on datetime serialization
        public DateTime? updatedOn { get; set; }

        public Guid createdBy { get; set; }
        public Guid? updatedBy { get; set; }

        // v1.7
        public List<string> tags {get; set;}

        // v1.12
        public bool isWebDatabase { get; set; }
        public string webDatabaseSite { get; set; }
        public string webDatabaseInstance { get; set; }
    }
}