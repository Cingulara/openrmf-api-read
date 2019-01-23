using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;

namespace openstig_read_api.Models
{
    [Serializable]
    public class Artifact
    {
        public Artifact () {
            id= Guid.NewGuid();
            Checklist = new CHECKLIST();
        }

        public DateTime created { get; set; }
        public string title { get; set; }
        public CHECKLIST Checklist { get; set; }
        public string rawChecklist { get; set; }
        [BsonId]
        public Guid id { get; set; }
        public STIGtype type { get; set; }

        [BsonDateTimeOptions]
        // attribute to gain control on datetime serialization
        public DateTime UpdatedOn { get; set; } = DateTime.Now;
    }

    public enum STIGtype{
        ASD = 0,
        DBInstance = 10,
        DBServer = 20,
        DOTNET = 30,
        JAVA = 40
    }

}