using System;
using System.Collections.Generic;
using MongoDB.Bson.Serialization.Attributes;
using MongoDB.Bson;

namespace openstig_read_api.Models
{
    [Serializable]
    public class ArtifactCount
    {
        public ArtifactCount () {
        }
        public string stigType { get; set; }

        public int count { get; set; }
    }
}