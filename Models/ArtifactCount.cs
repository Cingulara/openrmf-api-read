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
        public STIGtype type { get; set; }
        public string typeTitle { get { return Enum.GetName(typeof(STIGtype), type);} }

        public int count { get; set; }
    }
}