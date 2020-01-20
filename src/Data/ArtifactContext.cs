// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using MongoDB.Driver;
using openrmf_read_api.Models;
using Microsoft.Extensions.Options;

namespace openrmf_read_api.Data
{
    public class ArtifactContext
    {
        private readonly IMongoDatabase _database = null;

        public ArtifactContext(IOptions<Settings> settings)
        {
            var client = new MongoClient(settings.Value.ConnectionString);
            if (client != null)
                _database = client.GetDatabase(settings.Value.Database);
        }

        public IMongoCollection<Artifact> Artifacts
        {
            get
            {
                return _database.GetCollection<Artifact>("Artifacts");
            }
        }
    }
}