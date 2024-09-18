// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using openrmf_read_api.Models;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using System.Linq;
using MongoDB.Driver;
using MongoDB.Bson;
using MongoDB.Driver.Linq;
using Microsoft.Extensions.Options;

namespace openrmf_read_api.Data
{
    public class ArtifactRepository : IArtifactRepository
    {
        private readonly ArtifactContext _context = null;

        public ArtifactRepository(IOptions<Settings> settings)
        {
            _context = new ArtifactContext(settings);
        }

        public async Task<IEnumerable<Artifact>> GetAllArtifacts()
        {
            return await _context.Artifacts
                    .Find(_ => true).ToListAsync();
        }


        private ObjectId GetInternalId(string id)
        {
            ObjectId internalId;
            if (!ObjectId.TryParse(id, out internalId))
                internalId = ObjectId.Empty;

            return internalId;
        }

        // query after Id or InternalId (BSonId value)
        //
        public async Task<Artifact> GetArtifact(string id)
        {
            return await _context.Artifacts.Find(artifact => artifact.InternalId == GetInternalId(id)).FirstOrDefaultAsync();
        }

        // query after body text, updated time, and header image size
        //
        public async Task<IEnumerable<Artifact>> GetArtifact(string bodyText, DateTime updatedFrom, long headerSizeLimit)
        {
            var query = _context.Artifacts.Find(artifact => artifact.title.Contains(bodyText) &&
                                artifact.updatedOn >= updatedFrom);

            return await query.ToListAsync();
        }

        // query after Id or InternalId (BSonId value) by checking artifactId and systemGroupId
        public async Task<Artifact> GetArtifactBySystem(string systemGroupId, string artifactId)
        {
                ObjectId internalId = GetInternalId(artifactId);
                return await _context.Artifacts
                    .Find(artifact => artifact.InternalId == internalId && artifact.systemGroupId == systemGroupId).FirstOrDefaultAsync();
        }
        
        // query on the artifact stigType and version
        public async Task<IEnumerable<Artifact>> GetArtifactsByStigType(string systemGroupId, string stigType)
        {
                var query = _context.Artifacts.Find(artifact => artifact.stigType == stigType && 
                            artifact.systemGroupId == systemGroupId);
                return await query.ToListAsync();
        }

        public async Task<long> CountChecklists()
        {
            long result = await _context.Artifacts.CountDocumentsAsync(Builders<Artifact>.Filter.Empty);
            return result;
        }

        public async Task<IEnumerable<Artifact>> GetLatestArtifacts(int number)
        {
            return await _context.Artifacts.Find(_ => true).SortByDescending(y => y.updatedOn).Limit(number).ToListAsync();
        }

        public async Task<IEnumerable<object>> GetCountByType(string systemGroupId)
        {
            // show them all by type
            if (string.IsNullOrEmpty(systemGroupId))
            {
                var groupArtifactItemsByType = _context.Artifacts.Aggregate()
                        .Group(s => s.stigType,
                        g => new ArtifactCount { stigType = g.Key, count = g.Count() }).ToListAsync();
                return await groupArtifactItemsByType;
            }
            else
            {
                var groupArtifactItemsByType = _context.Artifacts.Aggregate().Match(artifact => artifact.systemGroupId == systemGroupId)
                        .Group(s => s.stigType,
                        g => new ArtifactCount { stigType = g.Key, count = g.Count() }).ToListAsync();
                return await groupArtifactItemsByType;
            }
        }

        public async Task<Artifact> GetArtifactBySystemHostnameAndType(string systemGroupId, string hostName, string stigType)
        {
            var query = _context.Artifacts.Find(artifact => artifact.systemGroupId == systemGroupId &&
                        !string.IsNullOrWhiteSpace(artifact.hostName) && !string.IsNullOrWhiteSpace(hostName) &&
                        artifact.hostName.ToLower() == hostName.ToLower() && artifact.stigType == stigType);

            return await query.FirstOrDefaultAsync();
        }

        public async Task<Artifact> GetArtifactBySystemHostnameAndTypeWithWebDatabase(string systemGroupId, string hostName, string stigType, 
            bool isWebDatabase, string webDatabaseSite, string webDatabaseInstance)
        {
            var query = _context.Artifacts.Find(artifact => artifact.systemGroupId == systemGroupId &&
                        !string.IsNullOrWhiteSpace(artifact.hostName) && !string.IsNullOrWhiteSpace(hostName) &&
                        artifact.hostName.ToLower() == hostName.ToLower() && artifact.stigType == stigType &&
                        artifact.isWebDatabase == isWebDatabase && artifact.webDatabaseSite == webDatabaseSite && 
                        artifact.webDatabaseInstance == webDatabaseInstance);

            return await query.FirstOrDefaultAsync();
        }

        public async Task<Artifact> AddArtifact(Artifact item)
        {
            await _context.Artifacts.InsertOneAsync(item);
            return item;
        }

        public async Task<bool> RemoveArtifact(string id)
        {
            DeleteResult actionResult
                = await _context.Artifacts.DeleteOneAsync(
                    Builders<Artifact>.Filter.Eq("Id", id));

            return actionResult.IsAcknowledged
                && actionResult.DeletedCount > 0;
        }

        public async Task<bool> UpdateArtifact(string id, Artifact body)
        {
            var filter = Builders<Artifact>.Filter.Eq(s => s.InternalId, GetInternalId(id));
            body.InternalId = GetInternalId(id);
            var actionResult = await _context.Artifacts.ReplaceOneAsync(filter, body);
            return actionResult.IsAcknowledged && actionResult.ModifiedCount > 0;
        }

        public async Task<bool> DeleteArtifact(string id)
        {
            var filter = Builders<Artifact>.Filter.Eq(s => s.InternalId, GetInternalId(id));
                Artifact art = new Artifact();
                art.InternalId = GetInternalId(id);
                // only save the data outside of the checklist, update the date
                var currentRecord = await _context.Artifacts.Find(artifact => artifact.InternalId == art.InternalId).FirstOrDefaultAsync();
                if (currentRecord != null){
                    DeleteResult actionResult = await _context.Artifacts.DeleteOneAsync(Builders<Artifact>.Filter.Eq("_id", art.InternalId));
                    return actionResult.IsAcknowledged && actionResult.DeletedCount > 0;
                } 
                else {
                    throw new KeyNotFoundException();
                }
        }

        #region Systems

        public async Task<IEnumerable<Artifact>> GetSystemArtifacts(string systemGroupId)
        {
            var query = await _context.Artifacts.FindAsync(artifact => artifact.systemGroupId == systemGroupId);
            return query.ToList().OrderBy(x => x.title);
        }
        #endregion
    }
}