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

namespace openrmf_read_api.Data {
    public class ArtifactRepository : IArtifactRepository
    {
        private readonly ArtifactContext _context = null;

        public ArtifactRepository(IOptions<Settings> settings)
        {
            _context = new ArtifactContext(settings);
        }

        public async Task<IEnumerable<Artifact>> GetAllArtifacts()
        {
            try
            {
                return await _context.Artifacts
                        .Find(_ => true).ToListAsync();
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
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
            try
            {
                return await _context.Artifacts.Find(artifact => artifact.InternalId == GetInternalId(id)).FirstOrDefaultAsync();
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }

        // query after body text, updated time, and header image size
        //
        public async Task<IEnumerable<Artifact>> GetArtifact(string bodyText, DateTime updatedFrom, long headerSizeLimit)
        {
            try
            {
                var query = _context.Artifacts.Find(artifact => artifact.title.Contains(bodyText) &&
                                    artifact.updatedOn >= updatedFrom);

                return await query.ToListAsync();
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }
        
        public async Task<long> CountChecklists(){
            try {
                long result = await _context.Artifacts.CountDocumentsAsync(Builders<Artifact>.Filter.Empty);
                return result;
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }

        public async Task<IEnumerable<Artifact>> GetLatestArtifacts(int number)
        {
            try
            {
                return await _context.Artifacts.Find(_ => true).SortByDescending(y => y.updatedOn).Limit(number).ToListAsync();
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }

        public async Task<IEnumerable<object>> GetCountByType(string systemGroupId)
        {
            try
            {
                // show them all by type
                if (string.IsNullOrEmpty(systemGroupId)) {
                    var groupArtifactItemsByType = _context.Artifacts.Aggregate()
                            .Group(s => s.stigType,
                            g => new ArtifactCount {stigType = g.Key, count = g.Count()}).ToListAsync();
                    return await groupArtifactItemsByType;
                }
                else {
                    var groupArtifactItemsByType = _context.Artifacts.Aggregate().Match(artifact => artifact.systemGroupId == systemGroupId)
                            .Group(s => s.stigType,
                            g => new ArtifactCount {stigType = g.Key, count = g.Count()}).ToListAsync();
                    return await groupArtifactItemsByType;
                }
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }
    
    
        #region Systems

        public async Task<IEnumerable<Artifact>> GetSystemArtifacts(string systemGroupId)
        {
            try
            {
                var query = await _context.Artifacts.FindAsync(artifact => artifact.systemGroupId == systemGroupId);
                return query.ToList().OrderBy(x => x.title);
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }
        #endregion
    }
}