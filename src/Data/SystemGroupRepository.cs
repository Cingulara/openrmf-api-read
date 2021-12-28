// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using openrmf_read_api.Models;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;
using MongoDB.Driver;
using MongoDB.Bson;
using Microsoft.Extensions.Options;

namespace openrmf_read_api.Data {
    public class SystemGroupRepository : ISystemGroupRepository
    {
        private readonly ArtifactContext _context = null;

        public SystemGroupRepository(IOptions<Settings> settings)
        {
            _context = new ArtifactContext(settings);
        }

        public async Task<IEnumerable<SystemGroup>> GetAllSystemGroups()
        {
            try
            {
                return await _context.SystemGroups
                        .Find(_ => true).SortBy(x => x.title).ToListAsync();
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
        public async Task<SystemGroup> GetSystemGroup(string id)
        {
            try
            {
                ObjectId internalId = GetInternalId(id);
                return await _context.SystemGroups
                                .Find(SystemGroup => SystemGroup.InternalId == internalId).FirstOrDefaultAsync();
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }


        public async Task<long> CountSystems(){
            try {
                long result = await _context.SystemGroups.CountDocumentsAsync(Builders<SystemGroup>.Filter.Empty);
                return result;
            }
            catch (Exception ex)
            {
                // log or manage the exception
                throw ex;
            }
        }

        // check that the database is responding and it returns at least one collection name
        public bool HealthStatus(){
            var result = _context.SystemGroups.Database.ListCollectionNamesAsync().GetAwaiter().GetResult().FirstOrDefault();
            if (!string.IsNullOrEmpty(result)) // we are good to go
                return true;
            return false;
        }
    }
}