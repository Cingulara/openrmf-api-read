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
        private readonly SystemGroupContext _context = null;

        public SystemGroupRepository(IOptions<Settings> settings)
        {
            _context = new SystemGroupContext(settings);
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

    }
}