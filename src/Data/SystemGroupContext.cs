using MongoDB.Driver;
using openrmf_read_api.Models;
using Microsoft.Extensions.Options;

namespace openrmf_read_api.Data
{
    public class SystemGroupContext
    {
        private readonly IMongoDatabase _database = null;

        public SystemGroupContext(IOptions<Settings> settings)
        {
            var client = new MongoClient(settings.Value.ConnectionString);
            if (client != null)
                _database = client.GetDatabase(settings.Value.Database);
        }

        public IMongoCollection<SystemGroup> SystemGroups
        {
            get
            {
                return _database.GetCollection<SystemGroup>("SystemGroups");
            }
        }
    }
}