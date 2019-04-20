using openrmf_read_api.Models;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;

namespace openrmf_read_api.Data {
    public interface IArtifactRepository
    {
        Task<IEnumerable<Artifact>> GetAllArtifacts();
        Task<Artifact> GetArtifact(string id);

        // query after multiple parameters
        Task<IEnumerable<Artifact>> GetArtifact(string bodyText, DateTime updatedFrom, long headerSizeLimit);

        /******************************************** 
         System specific calls
        ********************************************/
        // return a distinct list of Systems
        Task<List<string>> GetAllSystems();
        // return checklist records for a given system
        Task<IEnumerable<Artifact>> GetSystemArtifacts(string system);

        /******************************************** 
         Dashboard specific calls
        ********************************************/
        // get the # of checklists for the dashboard listing
        Task<long> CountChecklists();
        // get the last 5 checklists being edited in a sort order by last updated
        Task<IEnumerable<Artifact>> GetLatestArtifacts(int number);

        /******************************************** 
         Reports specific calls
        ********************************************/
        Task<IEnumerable<object>> GetCountByType(string system);
    }
}