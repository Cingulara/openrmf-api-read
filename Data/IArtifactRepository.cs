using openstig_read_api.Models;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;

namespace openstig_read_api.Data {
    public interface IArtifactRepository
    {
        Task<IEnumerable<Artifact>> GetAllArtifacts();
        Task<Artifact> GetArtifact(string id);

        // query after multiple parameters
        Task<IEnumerable<Artifact>> GetArtifact(string bodyText, DateTime updatedFrom, long headerSizeLimit);

        // add new note document
        Task AddArtifact(Artifact item);

        // remove a single document
        Task<bool> RemoveArtifact(string id);

        // update just a single document
        Task<bool> UpdateArtifact(string id, Artifact body);

        // should be used with high cautious, only in relation with demo setup
        Task<bool> RemoveAllArtifacts();

        /******************************************** 
         Dashboard specific calls
        */

        // get the # of checklists for the dashboard listing
        Task<long> CountChecklists();
        // get the last 5 checklists being edited in a sort order by last updated
        Task<IEnumerable<Artifact>> GetLatestArtifacts(int number);
    }
}