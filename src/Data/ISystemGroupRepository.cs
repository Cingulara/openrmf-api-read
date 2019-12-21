using openrmf_read_api.Models;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;

namespace openrmf_read_api.Data {
    public interface ISystemGroupRepository
    {
        Task<IEnumerable<SystemGroup>> GetAllSystemGroups();
        Task<SystemGroup> GetSystemGroup(string id);

        /******************************************** 
         Dashboard specific calls
        ********************************************/
        // get the # of checklists for the dashboard listing
        Task<long> CountSystems();
    }
}