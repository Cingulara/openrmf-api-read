// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using openrmf_read_api.Models;
using System.Collections.Generic;
using System;
using System.Threading.Tasks;

namespace openrmf_read_api.Data
{
    public interface IArtifactRepository
    {
        Task<IEnumerable<Artifact>> GetAllArtifacts();
        Task<Artifact> GetArtifact(string id);

        // query after multiple parameters
        Task<IEnumerable<Artifact>> GetArtifact(string bodyText, DateTime updatedFrom, long headerSizeLimit);

        Task<Artifact> GetArtifactBySystem(string systemGroupId, string artifactId);
        
        // get all artifacts by checklist STIG type and version
        Task<IEnumerable<Artifact>> GetArtifactsByStigType(string systemGroupId, string stigType);

        /******************************************** 
         System specific calls
        ********************************************/

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

        // add new note document
        Task<Artifact> AddArtifact(Artifact item);

        // remove a single document
        Task<bool> RemoveArtifact(string id);

        // update just a single document
        Task<bool> UpdateArtifact(string id, Artifact body);

        // see if there is a checklist based on the system, hostname, and STIG checklist type
        Task<Artifact> GetArtifactBySystemHostnameAndType(string systemGroupId, string hostName, string stigType);

        Task<bool> DeleteArtifact(string id);
    }
}