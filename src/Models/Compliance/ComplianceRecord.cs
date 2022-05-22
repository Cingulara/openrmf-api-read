// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;

namespace openrmf_read_api.Models.Compliance
{
    [Serializable]
    public class ComplianceRecord
    {
        public ComplianceRecord () {
            // any initialization here
        }
        public string artifactId { get; set; }
        public string title { get; set; }
        public string stigType { get; set; }
        public string stigRelease { get; set; }
        
        public string status { get; set; }
        public string hostName { get; set;}

        // the last time this was updated
        public DateTime updatedOn { get; set; }
    }

}