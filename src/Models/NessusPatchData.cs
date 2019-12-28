// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;

namespace openrmf_read_api.Models
{
    /// <summary>
    /// This is the class that shows the summary number of items for the ACAS Nessus Patch scan of systems
    /// </summary>

    public class NessusPatchData
    {
        public NessusPatchData () {
            summary = new List<NessusPatchSummary>();
        }
        public string reportName { get; set; }
        public List<NessusPatchSummary> summary { get; set; }
    }
}