// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace openrmf_read_api.Models.Compliance
{
    [Serializable]
    public class NISTCompliance
    {
        public NISTCompliance () {
            complianceRecords = new List<ComplianceRecord>();
        }

        // the control is the major piece to use here, AC-1, AU-9, etc.
        public string control { get; set;}

        // the index is the major control with all extra dots, dashes, and sub paragraphs
        public string index { get; set; }
        // This is the title of the index from the NIST site (tbd)
        public string sortString { get; set; } // sort by this to get the listing correct

        public string title { get; set; }
        public string version { get; set; }
        public string location { get; set; }
        public string CCI { get; set; }
        public List<ComplianceRecord> complianceRecords { get; set; }
    }

}