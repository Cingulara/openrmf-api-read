// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

namespace openrmf_read_api.Models
{
    /// <summary>
    /// This is the class that shows the summary number of items for the ACAS Nessus Patch scan of systems
    /// </summary>

    public class NessusPatchSummary
    {
        public NessusPatchSummary () {        }
        public string hostname { get; set;}
        public string operatingSystem { get; set;}
        public string systemType { get; set;}
        public string ipAddress { get; set;}
        public bool credentialed { get; set;}
        
        public string pluginId { get; set; }

        public string pluginIdSort { get {
            if (pluginId.Length >= 6)
                return pluginId;
            else 
                return "0" + pluginId;
        }}
        public string pluginName { get; set; }
        public string family { get; set; }
        public int severity { get; set; }
        public string severityName { get {
            if (severity == 4)
                return "Critical";
            else if (severity == 3)
                return "High";
            else if (severity == 2)
                return "Medium";
            else if (severity == 1)
                return "Low";
            else
                return "Informational";
        }}
        // how many hosts have this pluginId
        public int hostTotal { get; set; }
        // how many times has this pluginId come up in total
        public int total { get; set; }

        // specific data points for the report are below
        public string description { get; set; }
        public string publicationDate { get; set; }
        public string pluginType { get; set; }
        public string riskFactor { get; set; }
        public string synopsis { get; set; }

        // Nessus ACAS Scanner Version
        public string scanVersion { get; set; }
    }
}