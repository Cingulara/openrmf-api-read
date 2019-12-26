// Copyright (c) Cingulara 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using System.Xml.Serialization;

namespace openrmf_read_api.Models
{

    public class iSTIG {

        public iSTIG (){
            STIG_INFO = new STIG_INFO();
            VULN = new List<VULN>();
        }

        public STIG_INFO STIG_INFO { get; set; }

        [XmlElement("VULN")]
        public List<VULN> VULN { get; set; }
    }
}