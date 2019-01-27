using System;
using System.Collections.Generic;


namespace openstig_read_api.Models
{

    public class iSTIG {

        public iSTIG (){
            STIG_INFO = new STIG_INFO();
            VULNs = new List<VULN>();
        }

        public STIG_INFO STIG_INFO { get; set; }
        public List<VULN> VULNs { get; set; }
    }
}