using System;
using System.Collections.Generic;


namespace openstig_read_api.Models
{

    public class VULN {

        public VULN (){
            STIG_DATAs = new List<STIG_DATA>();
        }

        public List<STIG_DATA> STIG_DATAs { get; set;}
		public string STATUS { get; set;}
		public string FINDING_DETAILS { get; set;}
		public string COMMENTS { get; set;}
		public string SEVERITY_OVERRIDE { get; set;}
		public string SEVERITY_JUSTIFICATION { get; set;}
    }
}