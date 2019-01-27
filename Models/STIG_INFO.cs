using System;
using System.Collections.Generic;


namespace openstig_read_api.Models
{

    public class STIG_INFO {

        public STIG_INFO (){
            SI_DATAs = new List<SI_DATA>();
        }

        public List<SI_DATA> SI_DATAs { get; set;}
    }
}