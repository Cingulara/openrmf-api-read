using System;
using System.Collections.Generic;


namespace openstig_read_api.Models
{

    public class ASSET {

        public ASSET (){

        }

		public string ROLE { get; set; }
		public string ASSET_TYPE { get; set; }
		public string HOST_NAME { get; set; }
		public string HOST_IP { get; set; }
		public string HOST_MAC { get; set; }
		public string HOST_FQDN { get; set; }
		public string TECH_AREA { get; set; }
		public string TARGET_KEY { get; set; }
		public string WEB_OR_DATABASE { get; set; }
		public string WEB_DB_SITE { get; set; }
		public string WEB_DB_INSTANCE { get; set; }
    }
}