// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

namespace openrmf_read_api.Models
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