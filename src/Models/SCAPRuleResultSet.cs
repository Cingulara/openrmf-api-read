// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.
using System;
using System.Collections.Generic;

namespace openrmf_read_api.Models
{
    // there is a cdf:TestResult area under the cdf:Benchmark tag
    // read in each cdf:rule_result under the TestResult area
    // there is the *idref* field that matches to the rule Id field from the VULN in each checklist (i.e. SV-78007r1_rule)
    // the *cdf:result* will have pass or fail for that rule
    // save all that rule result data into a list to use for that VULN based on the rule idref field
    // use .Replace(xxx,"") to just get the SV-xxx rule information
    [Serializable]
    public class SCAPRuleResultSet
    {
        public SCAPRuleResultSet () {
            ruleResults = new List<SCAPRuleResult>();
            ipaddress = "";
            fqdn = "";
            macaddress = "";
        }

        public string title { get; set; }
        public string hostname { get; set; }
        public string ipaddress { get; set;}
        public string fqdn { get; set; }
        public string macaddress { get; set; }
        public string scanTool { get; set; }
        public string scanTime { get; set;}
        public List<SCAPRuleResult> ruleResults { get; set; }
    }
}