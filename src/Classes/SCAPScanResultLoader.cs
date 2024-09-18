// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.
using System.Collections.Generic;
using System.Linq;
using openrmf_read_api.Classes;
using System.IO;
using System.Xml;

namespace openrmf_read_api.Models
{
    // there is a cdf:TestResult area under the cdf:Benchmark tag
    // read in each cdf:rule_result under the TestResult area
    // there is the *idref* field that matches to the rule Id field from the VULN in each checklist (i.e. SV-78007r1_rule)
    // the *cdf:result* will have pass or fail for that rule
    // save all that rule result data into a list to use for that VULN based on the rule idref field
    // use .Replace(xxx,"") to just get the SV-xxx rule information

    public static class SCAPScanResultLoader
    {
        public static SCAPRuleResultSet LoadSCAPScan(string xmlfile) {
            SCAPRuleResultSet results = new SCAPRuleResultSet();
            // get the title of the SCAP scan we are using, which correlates to the Checklist
            // if a Nessus SCAP it uses "xccdf" tags
            xmlfile = xmlfile.Replace("\t","");
            string searchTag = "";
            // see if this is a DISA SCAP
            if (xmlfile.IndexOf("</cdf:") > 0)
                searchTag = "cdf:";
            // see if this is a Nessus SCAP
            if (xmlfile.IndexOf("</xccdf:") > 0)
                searchTag = "xccdf:";

            // now process the document
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(xmlfile);

            // get the template title from the SCAP to use to grab an empty Checklist
            XmlNodeList title = xmlDoc.GetElementsByTagName(searchTag + "title");
            if (title != null && title.Count > 0 && title.Item(0).FirstChild != null) {
                // get the title of the STIG so we can ask for the checklist later to fill in
                results.title = title.Item(0).FirstChild.InnerText;
            } else {
                // if not a DoD SCAP this is a Nessus SCAP (or trash)
                title = xmlDoc.GetElementsByTagName("xccdf:benchmark");
                if (title != null && title.Count > 0) {
                    // get the title of the STIG so we can ask for the checklist later to fill in
                    foreach (XmlNode node in title) {
                        if (node.Attributes.Count > 1 ) {
                            foreach (XmlAttribute attr in node.Attributes) {
                                if (attr.Name == "href" && !string.IsNullOrEmpty(attr.Value)) {
                                    // grab the Attribute's value
                                    if (!string.IsNullOrEmpty(attr.Value)) {
                                        results.title = attr.Value.Substring(0, attr.Value.IndexOf("_STIG_SCAP"));
                                        break; // we found it
                                    }
                                }
                            }
                        }
                    }
                }
            }
            if (string.IsNullOrEmpty(results.title))
                return results; // just return empty as we cannot match

            // get the target-address
            XmlNodeList targetAddresses = xmlDoc.GetElementsByTagName(searchTag + "target-address");
            if (targetAddresses != null && targetAddresses.Count > 0) {
                foreach (XmlNode node in targetAddresses) {
                    if (!string.IsNullOrEmpty(node.InnerText)) {
                        // grab the Node's InnerText
                        if (!string.IsNullOrEmpty(node.InnerText) && node.InnerText != "127.0.0.1" && 
                                    results.ipaddress != node.InnerText && results.ipaddress.IndexOf(", " + node.InnerText) < 0 && 
                                    results.ipaddress.IndexOf(node.InnerText + ", ") < 0) {
                            if (!string.IsNullOrEmpty(results.ipaddress))
                                results.ipaddress += ", " + node.InnerText;
                            else 
                                results.ipaddress = node.InnerText;
                        }
                    }
                }
            }

            // get the target (machine name) off the computer that was SCAP scanned
            XmlNodeList targetFacts = xmlDoc.GetElementsByTagName(searchTag + "target");
            if (targetFacts != null && targetFacts.Count > 0) {
                foreach (XmlNode node in targetFacts) {
                    if (!string.IsNullOrEmpty(node.InnerText)) {
                        // grab the Node's InnerText
                        results.hostname = node.InnerText;
                        break; // we found it
                    }
                }
            }

            // get the hostname and other facts off the computer that was SCAP scanned
            XmlNodeList targetFactsHostname = xmlDoc.GetElementsByTagName(searchTag + "fact");
            if (targetFactsHostname != null && targetFactsHostname.Count > 0) {
                foreach (XmlNode node in targetFactsHostname) {
                    if (node.Attributes.Count > 1 && (node.Attributes[0].InnerText.EndsWith("host_name") || 
                        node.Attributes[1].InnerText.EndsWith("host_name")) && 
                        string.IsNullOrEmpty(results.hostname))  {
                        // grab the Node's InnerText
                        results.hostname = node.InnerText;
                        //break;
                    } else if (node.Attributes.Count > 1 && (node.Attributes[0].InnerText.EndsWith("fqdn") || 
                        node.Attributes[1].InnerText.EndsWith("fqdn")) && 
                        string.IsNullOrEmpty(results.fqdn)) {
                        // grab the Node's InnerText
                        results.fqdn = node.InnerText;
                    } else if (node.Attributes.Count > 1 && (node.Attributes[0].InnerText.EndsWith("mac") || 
                        node.Attributes[1].InnerText.EndsWith("mac"))) {
                        // grab the Node's InnerText
                        if (!string.IsNullOrEmpty(node.InnerText) && node.InnerText != "00:00:00:00:00:00" && 
                                    results.macaddress != node.InnerText && results.macaddress.IndexOf(", " + node.InnerText) < 0 && 
                                    results.macaddress.IndexOf(node.InnerText + ", ") < 0) {
                            if (!string.IsNullOrEmpty(results.macaddress))
                                results.macaddress += ", " + node.InnerText;
                            else 
                                results.macaddress = node.InnerText;
                        }
                    } else if (node.Attributes.Count > 1 && (node.Attributes[0].InnerText.EndsWith("ipv4") || 
                        node.Attributes[1].InnerText.EndsWith("ipv4"))) {
                        // grab the Node's InnerText
                        if (!string.IsNullOrEmpty(node.InnerText) && node.InnerText != "127.0.0.1" && 
                                    !string.IsNullOrEmpty(results.ipaddress) && 
                                    results.ipaddress != node.InnerText && results.ipaddress.IndexOf(", " + node.InnerText) < 0 && 
                                    results.ipaddress.IndexOf(node.InnerText + ", ") < 0) {
                            if (!string.IsNullOrEmpty(results.ipaddress))
                                results.ipaddress += ", " + node.InnerText;
                            else 
                                results.ipaddress = node.InnerText;
                        }
                    }
                }
            }

            // GET the TestResult XML and grab test-system attribute and the end-time attribute
            // put into the format below IF you find it:
            // Tool: cpe:/a:spawar:scc:5.0.1
            // Time: 2019-04-19T17:13:08
            // Result: pass

            // get the CPE tool information and scan time
            XmlNodeList toolInformation = xmlDoc.GetElementsByTagName(searchTag + "TestResult");
            if (toolInformation != null && toolInformation.Count > 0 && toolInformation.Item(0).FirstChild != null) {
                foreach (XmlNode node in toolInformation) {
                    foreach (XmlAttribute attr in node.Attributes) {
                        if (attr.Name == "end-time") {
                            results.scanTime = attr.InnerText;
                        }
                        if (attr.Name == "test-system") {
                            results.scanTool = attr.InnerText;
                        }
                    }
                }
            } 

            // get all the rules and their pass/fail results
            XmlNodeList ruleResults = xmlDoc.GetElementsByTagName(searchTag + "rule-result");
            if (ruleResults != null && ruleResults.Count > 0 && ruleResults.Item(0).FirstChild != null) {
                results.ruleResults = GetResultsListing(ruleResults, searchTag);
            }
            return results;
        }

        private static List<SCAPRuleResult> GetResultsListing(XmlNodeList nodes, string searchTag) {
            List<SCAPRuleResult> ruleResults = new List<SCAPRuleResult>();
            SCAPRuleResult result;
            
            foreach (XmlNode node in nodes) {
                result = new SCAPRuleResult();
                foreach (XmlAttribute attr in node.Attributes) {
                    if (attr.Name == "idref") {
                        result.ruleId = attr.InnerText.Replace("xccdf_mil.disa.stig_rule_","");
                    } else if (attr.Name == "version") {
                        result.ruleVersion = attr.InnerText;
                    }
                }
                if (node.ChildNodes.Count > 0) {
                    foreach (XmlElement child in node.ChildNodes) {
                        // switch on the fields left over to fill them in the SCAPRuleResult class 
                        if (child.Name == searchTag + "result") {
                                // pass or fail
                                result.result = child.InnerText;
                                break;
                        }
                    }
                }
                ruleResults.Add(result);
            }
            return ruleResults;
        }
        
        /// <summary>
        /// Return a checklist raw string based on the SCAP XML file results. 
        /// </summary>
        /// <param name="results">The results list of pass and fail information rules from the SCAP scan</param>
        /// <returns>A checklist raw XML string, if found</returns>
        public static string GenerateChecklistData(SCAPRuleResultSet results) {
            string checklistString = NATSClient.GetArtifactByTemplateTitle(results.title);

            // generate the checklist from reading the template in using a Request/Reply to openrmf.template.read
            if (!string.IsNullOrEmpty(checklistString)) {
                return UpdateChecklistData(results, checklistString, true);
            }            
            // return the default template string
            return checklistString;
        }

        /// <summary>
        /// Return a checklist raw string based on the SCAP XML file results of an existing checklist file.
        /// </summary>
        /// <param name="results">The results list of pass and fail information rules from the SCAP scan</param>
        /// <param name="checklistString">The raw XML of the checklist</param>
        /// <param name="newChecklist">True/False on a new checklist (template). If true, add pass and fail items.</param>
        /// <returns>A checklist raw XML string, if found</returns>
        public static string UpdateChecklistData(SCAPRuleResultSet results, string checklistString, bool newChecklist) {
            // process the raw checklist into the CHECKLIST structure
            CHECKLIST chk = ChecklistLoader.LoadChecklist(checklistString);
            STIG_DATA data;
            STIG_DATA vulnNum;
            SCAPRuleResult result;
            if (chk != null) {
                // if we read in the hostname, then use it in the Checklist data
                if (!string.IsNullOrEmpty(results.hostname)) {
                    chk.ASSET.HOST_NAME = results.hostname;
                }
                // if we have the IP Address, use that as well
                if (!string.IsNullOrEmpty(results.ipaddress)) {
                    chk.ASSET.HOST_IP = results.ipaddress;
                }
                // if we have the MAC Address, use that as well
                if (!string.IsNullOrEmpty(results.macaddress)) {
                    chk.ASSET.HOST_MAC = results.macaddress;
                }
                // if we have the FQDN, use that as well
                if (!string.IsNullOrEmpty(results.fqdn)) {
                    chk.ASSET.HOST_FQDN = results.fqdn;
                }

                string findingDetails = string.Format("Tool: {0}\nTime: {1}\nResult: ", results.scanTool, results.scanTime);

                // for each VULN see if there is a rule matching the rule in the 
                foreach (VULN v in chk.STIGS.iSTIG.VULN) {
                    data = v.STIG_DATA.Where(y => y.VULN_ATTRIBUTE == "Rule_Ver").FirstOrDefault();
                    vulnNum = v.STIG_DATA.Where(y => y.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault();
                    // if we find the VULN id and the Rule Ver, and the VULN id is NOT in the list of "locked" you can update it
                    if (data != null && vulnNum != null) {
                        // find if there is a matching rule
                        result = results.ruleResults.Where(z => z.ruleVersion.ToLower() == data.ATTRIBUTE_DATA.ToLower()).FirstOrDefault();
                        if (result != null) {
                            // set the status
                            // only mark fails IF this is a new one, otherwise leave alone
                            if (result.result.ToLower() == "fail") {
                                v.STATUS = "Open";
                                v.FINDING_DETAILS = findingDetails + "fail";
                            } 
                            // mark the pass on any checklist item we find that passed
                            else if (result.result.ToLower() == "pass") {
                                v.STATUS = "NotAFinding";
                                v.FINDING_DETAILS = findingDetails + "pass";
                            }
                            // mark the not_applicable on any checklist item we find that passed
                            else if (result.result.ToLower() == "notapplicable") {
                                v.STATUS = "Not_Applicable";
                                v.FINDING_DETAILS = findingDetails + "Not Applicable";
                            }
                        }
                    }
                }
            } else {
                // this is not a valid file
                return "";
            }
            // serialize into a string again
            System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(chk.GetType());
            using(StringWriter textWriter = new StringWriter())                
            {
                xmlSerializer.Serialize(textWriter, chk);
                checklistString = textWriter.ToString();
            }
            // strip out all the extra formatting crap and clean up the XML to be as simple as possible
            System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(checklistString, System.Xml.Linq.LoadOptions.None);
            checklistString = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            return checklistString;
        }
    }
}