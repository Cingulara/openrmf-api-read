// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using openrmf_read_api.Models;
using System.Xml;
using System.Linq;

namespace openrmf_read_api.Classes
{
    public static class ChecklistLoader
    {        
        public static CHECKLIST LoadChecklist(string rawChecklist) {
            CHECKLIST myChecklist = new CHECKLIST();
            rawChecklist = rawChecklist.Replace("\t","");
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawChecklist);
            XmlNodeList assetList = xmlDoc.GetElementsByTagName("ASSET");
            XmlNodeList vulnList = xmlDoc.GetElementsByTagName("VULN");
            XmlNodeList stiginfoList = xmlDoc.GetElementsByTagName("STIG_INFO");
            // ensure all three are valid otherwise this XML is junk
            if (assetList != null && stiginfoList != null && vulnList != null) {
                // fill in the ASSET listing
                if (assetList.Count >= 1)
                    myChecklist.ASSET = getAssetListing(assetList.Item(0));
                // now get the STIG_INFO Listing
                if (stiginfoList.Count >= 1)
                    myChecklist.STIGS.iSTIG.STIG_INFO = getStigInfoListing(stiginfoList.Item(0));
                // now get the VULN listings until the end!
                if (vulnList.Count > 0) {
                    myChecklist.STIGS.iSTIG.VULN = getVulnerabilityListing(vulnList);
                }
            }            
            return myChecklist;
        }

        private static ASSET getAssetListing(XmlNode node) {
            ASSET asset = new ASSET();
            foreach (XmlElement child in node.ChildNodes)
            {
                switch (child.Name) {
                    case "ROLE":
                        asset.ROLE = child.InnerText;
                        break;
                    case "ASSET_TYPE":
                        asset.ASSET_TYPE = child.InnerText;
                        break;
                    case "MARKING": 
                        asset.MARKING = child.InnerText;
                        break;
                    case "HOST_NAME":
                        asset.HOST_NAME = child.InnerText;
                        break;
                    case "HOST_IP":
                        asset.HOST_IP = child.InnerText;
                        break;
                    case "HOST_MAC":
                        asset.HOST_MAC = child.InnerText;
                        break;
                    case "HOST_FQDN":
                        asset.HOST_FQDN = child.InnerText;
                        break;
                    case "TECH_AREA":
                        asset.TECH_AREA = child.InnerText;
                        break;
                    case "TARGET_KEY":
                        asset.TARGET_KEY = child.InnerText;
                        break;
                    case "WEB_OR_DATABASE":
                        asset.WEB_OR_DATABASE = child.InnerText;
                        break;
                    case "WEB_DB_SITE":
                        asset.WEB_DB_SITE = child.InnerText;
                        break;
                    case "WEB_DB_INSTANCE":
                        asset.WEB_DB_INSTANCE = child.InnerText;
                        break;
                }
            }
            return asset;
        }

        private static STIG_INFO getStigInfoListing(XmlNode node) {
            STIG_INFO info = new STIG_INFO();
            SI_DATA data; // used for the name/value pairs

            // cycle through the children in STIG_INFO and get the SI_DATA
            foreach (XmlElement child in node.ChildNodes) {
                // get the SI_DATA record for SID_DATA and SID_NAME and then return them
                // each SI_DATA has 2
                data = new SI_DATA();
                foreach (XmlElement siddata in child.ChildNodes) {
                    if (siddata.Name == "SID_NAME")
                        data.SID_NAME = siddata.InnerText;
                    else if (siddata.Name == "SID_DATA")
                        data.SID_DATA = siddata.InnerText;
                }
                info.SI_DATA.Add(data);
            }            
            return info;
        }
 
        private static List<VULN> getVulnerabilityListing(XmlNodeList nodes) {
            List<VULN> vulns = new List<VULN>();
            VULN vuln;
            STIG_DATA data;
            foreach (XmlNode node in nodes) {
                vuln = new VULN();
                if (node.ChildNodes.Count > 0) {
                    foreach (XmlElement child in node.ChildNodes) {
                        data = new STIG_DATA();
                        if (child.Name == "STIG_DATA") {
                            foreach (XmlElement stigdata in child.ChildNodes) {
                                if (stigdata.Name == "VULN_ATTRIBUTE")
                                    data.VULN_ATTRIBUTE = stigdata.InnerText;
                                else if (stigdata.Name == "ATTRIBUTE_DATA")
                                    data.ATTRIBUTE_DATA = stigdata.InnerText;
                            }
                            vuln.STIG_DATA.Add(data);
                        }
                        else {
                            // switch on the fields left over to fill them in the VULN class 
                            switch (child.Name) {
                                case "STATUS":
                                    vuln.STATUS = child.InnerText;
                                    break;
                                case "FINDING_DETAILS":
                                    vuln.FINDING_DETAILS = child.InnerText;
                                    break;
                                case "COMMENTS":
                                    vuln.COMMENTS = child.InnerText;
                                    break;
                                case "SEVERITY_OVERRIDE":
                                    vuln.SEVERITY_OVERRIDE = child.InnerText;
                                    break;
                                case "SEVERITY_JUSTIFICATION":
                                    vuln.SEVERITY_JUSTIFICATION = child.InnerText;
                                    break;
                            }
                        }
                    }
                }
                vulns.Add(vuln);
            }
            return vulns;
        }
 
 
        public static string UpdateChecklistVulnerabilityOrder(string rawChecklist) {
            rawChecklist = rawChecklist.Replace("\t","");
            // load the CHECKLIST structure
            CHECKLIST myChecklist = LoadChecklist(rawChecklist);

            string finalChecklistFormat = "";
            bool badFormat = false;
            
            // test the first VULN as they should all be structured the same
            // if the first one is in the right order, off you go! just return the string
            if (myChecklist.STIGS.iSTIG.VULN != null && myChecklist.STIGS.iSTIG.VULN.Count > 0) {
                VULN checkVulnRecord = myChecklist.STIGS.iSTIG.VULN[0];
                if (checkVulnRecord != null ) {
                    if (checkVulnRecord.STIG_DATA.Count > 10) {
                        // check vuln_num
                        if (checkVulnRecord.STIG_DATA[0].VULN_ATTRIBUTE.ToLower() != "vuln_num")
                            badFormat = true;
                        // check severity
                        if (checkVulnRecord.STIG_DATA[1].VULN_ATTRIBUTE.ToLower() != "severity")
                            badFormat = true;
                        // check STIG ID
                        if (checkVulnRecord.STIG_DATA[4].VULN_ATTRIBUTE.ToLower() != "rule_ver")
                            badFormat = true;
                        // check rule title
                        if (checkVulnRecord.STIG_DATA[5].VULN_ATTRIBUTE.ToLower() != "rule_title")
                            badFormat = true;
                        // make sure the last one is CCI_REF
                        if (checkVulnRecord.STIG_DATA[checkVulnRecord.STIG_DATA.Count-1].VULN_ATTRIBUTE.ToLower() != "cci_ref")
                            badFormat = true;
                    }
                }
                else {
                    // send back what junk they sent you!!
                    return rawChecklist;
                }
            }

            // check that the CCI References do not have a space in them
            // if there are multiple 1 CCI_REF record per value -- Evalute-STIG was generating bad ones for the cisco config file scanner
            if (myChecklist.STIGS.iSTIG.VULN != null && myChecklist.STIGS.iSTIG.VULN.Count > 0 &&
                myChecklist.STIGS.iSTIG.VULN.Where(
                    z => z.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "CCI_REF" && x.ATTRIBUTE_DATA.Contains(" ")).FirstOrDefault() != null
                ).FirstOrDefault() != null
            ) {
                badFormat = true;
            }

            if (!badFormat) 
                return rawChecklist; // good to go!

            foreach (VULN currentVuln in myChecklist.STIGS.iSTIG.VULN) {
                // make the new one in the right order, copying data from the current one
                currentVuln.STIG_DATA = getCorrectSTIGDataRecordsOrdered(currentVuln.STIG_DATA);
                // go on to the next one
            }

            System.Xml.Serialization.XmlSerializer writer = new System.Xml.Serialization.XmlSerializer(typeof(CHECKLIST)); 
            using(System.IO.StringWriter textWriter = new System.IO.StringWriter())
            {
                writer.Serialize(textWriter, myChecklist);
                finalChecklistFormat = textWriter.ToString();
            }
            // strip out all the extra formatting crap and clean up the XML to be as simple as possible
            System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(finalChecklistFormat, System.Xml.Linq.LoadOptions.PreserveWhitespace);
            // get the finalized checklist string format
            finalChecklistFormat = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
            rawChecklist = finalChecklistFormat.Substring(finalChecklistFormat.IndexOf("<STIGS>")); 
            // save the rest but redo the top part up to the STIGS area so the XML looks pretty
            rawChecklist = string.Format("<?xml version=\"1.0\" encoding=\"UTF-8\"?><CHECKLIST><ASSET><ROLE>{0}</ROLE><ASSET_TYPE>{1}</ASSET_TYPE><MARKING>{2}</MARKING><HOST_NAME>{3}</HOST_NAME><HOST_IP>{4}</HOST_IP><HOST_MAC>{5}</HOST_MAC><HOST_FQDN>{6}</HOST_FQDN><TARGET_COMMENT></TARGET_COMMENT><TECH_AREA>{7}</TECH_AREA><TARGET_KEY>{8}</TARGET_KEY><WEB_OR_DATABASE>{9}</WEB_OR_DATABASE><WEB_DB_SITE>{10}</WEB_DB_SITE><WEB_DB_INSTANCE>{11}</WEB_DB_INSTANCE></ASSET>",
                myChecklist.ASSET.ROLE,myChecklist.ASSET.ASSET_TYPE,myChecklist.ASSET.MARKING,
                myChecklist.ASSET.HOST_NAME,myChecklist.ASSET.HOST_IP,
                myChecklist.ASSET.HOST_MAC,myChecklist.ASSET.HOST_FQDN,myChecklist.ASSET.TECH_AREA,
                myChecklist.ASSET.TARGET_KEY,myChecklist.ASSET.WEB_OR_DATABASE,myChecklist.ASSET.WEB_DB_SITE,
                myChecklist.ASSET.WEB_DB_INSTANCE).Trim() + rawChecklist.Trim();

            return RecordGenerator.CleanData(rawChecklist.Trim());
        }

        private static List<STIG_DATA> getCorrectSTIGDataRecordsOrdered (List<STIG_DATA> currentVulnData) {
            List<STIG_DATA> newOrderedData = new List<STIG_DATA>();
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Vuln_Num"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Severity"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Group_Title"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Rule_ID"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Rule_Ver"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Rule_Title"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Vuln_Discuss"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "IA_Controls"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Check_Content"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Fix_Text"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "False_Positives"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "False_Negatives"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Documentable"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Mitigations"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Potential_Impact"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Third_Party_Tools"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Mitigation_Control"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Responsibility"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Security_Override_Guidance"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Check_Content_Ref"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Weight"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "Class"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "STIGRef"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "TargetKey"));
            newOrderedData.Add(getSTIGDataFromList(currentVulnData, "STIG_UUID"));
            // Legacy ID can be multiples, so do 1 at a time
            newOrderedData.AddRange(getSTIGDataListFromCurrentList(currentVulnData, "LEGACY_ID"));
            // CCI REF can be multiple so do 1 at a time
            newOrderedData.AddRange(getSTIGDataCCIREFFromCurrentList(currentVulnData, "CCI_REF"));
            return newOrderedData;
        }

        // method to look through the listing and find the correct value pair
        private static STIG_DATA getSTIGDataFromList (List<STIG_DATA> currentVulnData, string attribute) {
            STIG_DATA newData = new STIG_DATA();
            // get the actual correct attribute/data combination
            if (currentVulnData.Where(z => z.VULN_ATTRIBUTE == attribute).FirstOrDefault() != null) {
                newData.VULN_ATTRIBUTE = attribute;
                newData.ATTRIBUTE_DATA = currentVulnData.Where(z => z.VULN_ATTRIBUTE == attribute).First().ATTRIBUTE_DATA;
            }
            else {
                newData.VULN_ATTRIBUTE = attribute;
                newData.ATTRIBUTE_DATA = "";
            }
            return newData;
        }

        // method to look through the listing and find the correct value pairs from 1 or more records
        private static List<STIG_DATA> getSTIGDataListFromCurrentList (List<STIG_DATA> currentVulnData, string attribute) {
            List<STIG_DATA> newDataList = new List<STIG_DATA>();
            STIG_DATA newData;
            // get the actual correct attribute/data combination
            foreach (STIG_DATA multipleId in currentVulnData.Where(z => z.VULN_ATTRIBUTE == attribute).ToList()) {
                newData = new STIG_DATA();
                newData.VULN_ATTRIBUTE = attribute;
                newData.ATTRIBUTE_DATA = multipleId.ATTRIBUTE_DATA;
                // add to the list we return
                newDataList.Add(newData);
            }
            return newDataList;
        }

        // method to look through the CCI_REF listing and find the correct value pairs from 1 or more records
        // also if more than one record based on spaces in between data, separate them out
        private static List<STIG_DATA> getSTIGDataCCIREFFromCurrentList (List<STIG_DATA> currentVulnData, string attribute) {
            List<STIG_DATA> newDataList = new List<STIG_DATA>();
            STIG_DATA newData;
            // get the actual correct attribute/data combination
            foreach (STIG_DATA multipleId in currentVulnData.Where(z => z.VULN_ATTRIBUTE == attribute).ToList()) {
                if (multipleId.ATTRIBUTE_DATA.IndexOf(" ") > 0) {
                    foreach (string cci in multipleId.ATTRIBUTE_DATA.Split(" ")) {
                        newData = new STIG_DATA();
                        newData.VULN_ATTRIBUTE = attribute;
                        newData.ATTRIBUTE_DATA = cci; // the individiual CCI_REF
                        // add to the list we return
                        newDataList.Add(newData);
                    }
                } 
                else {
                    newData = new STIG_DATA();
                    newData.VULN_ATTRIBUTE = attribute;
                    newData.ATTRIBUTE_DATA = multipleId.ATTRIBUTE_DATA;
                    // add to the list we return
                    newDataList.Add(newData);
                }
            }
            return newDataList;
        }

    }
}