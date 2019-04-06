using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using openstig_read_api.Models;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using System.Xml;

namespace openstig_read_api.Classes
{
    public static class ChecklistLoader
    {        
        public static CHECKLIST LoadChecklist(string rawChecklist) {
            CHECKLIST myChecklist = new CHECKLIST();
            XmlSerializer serializer = new XmlSerializer(typeof(CHECKLIST));
            rawChecklist = rawChecklist.Replace("\n","").Replace("\t","");
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
    }
}