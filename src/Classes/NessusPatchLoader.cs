// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using openrmf_read_api.Models;
using System.Xml;
using System;

namespace openrmf_read_api.Classes
{
    public static class NessusPatchLoader
    {        
        public static NessusPatchData LoadPatchData(string rawNessusPatchFile) {
            NessusPatchData myPatchData = new NessusPatchData();            
            rawNessusPatchFile = rawNessusPatchFile.Replace("\t","");
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawNessusPatchFile);

            XmlNodeList reportList = xmlDoc.GetElementsByTagName("Report");
            XmlNodeList reportHostList = xmlDoc.GetElementsByTagName("ReportHost");
            // ensure all three are valid otherwise this XML is junk
            if (reportList != null && reportHostList != null) {
                // fill in the Report name
                if (reportList.Count >= 1)
                    myPatchData.reportName  = getReportName(reportList.Item(0));
                // now get the ReportHost Listing
                if (reportHostList.Count > 0) {
                    myPatchData.summary = getReportHostListing(reportHostList);
                }
            }            
            return myPatchData;
        }

        private static string getReportName(XmlNode node) {
            XmlAttributeCollection colAttributes = node.Attributes;
            string title = "";
            foreach (XmlAttribute attr in colAttributes) {
                if (attr.Name == "name") {
                    title = attr.Value;
                }
                break;
            }
            return title;
        }

        private static List<NessusPatchSummary> getReportHostListing(XmlNodeList nodes) {
            List<NessusPatchSummary> summaryListing = new List<NessusPatchSummary>();
            NessusPatchSummary summary = new NessusPatchSummary();
            XmlAttributeCollection colAttributes;
            string hostname = "";
            string netbiosname = "";
            string operatingSystem = "";
            string systemType = "";
            bool credentialed = false;
            string ipAddress = "";

            foreach (XmlNode node in nodes) {
                // reset the variables for each reporthost listing
                hostname = "";
                netbiosname = "";
                operatingSystem = "";
                systemType = "";
                credentialed = false;
                ipAddress = "";
                colAttributes = node.Attributes;
                foreach (XmlAttribute attr in colAttributes) {
                    if (attr.Name == "name") {
                        hostname = SanitizeHostname(attr.Value);
                        break;
                    }
                }              
                if (node.ChildNodes.Count > 0) {
                    foreach (XmlElement child in node.ChildNodes) {
                        if (child.Name == "HostProperties") {
                            // for each child node in here
                            netbiosname = "";
                            operatingSystem = "";
                            systemType = "";
                            credentialed = false;
                            ipAddress = "";
                            foreach (XmlElement hostChild in child.ChildNodes) {
                                // get the child
                                foreach (XmlAttribute childAttr in hostChild.Attributes) {
                                    // cycle through attributes where attribute.innertext == netbios-name
                                    if (childAttr.InnerText == "netbios-name") {
                                        netbiosname = hostChild.InnerText; // get the outside child text;
                                    } else if (childAttr.InnerText == "hostname") {
                                        hostname = hostChild.InnerText; // get the outside child text;
                                    } else if (childAttr.InnerText == "operating-system") {
                                        operatingSystem = hostChild.InnerText; // get the outside child text;
                                    } else if (childAttr.InnerText == "system-type") {
                                        systemType = hostChild.InnerText; // get the outside child text;
                                    } else if (childAttr.InnerText == "Credentialed_Scan") {
                                        bool.TryParse(hostChild.InnerText, out credentialed); // get the outside child text;
                                    } else if (childAttr.InnerText == "host-rdns") {
                                        ipAddress = hostChild.InnerText; // get the outside child text;
                                    }
                                }// for each childAttr in hostChild
                            } // for each hostChild
                        }
                        else if (child.Name == "ReportItem") {
                            // get the report host name
                            // get all ReportItems and their attributes in the tag 
                            colAttributes = child.Attributes;
                            summary = new NessusPatchSummary();
                            // set the hostname and other host data for every single record
                            summary.hostname = hostname;
                            summary.operatingSystem = operatingSystem;
                            summary.ipAddress = SanitizeHostname(ipAddress); // if an IP clean up the information octets
                            summary.systemType = systemType;
                            summary.credentialed = credentialed;
                            // get all the attributes
                            foreach (XmlAttribute attr in colAttributes) {
                                if (attr.Name == "severity") {
                                    // store the integer
                                    summary.severity = Convert.ToInt32(attr.Value);
                                } else if (attr.Name == "pluginID") {
                                    summary.pluginId = attr.Value;
                                } else if (attr.Name == "pluginName") { 
                                    summary.pluginName = attr.Value;
                                } else if (attr.Name == "pluginFamily") {
                                    summary.family = attr.Value;
                                }
                            }
                            // get all the child record data we need
                            foreach (XmlElement reportData in child.ChildNodes) {
                                if (reportData.Name == "description")
                                    summary.description = reportData.InnerText;
                                else if (reportData.Name == "plugin_publication_date")
                                    summary.publicationDate = reportData.InnerText;
                                else if (reportData.Name == "plugin_type")
                                    summary.pluginType = reportData.InnerText;
                                else if (reportData.Name == "risk_factor")
                                    summary.riskFactor = reportData.InnerText;
                                else if (reportData.Name == "synopsis")
                                    summary.synopsis = reportData.InnerText;
                            }
                            // add the record
                            summaryListing.Add(summary);
                        }
                    }
                }
            }
            return summaryListing;
        }

        /// <summary>
        /// Called to remove the first two octets from an IP Address if this is an IP
        /// </summary>
        /// <param name="hostname">The hostname or IP of the system</param>
        /// <returns>
        /// The hostname if just a string, the IP address if an IP with xxx.xxx. to start 
        /// the IP range. So the first two octets are hidden from view for security reasons.
        /// </returns>
        private static string SanitizeHostname(string hostname){
            // if this is not an IP, just return the host
            if (hostname.IndexOf(".") <= 0)
                return hostname;
            else {
                System.Net.IPAddress hostAddress;
                if (System.Net.IPAddress.TryParse(hostname.Trim(), out hostAddress)){
                    // this is an IP address so return the last two octets
                    return "xxx.xxx." + hostAddress.GetAddressBytes()[2] + "." + hostAddress.GetAddressBytes()[3];
                }
                else 
                    return hostname;
            }
        }
    }
}