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
            colAttributes = nodes[0].Attributes;
            foreach (XmlAttribute attr in colAttributes) {
                if (attr.Name == "name") {
                    hostname = attr.Value;
                }
                break;
            }
            foreach (XmlNode node in nodes) {                
                if (node.ChildNodes.Count > 0) {
                    foreach (XmlElement child in node.ChildNodes) {
                        if (child.Name == "ReportItem") {
                            // get the report host name
                            // get all ReportItems and their attributes in the tag 
                            colAttributes = child.Attributes;
                            summary = new NessusPatchSummary();
                            // set the hostname
                            summary.hostname = hostname;
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
                            summaryListing.Add(summary);
                        }
                    }
                }
            }
            return summaryListing;
        }
    }
}