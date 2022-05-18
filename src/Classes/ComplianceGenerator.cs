// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Threading.Tasks;
using openrmf_read_api.Models;
using openrmf_read_api.Models.Compliance;
using System.Collections.Generic;
using System.Linq;

namespace openrmf_read_api.Classes
{
    public static class ComplianceGenerator 
    {
      public static async Task<List<NISTCompliance>> GetSystemControls(string systemId, List<Artifact> checklists, 
          string filter, bool pii, string majorcontrol)
      {
        // for each system
        //  for each checklist in the system
        //    for each VULN listing
        //      for each CCI within the VULN listed
        //        match up the CCI to the NIST, then get the status, checklist and ID, STIG ID , VULN ID, and type and return it
        try {
          // Call the NATS subscription to get all CCI to NIST Major Controls
          List<CciItem> cciItems = NATSClient.GetCCIListing();
          // list of the NIST controls down to the index-to-CCI level we cycle through
          List<NISTControl> controls = CreateListOfNISTControls(cciItems);
          // the end result grouped by control and listing checklists and their status
          List<NISTCompliance> complianceList = new List<NISTCompliance>();
          if (!string.IsNullOrEmpty(majorcontrol)) {
            // filter the list to just do the one major control
            controls = controls.Where(x => x.control == majorcontrol).ToList();
            // get all cciItems where at least one of the major controls exists in the references listing
            List<CciItem> filteredCCIItems = new List<CciItem>();
            foreach (CciItem cci in cciItems) { // see if the major control is in here; if so keep it otherwise discard
              if (cci.references != null && cci.references.Count > 0 && cci.references.Where(x => x.majorControl == majorcontrol).FirstOrDefault() != null)
                filteredCCIItems.Add(cci);
            }
            cciItems = filteredCCIItems;
            filteredCCIItems = null;
          }

          // get all the variables ready
          List<STIG_DATA> sd;
          //Artifact art;
          ComplianceRecord rec;
          List<ControlSet> controlSet;
          ControlSet controlRecord;
          NISTCompliance compliance;
          string host = "";
          int parentIndex = 0;

          // this routine uses "API to API" messaging via NATS to get the list of checklists per system, each checklist, 
          // as well as the controls that we need to run against for generating compliance
          if (checklists != null && checklists.Count > 0) {
            controlSet = NATSClient.GetControlRecords(filter, pii);
            foreach (Artifact art in checklists) {
              //art = NATSClient.GetChecklist(a.InternalId.ToString());
              if (art != null) {
                host = !string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_NAME)? art.CHECKLIST.ASSET.HOST_NAME : "";
                foreach (VULN v in art.CHECKLIST.STIGS.iSTIG.VULN){
                  // grab each CCI and then match to one or more NIST Control records
                  // fill in the compliance record for the control and add the compliance record to that control w/in the larger list
                  sd = v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "CCI_REF").ToList();
                  foreach (STIG_DATA d in sd) {
                    foreach (NISTControl ctrl in controls.Where(x => x.CCI == d.ATTRIBUTE_DATA).ToList()) {
                      // for each CTRL, if it already has a complianceList record for the checklist and this control, then update the record
                      // if no record, then make a new one
                      if (complianceList.Where(z => z.control == ctrl.control).Count() > 0 ) { // should at most be 1
                        compliance = complianceList.Where(z => z.control == ctrl.control).First();
                      }
                      else {
                        compliance = new NISTCompliance();
                        compliance.control = ctrl.control; // major control family
                        compliance.title = "Unknown";
                        controlRecord = controlSet.Where(x => x.number == ctrl.control.Replace(" ", "")).FirstOrDefault();
                        if (controlRecord != null) {
                          compliance.title = controlRecord.title;
                          compliance.sortString = GenerateControlIndexSort(ctrl.index);
                          complianceList.Add(compliance); // add it to the listing
                        }
                        else { // get the generic family name of the control if any if this is an allowed control
                          parentIndex = GetFirstIndex(ctrl.index);
                          if (parentIndex > 0) {
                            controlRecord = controlSet.Where(x => x.number == ctrl.index.Substring(0, parentIndex) || 
                              x.subControlNumber == ctrl.index.Substring(0, parentIndex)).FirstOrDefault();
                            if (controlRecord != null) {
                              if (!string.IsNullOrEmpty(controlRecord.title))
                                compliance.title = controlRecord.title;
                              compliance.sortString = GenerateControlIndexSort(ctrl.index);
                              complianceList.Add(compliance); // add it to the listing
                            }
                          }
                        }
                        // moved this above where there is a real title, otherwise this is a ghost control not used anymore 
                        // also means it is not linked to Low / Moderate / High which is a required param calling this code
                        //compliance.sortString = GenerateControlIndexSort(ctrl.index);
                        //complianceList.Add(compliance); // add it to the listing
                      }
                      // For the compliance record, does it have a listing for the checklist/artifactId
                      if (compliance.complianceRecords.Where(c => c.artifactId == art.InternalId.ToString()).Count() > 0) { // if a new record, will be 0
                        rec = compliance.complianceRecords.Where(c => c.artifactId == art.InternalId.ToString()).First(); //grab the the record to update the status
                        rec.status = GenerateStatus(rec.status, v.STATUS);
                      }
                      else {
                        rec = new ComplianceRecord();
                        rec.artifactId = art.InternalId.ToString();
                        rec.status = v.STATUS;
                        rec.updatedOn = art.updatedOn.Value;
                        rec.title = art.title;
                        rec.stigType = art.stigType;
                        rec.stigRelease = art.stigRelease;
                        rec.hostName = host;
                        compliance.complianceRecords.Add(rec); // add the new compliance record to the control we are making
                      }
                    }
                  }
                }
              }
            }
            // // fill the compliance list with those in the controls not yet in the complianceList but in the valid control set
            List<string> missingIndexes = controls.Where(x => !complianceList.Any(x2 => x2.control == x.control)).Select(y => y.control).Distinct().ToList();
            foreach (string index in missingIndexes) {
              compliance = new NISTCompliance();
              compliance.control = index; // add the control family
              compliance.title = "Unknown";
              //controlRecord = controlSet.Where(x => x.number == index.Replace(" ", "") || x.subControlNumber == index.Replace(" ", "")).FirstOrDefault();
              controlRecord = controlSet.Where(x => x.number == index.Replace(" ", "")).FirstOrDefault();
              if (controlRecord != null) {
                compliance.title = controlRecord.title;
                compliance.sortString = GenerateControlIndexSort(index);
                complianceList.Add(compliance); // add it to the listing
              }
              else { // get the generic family name of the control if any
                parentIndex = GetFirstIndex(index);
                if (parentIndex > 0) {
                  controlRecord = controlSet.Where(x => x.number == index.Substring(0, GetFirstIndex(index)) || 
                    x.subControlNumber == index.Substring(0, GetFirstIndex(index))).FirstOrDefault();
                  if (controlRecord != null) {
                    if (!string.IsNullOrEmpty(controlRecord.title))
                      compliance.title = controlRecord.title;
                      compliance.sortString = GenerateControlIndexSort(index);
                      complianceList.Add(compliance); // add it to the listing
                  }
                }
                else {
                  Console.WriteLine(string.Format("control not found: {0}", index));
                }
              }
              // moved this to above where we find a title, otherwise these are orphaned are not listed
              //compliance.sortString = GenerateControlIndexSort(index);
              //complianceList.Add(compliance); // add it to the listing
            }
            // order by the index, which also groups them by the major control
            return complianceList.OrderBy(x => x.sortString).ToList();
          }
          else
              return null;
        }
        catch (Exception ex) {
            // log it here
            throw ex;
        }
      }

      private static int GetFirstIndex(string term) {
          int space = term.IndexOf(" ");
          int period = term.IndexOf(".");
          if (space < 0 && period < 0)
              return -1;
          else if (space > 0 && period > 0 && space < period ) // see which we hit first
              return space;
          else if (space > 0 && period > 0 && space > period )
              return period;
          else if (space > 0) 
              return space;
          else 
              return period;
      }
      
      // use the rules in https://github.com/Cingulara/openrmf-api-compliance/issues/1: 
      //   if anything is open, mark it open
      //   else if the old one is not open or N/R
      //         anything is not reviewed, mark it not reviewed
      //         else if old is notafinding and new is notafinding or not_reviewed, not a finding
      private static string GenerateStatus(string oldStatus, string newStatus) {
        if (newStatus.ToLower() == "open")
          return newStatus.ToLower();
        else if (oldStatus.ToLower() != "open" && oldStatus.ToLower() != "not_reviewed") { // otherwise keep it the same
          // this was already not_reviewed or it was notafinding from being NaF or N/A
          if (newStatus.ToLower() == "not_reviewed")
            return newStatus.ToLower();
          else
            return "notafinding"; // catch all cause it is either NaF or N/A
        }
        else
          return oldStatus.ToLower(); // was already marked open or not_reviewed
      }

      // for each of the CCI items in the list 
        //   get the CCI ID
        //   cycle through all references
        //     for each reference add a new control record to the list
        //   when done return the list of controls, somewhere around 6,368 of them
      private static List<NISTControl> CreateListOfNISTControls (List<CciItem> cciItems) {
          List<NISTControl> controls = new List<NISTControl>();
          NISTControl control;
          foreach (CciItem item in cciItems) {
              foreach (CciReference reference in item.references) {
                  control = new NISTControl();
                  control.CCI = item.cciId;
                  control.control = reference.majorControl;
                  control.index = reference.index;
                  control.location = reference.location;
                  control.title = reference.title;
                  control.version = reference.version;
                  // now add the unique record from the xml file to the list of all controls
                  controls.Add(control); // add each one as the combination of index and CCI is unique
              }
          }
          return controls;
      }

      private static string GenerateControlIndexSort(string index) {
        string sort = "";
        int dash = index.IndexOf('-')+1;
        sort = index.Substring(0, dash); // gets you the XX and the -

        int space = index.IndexOf(" ", sort.Length);
        if (space > -1) {// there is something
          index = index.Substring(0, space);
        }
        int period = index.IndexOf(".", sort.Length);
        if (period > -1) { // there is a period so shorten it
          index = index.Substring(0, period);
        }
        if (index.Substring(dash).Length == 1) // need to pad the number with a 0}
          sort += "0" + index.Substring(dash);
        else 
          sort += index.Substring(dash);
        return sort;
      }
    }
}