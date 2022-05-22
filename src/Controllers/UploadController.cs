// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using System.IO;
using System.Text;
using Microsoft.AspNetCore.Http;
using System.Xml;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Logging;
using NATS.Client;
using Microsoft.AspNetCore.Authorization;
using Newtonsoft.Json;

using openrmf_read_api.Data;
using openrmf_read_api.Models;
using openrmf_read_api.Classes;

namespace openrmf_read_api.Controllers
{
    [Route("/upload")]
    public class UploadController : Controller
    {
	    private readonly IArtifactRepository _artifactRepo;
	    private readonly ISystemGroupRepository _systemRepo;
      private readonly ILogger<UploadController> _logger;
      private readonly IConnection _msgServer;

        public UploadController(IArtifactRepository artifactRepo, ILogger<UploadController> logger, IOptions<NATSServer> msgServer, ISystemGroupRepository systemRepo )
        {
            _logger = logger;
            _artifactRepo = artifactRepo;
            _systemRepo = systemRepo;
            _msgServer = msgServer.Value.connection;
        }

        /// <summary>
        /// POST Called from the OpenRMF UI (or external access) to create one or more checklist/artifact records within a system.
        /// </summary>
        /// <param name="checklistFiles">The CKL files to add into the system</param>
        /// <param name="systemGroupId">The system Id if adding to a current system</param>
        /// <param name="system">A new System title if creating a new system from checklists</param>
        /// <returns>
        /// HTTP Status showing they were created or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly updated item</response>
        /// <response code="400">If the item did not update correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpPost]
        [Authorize(Roles = "Administrator,Editor,Assessor")]
        public async Task<IActionResult> UploadNewChecklist(List<IFormFile> checklistFiles, string systemGroupId, string system)
        {
          try {
            _logger.LogInformation("Calling UploadNewChecklist() with {0} checklists", checklistFiles.Count.ToString());
            if (checklistFiles.Count > 0) {

              // grab the user/system ID from the token if there which is *should* always be
              var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
              // make sure the SYSTEM GROUP is valid here and then add the files...
              SystemGroup sg;
              SystemGroup recordSystem = null;

              if (string.IsNullOrEmpty(systemGroupId)) {
                sg = new SystemGroup();
                sg.title = system;
                sg.created = DateTime.Now;
                if (claim != null && claim.Value != null) {
                  sg.createdBy = Guid.Parse(claim.Value);
                }
                recordSystem = _systemRepo.AddSystemGroup(sg).GetAwaiter().GetResult();
              } else {
                sg = await _systemRepo.GetSystemGroup(systemGroupId);
                if (sg == null) {
                  sg = new SystemGroup();
                  sg.title = "None";
                  sg.created = DateTime.Now;
                  if (claim != null && claim.Value != null) {
                    sg.createdBy = Guid.Parse(claim.Value);
                  }
                recordSystem = _systemRepo.AddSystemGroup(sg).GetAwaiter().GetResult();
                }
                else {
                  sg.updatedOn = DateTime.Now;
                  if (claim != null && claim.Value != null) {
                    sg.updatedBy = Guid.Parse(claim.Value);
                  }
                  var updated = _systemRepo.UpdateSystemGroup(systemGroupId, sg).GetAwaiter().GetResult();
                }
              }

              // result we send back
              UploadResult uploadResult = new UploadResult();
              bool updatedChecklist = false;

              // now go through the Checklists and set them up
              foreach(IFormFile file in checklistFiles) {
                try {
                    string rawChecklist =  string.Empty;
                    updatedChecklist = false;

                    if (file.FileName.ToLower().EndsWith(".xml")) {
                      // if an XML XCCDF SCAP scan file
                      _logger.LogInformation("UploadNewChecklist() parsing the SCAP Scan file for {0}.", file.FileName.ToLower());
                      using (var reader = new StreamReader(file.OpenReadStream()))
                      {
                        // read in the file
                        string xmlfile = reader.ReadToEnd();
                        // pull out the rule IDs and their results of pass or fail and the title/type of SCAP scan done
                        SCAPRuleResultSet results = SCAPScanResultLoader.LoadSCAPScan(xmlfile);
                        // get the rawChecklist data so we can move on
                        // generate a new checklist from a template based on the type and revision
                        rawChecklist = SCAPScanResultLoader.GenerateChecklistData(results);
                      }
                    }
                    else if (file.FileName.ToLower().EndsWith(".ckl")) {
                      // if a CKL file
                      _logger.LogInformation("UploadNewChecklist() parsing the Checklist CKL file for {0}.", file.FileName.ToLower());
                      using (var reader = new StreamReader(file.OpenReadStream()))
                      {
                          rawChecklist = reader.ReadToEnd();  
                      }
                    }
                    else {
                      // log this is a bad file
                      return BadRequest();
                    }

                    // clean up any odd data that can mess us up moving around, via JS, and such
                    _logger.LogInformation("UploadNewChecklist() sanitizing the checklist for {0}.", file.FileName.ToLower());
                    rawChecklist = SanitizeData(rawChecklist);

                    // create the new record for saving into the DB
                    Artifact newArtifact = MakeArtifactRecord(rawChecklist);
                    Artifact oldArtifact = null; 
                    // add the system record ID to the Artifact to know how to query it
                    _logger.LogInformation("UploadNewChecklist() setting the title of the checklist {0}.", file.FileName.ToLower());
                    if (recordSystem != null) {
                      newArtifact.systemGroupId = recordSystem.InternalId.ToString();
                      // store the title for ease of use
                      newArtifact.systemTitle = recordSystem.title;
                    }
                    else {
                      newArtifact.systemGroupId = sg.InternalId.ToString();
                      // store the title for ease of use
                      newArtifact.systemTitle = sg.title;
                    }
                    
                    // if there is a hostname, see if this is an UPDATE or a new one based on the checklist type for this system
                    if (!string.IsNullOrEmpty(newArtifact.hostName) && newArtifact.hostName.ToLower() != "unknown") {
                      // we got this far, so it is a valid checklist. Let's see if it is an update or a new one that we uploaded.
                      oldArtifact = await _artifactRepo.GetArtifactBySystemHostnameAndType(newArtifact.systemGroupId, newArtifact.hostName, newArtifact.stigType);
                      if (oldArtifact != null && oldArtifact.createdBy != Guid.Empty) {
                        _logger.LogInformation("UploadNewChecklist({0}) this is an update, not a new checklist", newArtifact.systemGroupId);
                        // this is an update of an older one, keep the createdBy intact
                        newArtifact.createdBy = oldArtifact.createdBy;
                        newArtifact.InternalId = oldArtifact.InternalId; // copy the ID over
                        // keep it a part of the same system group
                        oldArtifact.updatedBy = Guid.Parse(claim.Value);
                        oldArtifact.updatedOn = DateTime.Now;
                        updatedChecklist = true;
                      } else {
                        newArtifact.createdBy = Guid.Parse(claim.Value);
                        newArtifact.created = DateTime.Now;
                      }
                      //oldArtifact = null;
                    }

                    if (claim != null) { // get the value
                      _logger.LogInformation("UploadNewChecklist() setting the created by ID of the checklist {0}.", file.FileName.ToLower());
                      newArtifact.createdBy = Guid.Parse(claim.Value);
                      if (sg.createdBy == Guid.Empty)
                        sg.createdBy = Guid.Parse(claim.Value);
                      else 
                        sg.updatedBy = Guid.Parse(claim.Value);
                    }

                    // save the artifact record and checklist to the database
                    Artifact record = new Artifact();
                    if (!updatedChecklist || oldArtifact == null) {
                      _logger.LogInformation("UploadNewChecklist() saving the new checklist {0} to the database", file.FileName.ToLower());
                      record  = await _artifactRepo.AddArtifact(newArtifact);
                    }
                    else {
                      _logger.LogInformation("UploadNewChecklist() saving the updated checklist {0} to the database", file.FileName.ToLower());
                      // we need to update in place and copy the VULN information over
                      // we also only copy the Status and Finding Details if an XML SCAP XCCDF file; otherwise we copy all the fields
                      // then we call update in place
                      // get the checklist data loaded correctly
                      oldArtifact.CHECKLIST = ChecklistLoader.LoadChecklist(oldArtifact.rawChecklist);
                      newArtifact.CHECKLIST = ChecklistLoader.LoadChecklist(newArtifact.rawChecklist);
                      VULN oldVulnerability;
                      string vulnid = "";
                      foreach (VULN v in newArtifact.CHECKLIST.STIGS.iSTIG.VULN) {
                          vulnid = v.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault().ATTRIBUTE_DATA;
                          // if the vulnerability number/id matches, and 
                          //  1) it is either Open or Not a Finding from a SCAP; OR 
                          //  2) a checklist file finding from an uploaded CKL
                          if (oldArtifact.CHECKLIST.STIGS.iSTIG.VULN.Where(y => y.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault()
                            .ATTRIBUTE_DATA  == vulnid).FirstOrDefault() != null && 
                            (!file.FileName.ToLower().EndsWith(".xml") || v.STATUS.ToLower() == "notafinding" || v.STATUS.ToLower() == "open")) { 

                              // find the vulnerability
                              oldVulnerability = oldArtifact.CHECKLIST.STIGS.iSTIG.VULN.Where(y => y.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault().ATTRIBUTE_DATA == vulnid).FirstOrDefault();
                              if (v.STATUS.ToLower() == "notafinding")
                                  oldVulnerability.STATUS = "NotAFinding";
                              else if (v.STATUS.ToLower() == "open")
                                  oldVulnerability.STATUS = "Open";
                              else if (v.STATUS.ToLower() == "not_applicable")
                                  oldVulnerability.STATUS = "Not_Applicable";
                              else if (v.STATUS.ToLower() == "not_reviewed")
                                  oldVulnerability.STATUS = "Not_Reviewed";

                              if (!string.IsNullOrEmpty(v.FINDING_DETAILS)) oldVulnerability.FINDING_DETAILS = v.FINDING_DETAILS;
                              else oldVulnerability.FINDING_DETAILS = "";
                              if (!file.FileName.ToLower().EndsWith(".xml")) {
                                  if (!string.IsNullOrEmpty(v.COMMENTS)) oldVulnerability.COMMENTS = v.COMMENTS;
                                  else oldVulnerability.COMMENTS = "";
                                  if (!string.IsNullOrEmpty(v.SEVERITY_OVERRIDE)) oldVulnerability.SEVERITY_OVERRIDE = v.SEVERITY_OVERRIDE;
                                  else oldVulnerability.SEVERITY_OVERRIDE = "";
                                  if (!string.IsNullOrEmpty(v.SEVERITY_JUSTIFICATION)) oldVulnerability.SEVERITY_JUSTIFICATION = v.SEVERITY_JUSTIFICATION;
                                  else oldVulnerability.SEVERITY_JUSTIFICATION = "";
                              }
                          }
                          // cycle to the next one
                      }

                      // format the XML string
                      string newChecklistString = "";
                      System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(oldArtifact.CHECKLIST.GetType());
                      using(StringWriter textWriter = new StringWriter())                
                      {
                          xmlSerializer.Serialize(textWriter, oldArtifact.CHECKLIST);
                          newChecklistString = textWriter.ToString();
                      }
                      // strip out all the extra formatting crap and clean up the XML to be as simple as possible
                      System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(newChecklistString, System.Xml.Linq.LoadOptions.None);
                      // save the new serialized checklist record to the database
                      oldArtifact.rawChecklist = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                      // clear out the data we do not need now before saving this updated record
                      oldArtifact.CHECKLIST = new CHECKLIST();
                      // save the update                      
                      await _artifactRepo.UpdateArtifact(oldArtifact.InternalId.ToString(), oldArtifact);
                    }

                    _logger.LogInformation("UploadNewChecklist() saved the checklist {0} to the database.", file.FileName.ToLower());

                    // add to the number of successful uploads
                    uploadResult.successful++;

                    // publish to the openrmf save message the artifactId we can use
                    if (!updatedChecklist) {
                      _logger.LogInformation("UploadNewChecklist() publish a message on a new checklist {0} for the scoring of it.", file.FileName.ToLower());
                      _msgServer.Publish("openrmf.checklist.save.new", Encoding.UTF8.GetBytes(record.InternalId.ToString()));
                      // publish to update the system checklist count
                      _logger.LogInformation("UploadNewChecklist() publish a message on a new checklist {0} for updating the count of checklists in the system.", file.FileName.ToLower());
                      _msgServer.Publish("openrmf.system.count.add", Encoding.UTF8.GetBytes(record.systemGroupId));
                      _msgServer.Flush();

                      // publish an audit event
                      _logger.LogInformation("UploadNewChecklist() publish an audit message on a new checklist {0}.", file.FileName.ToLower());
                      Audit newAudit = GenerateAuditMessage(claim, "add checklist");
                      newAudit.message = string.Format("UploadNewChecklist() uploaded a new checklist {0} {1} in system group ({2}) {3}.", file.FileName.ToLower(),record.title, sg.InternalId.ToString(), sg.title);
                      newAudit.url = "POST /";
                      _msgServer.Publish("openrmf.audit.upload", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                      _msgServer.Flush();
                    } else { // send a different message, this is an update
                      _logger.LogInformation("UploadNewChecklist({0}) publishing the updated checklist for scoring", oldArtifact.InternalId.ToString());
                      _msgServer.Publish("openrmf.checklist.save.update", Encoding.UTF8.GetBytes(oldArtifact.InternalId.ToString()));
                      _msgServer.Flush();
                      _logger.LogInformation("Called UploadNewChecklist() with an updated checklist {0} successfully", oldArtifact.InternalId.ToString());
                      
                      // publish an audit event
                      _logger.LogInformation("UploadNewChecklist() publish an audit message on an updated checklist {0}.", file.FileName.ToLower());
                      Audit newAudit = GenerateAuditMessage(claim, "update checklist");
                      newAudit.message = string.Format("UploadNewChecklist() updated checklist {0} {1} with file {2} in system group ({3}) {4}.", oldArtifact.InternalId.ToString(), oldArtifact.title, file.FileName.ToLower(), sg.InternalId.ToString(), sg.title);
                      newAudit.url = "POST /";
                      _msgServer.Publish("openrmf.audit.upload", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                      _msgServer.Flush();
                    }
                }
                catch (Exception ex) {
                  // add to the list of failed uploads
                  uploadResult.failed++;
                  uploadResult.failedUploads.Add(file.FileName);
                  // log it
                  _logger.LogError(ex, "UploadNewChecklist() error on checklist file not parsing right: {0}.", file.FileName.ToLower());
                  // see if there are any left
                }
              }
              _logger.LogInformation("Called UploadNewChecklist() with {0} checklists successfully", checklistFiles.Count.ToString());
              return Ok(uploadResult);
            }
            else {              
              _logger.LogWarning("Called UploadNewChecklist() with NO checklists!");
              return BadRequest();
            }
          }
          catch (Exception ex) {
              _logger.LogError(ex, "Error uploading checklist file");
              return BadRequest();
          }
        }
      
        /// <summary>
        /// POST Called from the OpenRMF UI to create a new checklist in a system package from a template page.
        /// </summary>
        /// <param name="systemGroupId">The system Id if adding to a current system</param>
        /// <param name="templateId">The template Id to use</param>
        /// <returns>
        /// HTTP Status showing it was created or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly updated item</response>
        /// <response code="400">If the item did not create correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpPost("{systemGroupId}/template/{templateId}")]
        [Authorize(Roles = "Administrator,Editor,Assessor")]
        public async Task<IActionResult> CreateNewChecklist(string systemGroupId, string templateId)
        {
            try {
                // grab the user/system ID from the token if there which is *should* always be
                var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                // make sure the SYSTEM GROUP is valid here and then add the checklist
                SystemGroup sg = await _systemRepo.GetSystemGroup(systemGroupId);
                if (sg == null) {
                    _logger.LogWarning("CreateNewChecklist() passed an invalid systemGroupId {0}.", systemGroupId);
                    return NotFound();
                }

                // go get this as a template ID and if not valid, return NotFound() as well
                string rawChecklist =  NATSClient.GetArtifactByTemplateTitle(templateId);
                if (string.IsNullOrEmpty(rawChecklist)) {
                    _logger.LogWarning("CreateNewChecklist() passed an invalid templateId {0}.", templateId);
                    return NotFound();
                }
                rawChecklist = SanitizeData(rawChecklist);

                // create the new record for saving into the DB
                Artifact newArtifact = MakeArtifactRecord(rawChecklist);
                // add the system record ID to the Artifact to know how to query it
                _logger.LogInformation("CreateNewChecklist() setting the title of the checklist {0}.", newArtifact.title);
                newArtifact.systemGroupId = sg.InternalId.ToString();
                // store the title for ease of use
                newArtifact.systemTitle = sg.title;
                newArtifact.created = DateTime.Now;

                if (claim != null) { // get the value
                    _logger.LogInformation("CreateNewChecklist() setting the created by ID of the checklist {0}.", newArtifact.title);
                    newArtifact.createdBy = Guid.Parse(claim.Value);
                    sg.updatedBy = Guid.Parse(claim.Value);
                }

                // save the artifact record and checklist to the database
                _logger.LogInformation("CreateNewChecklist() saving the new checklist {0} to the database", newArtifact.title);
                Artifact record  = await _artifactRepo.AddArtifact(newArtifact);
                _logger.LogInformation("CreateNewChecklist() saved the checklist {0} to the database.", newArtifact.title);

                sg.updatedOn = DateTime.Now;
                var updated = _systemRepo.UpdateSystemGroup(systemGroupId, sg).GetAwaiter().GetResult();

                // publish to the openrmf save message the artifactId we can use
                if (record != null)
                _logger.LogInformation("CreateNewChecklist() publish a message on a new checklist {0} for the scoring of it.", record.title);
                _msgServer.Publish("openrmf.checklist.save.new", Encoding.UTF8.GetBytes(record.InternalId.ToString()));
                // publish to update the system checklist count
                _logger.LogInformation("CreateNewChecklist() publish a message on a new checklist {0} for updating the count of checklists in the system.", newArtifact.title);
                _msgServer.Publish("openrmf.system.count.add", Encoding.UTF8.GetBytes(record.systemGroupId));
                _msgServer.Flush();

                // publish an audit event
                _logger.LogInformation("CreateNewChecklist() publish an audit message on a new checklist {0}.", newArtifact.title);
                Audit newAudit = GenerateAuditMessage(claim, "add checklist");
                newAudit.message = string.Format("CreateNewChecklist() created a new checklist {0} in system group ({1}) {2}.", record.title, sg.InternalId.ToString(), sg.title);
                newAudit.url = string.Format("POST /{0}/template/{1}", systemGroupId, templateId);
                _msgServer.Publish("openrmf.audit.upload", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                _msgServer.Flush();

                _logger.LogInformation("Called CreateNewChecklist() for checklist {0} successfully", newArtifact.title);
                return Ok();

            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error creating checklist from template");
                return BadRequest();
            }
        }

      // this parses the text and system, generates the pieces, and returns the artifact to save
      private Artifact MakeArtifactRecord(string rawChecklist) {
        Artifact newArtifact = new Artifact();
        newArtifact.created = DateTime.Now;
        newArtifact.updatedOn = DateTime.Now;
        newArtifact.rawChecklist = rawChecklist;
        newArtifact.hostName = "Unknown"; // default

        // parse the checklist and get the data needed
        rawChecklist = rawChecklist.Replace("\t","");
        XmlDocument xmlDoc = new XmlDocument();
        xmlDoc.LoadXml(rawChecklist);

        XmlNodeList assetList = xmlDoc.GetElementsByTagName("ASSET");
        // get the host name from here
        foreach (XmlElement child in assetList.Item(0).ChildNodes)
        {
          switch (child.Name) {
            case "HOST_NAME":
              if (!string.IsNullOrEmpty(child.InnerText)) 
                newArtifact.hostName = child.InnerText;
              break;
          }
        }
        // get the title and release which is a list of children of child nodes buried deeper :face-palm-emoji:
        XmlNodeList stiginfoList = xmlDoc.GetElementsByTagName("STIG_INFO");
        foreach (XmlElement child in stiginfoList.Item(0).ChildNodes) {
          if (child.FirstChild.InnerText == "releaseinfo")
            newArtifact.stigRelease = child.LastChild.InnerText;
          else if (child.FirstChild.InnerText == "title")
            newArtifact.stigType = child.LastChild.InnerText;
          else if (child.FirstChild.InnerText == "version")
              newArtifact.version = child.LastChild.InnerText;
        }

        // shorten the names a bit
        if (newArtifact != null && !string.IsNullOrEmpty(newArtifact.stigType)){
          newArtifact.stigType = newArtifact.stigType.Replace("Security Technical Implementation Guide", "STIG");
          newArtifact.stigType = newArtifact.stigType.Replace("Windows", "WIN");
          newArtifact.stigType = newArtifact.stigType.Replace("Application Security and Development", "ASD");
          newArtifact.stigType = newArtifact.stigType.Replace("Microsoft Internet Explorer", "MSIE");
          newArtifact.stigType = newArtifact.stigType.Replace("Red Hat Enterprise Linux", "REL");
          newArtifact.stigType = newArtifact.stigType.Replace("MS SQL Server", "MSSQL");
          newArtifact.stigType = newArtifact.stigType.Replace("Server", "SVR");
          newArtifact.stigType = newArtifact.stigType.Replace("Workstation", "WRK");
        }
        if (newArtifact != null && !string.IsNullOrEmpty(newArtifact.stigRelease)) {
          newArtifact.stigRelease = newArtifact.stigRelease.Replace("Release: ", "R"); // i.e. R11, R2 for the release number
          newArtifact.stigRelease = newArtifact.stigRelease.Replace("Benchmark Date:","dated");
        }
        return newArtifact;
      }
      private string SanitizeData (string rawdata) {
        return rawdata.Replace("\t","");
      }

      private Audit GenerateAuditMessage(System.Security.Claims.Claim claim, string action) {
        Audit audit = new Audit();
        audit.program = "Upload API";
        audit.created = DateTime.Now;
        audit.action = action;
        if (claim != null) {
          audit.userid = claim.Value;
          var fullname = claim.Subject.Claims.Where(x => x.Type == "name").FirstOrDefault();
          if (fullname != null) 
            audit.fullname = fullname.Value;
          var username = claim.Subject.Claims.Where(x => x.Type == "preferred_username").FirstOrDefault();
          if (username != null) 
            audit.username = username.Value;
          var useremail = claim.Subject.Claims.Where(x => x.Type.Contains("emailaddress")).FirstOrDefault();
          if (useremail != null) 
            audit.email = useremail.Value;
        }
        return audit;
      }
    }
}
