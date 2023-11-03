// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.
using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using Microsoft.AspNetCore.Http;
using openrmf_read_api.Models;
using System.Text;
using System.IO;
using Microsoft.Extensions.Options;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Authorization;
using NATS.Client;
using Newtonsoft.Json;
using System.Xml;

using openrmf_read_api.Data;
using openrmf_read_api.Classes;

namespace openrmf_read_api.Controllers
{
    [Route("/save")]
    public class SaveController : Controller
    {
        private readonly IArtifactRepository _artifactRepo;
        private readonly ISystemGroupRepository _systemGroupRepo;
        private readonly ILogger<SaveController> _logger;
        private readonly IConnection _msgServer;

        public SaveController(IArtifactRepository artifactRepo, ISystemGroupRepository systemGroupRepo, ILogger<SaveController> logger, IOptions<NATSServer> msgServer)
        {
            _logger = logger;
            _artifactRepo = artifactRepo;
            _msgServer = msgServer.Value.connection;
            _systemGroupRepo = systemGroupRepo;
        }

        #region Checklists

        /// <summary>
        /// DELETE Called from the OpenRMF UI (or external access) to delete a checklist by its ID.
        /// Also deletes all scores for those checklists.
        /// </summary>
        /// <param name="id">The ID of the artifact passed in</param>
        /// <returns>
        /// HTTP Status showing it was deleted or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not delete correctly</response>
        /// <response code="404">If the ID was not found</response>
        [HttpDelete("artifact/{id}")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> DeleteArtifact(string id)
        {
            try
            {
                _logger.LogInformation("Calling DeleteArtifact({0})", id);
                Artifact art = _artifactRepo.GetArtifact(id).Result;
                if (art != null)
                {
                    _logger.LogInformation("Deleting Checklist {0}", id);
                    var deleted = await _artifactRepo.DeleteArtifact(id);
                    if (deleted)
                    {
                        var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                        // publish to the openrmf delete realm the new ID passed in to remove the score
                        _logger.LogInformation("Publishing the openrmf.checklist.delete message for {0}", id);
                        _msgServer.Publish("openrmf.checklist.delete", Encoding.UTF8.GetBytes(id));
                        _msgServer.Flush();
                        // decrement the system # of checklists by 1
                        _logger.LogInformation("Publishing the openrmf.system.count.delete message for {0}", id);
                        _msgServer.Publish("openrmf.system.count.delete", Encoding.UTF8.GetBytes(art.systemGroupId));
                        _msgServer.Flush();
                        _logger.LogInformation("Called DeleteArtifact({0}) successfully", id);

                        // publish an audit event
                        _logger.LogInformation("DeleteArtifact() publish an audit message on a deleted checklist {0}.", id);
                        Audit newAudit = GenerateAuditMessage(claim, "delete checklist");
                        newAudit.message = string.Format("DeleteArtifact() delete a single checklist {0}.", id);
                        newAudit.url = string.Format("DELETE /artifact/{0}", id);
                        _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                        _msgServer.Flush();
                        return Ok();
                    }
                    else
                    {
                        _logger.LogWarning("DeleteArtifact() Checklist id {0} not deleted correctly", id);
                        return NotFound();
                    }
                }
                else
                {
                    _logger.LogWarning("DeleteArtifact() Checklist id {0} not found", id);
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DeleteArtifact() Error Deleting Checklist {0}", id);
                return BadRequest();
            }
        }

        /// <summary>
        /// PUT Updating a checklist record from the UI or external that can update the checklist asset information.
        /// </summary>
        /// <param name="artifactId">The ID of the checklist passed in</param>
        /// <param name="systemGroupId">The ID of the system passed in</param>
        /// <param name="hostname">The hostname of the checklist machine</param>
        /// <param name="domainname">The full domain name of the checklist machine</param>
        /// <param name="techarea">The technology area of the checklist machine</param>
        /// <param name="assettype">The asset type of the checklist machine</param>
        /// <param name="machinerole">The role of the checklist machine</param>
        /// <param name="checklistFile">A new Checklist or SCAP scan results file, if any</param>
        /// <param name="tagList">List of tags to add to a checklist, pipe delimited</param>
        /// <returns>
        /// HTTP Status showing it was updated or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not create correctly</response>
        /// <response code="404">If the system ID was not found</response>
        [HttpPut("artifact/{artifactId}")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> UpdateChecklist(string artifactId, string systemGroupId, string hostname, string domainname,
            string techarea, string assettype, string machinerole, IFormFile checklistFile, string tagList)
        {
            try
            {
                _logger.LogInformation("Calling UpdateChecklist(system: {0}, checklist: {1})", systemGroupId, artifactId);
                // see if this is a valid system
                // update and fill in the same info
                Artifact checklist = _artifactRepo.GetArtifactBySystem(systemGroupId, artifactId).GetAwaiter().GetResult();
                // the new raw checklist string
                string newChecklistString = "";

                if (checklist == null)
                {
                    // not a valid system group Id or checklist Id passed in
                    _logger.LogWarning("UpdateChecklist() Error with the System: {0}, Checklist: {1} not a valid system Id or checklist Id", systemGroupId, artifactId);
                    return NotFound();
                }
                checklist.updatedOn = DateTime.Now;

                var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                // grab the user/system ID from the token if there which is *should* always be
                if (claim != null)
                { // get the value
                    checklist.updatedBy = Guid.Parse(claim.Value);
                }
                // get the raw checklist, put into the classes, update the asset information, then save the checklist back to a string
                CHECKLIST chk = ChecklistLoader.LoadChecklist(checklist.rawChecklist);
                // hostname
                if (!string.IsNullOrEmpty(hostname))
                {
                    // set the checklist asset hostname
                    chk.ASSET.HOST_NAME = hostname;
                    // set the artifact record metadata hostname field
                    checklist.hostName = hostname;
                }
                else
                    chk.ASSET.HOST_NAME = "";
                // domain name
                if (!string.IsNullOrEmpty(domainname))
                    chk.ASSET.HOST_FQDN = domainname;
                else
                    chk.ASSET.HOST_FQDN = "";
                // tech area
                if (!string.IsNullOrEmpty(techarea))
                    chk.ASSET.TECH_AREA = techarea;
                else
                    chk.ASSET.TECH_AREA = "";
                // asset type
                if (!string.IsNullOrEmpty(assettype))
                    chk.ASSET.ASSET_TYPE = assettype;
                else
                    chk.ASSET.ASSET_TYPE = "";
                // role
                if (!string.IsNullOrEmpty(machinerole))
                    chk.ASSET.ROLE = machinerole;
                else
                    chk.ASSET.ROLE = "None";

                checklist.tags = new List<string>(); // by default reset it, add if passed in
                if (!string.IsNullOrEmpty(tagList))
                {
                    string[] tags = tagList.Split('|');
                    foreach (string tag in tags)
                    {
                        if (!string.IsNullOrEmpty(tag))
                            checklist.tags.Add(tag);
                    }
                }

                // serialize into a string again
                System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(chk.GetType());
                using (StringWriter textWriter = new StringWriter())
                {
                    xmlSerializer.Serialize(textWriter, chk);
                    newChecklistString = textWriter.ToString();
                }
                // strip out all the extra formatting crap and clean up the XML to be as simple as possible
                System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(newChecklistString, System.Xml.Linq.LoadOptions.None);
                // save the new serialized checklist record to the database
                checklist.rawChecklist = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

                // now save the new record
                _logger.LogInformation("UpdateChecklist() Saving the updated system: {0}, checklist: {1}", systemGroupId, artifactId);
                await _artifactRepo.UpdateArtifact(artifactId, checklist);
                _logger.LogInformation("Called UpdateChecklist(system:{0}, checklist:{1}) successfully", systemGroupId, artifactId);
                // we are finally done

                // publish an audit event
                _logger.LogInformation("UpdateChecklist() publish an audit message on updating a system: {0}, checklist: {1}.", systemGroupId, artifactId);
                Audit newAudit = GenerateAuditMessage(claim, "update checklist data");
                newAudit.message = string.Format("UpdateChecklist() update the system: {0}, checklist: {1}.", systemGroupId, artifactId);
                newAudit.url = string.Format("PUT /artifact/{0}", artifactId);
                _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                _msgServer.Flush();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "UpdateChecklist() Error Updating the Checklist Data: {0}, checklist: {1}", systemGroupId, artifactId);
                return BadRequest();
            }
        }

        /// <summary>
        /// PUT Updating a checklist vulnerability record from the UI or external that can update the checklist information.
        /// </summary>
        /// <param name="artifactId">The ID of the checklist passed in</param>
        /// <param name="systemGroupId">The ID of the system passed in</param>
        /// <param name="vulnid">The vulnerability ID in the checklist</param>
        /// <param name="status">The vulnerability status</param>
        /// <param name="comments">The vulnerability comments entered</param>
        /// <param name="details">The vulnerability finding details</param>
        /// <param name="severityoverride">The severity override for the vulnerability</param>
        /// <param name="justification">The justification for the severity override</param>
        /// <param name="bulkUpdate">Is this a bulk update across all the same checklist types in this system?</param>
        /// <returns>
        /// HTTP Status showing it was updated or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not create correctly</response>
        /// <response code="404">If the system ID was not found</response>
        [HttpPut("artifact/{artifactId}/vulnid/{vulnid}")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> UpdateChecklistVulnerability(string artifactId, string systemGroupId, string vulnid,
            string status, string comments, string details, string severityoverride, string justification, bool bulkUpdate)
        {
            try
            {
                _logger.LogInformation("Calling UpdateChecklistVulnerability(system: {0}, checklist: {1})", systemGroupId, artifactId);

                // who issued this edit/save
                var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                string newChecklistString = "";
                Artifact checklist;
                List<Artifact> bulkChecklists = new List<Artifact>();
                CHECKLIST chk;
                VULN vulnerability;

                // get the one checklist to base this all on
                checklist = _artifactRepo.GetArtifactBySystem(systemGroupId, artifactId).GetAwaiter().GetResult();
                if (checklist != null)
                    bulkChecklists.Add(checklist); // will only be 1
                else
                {
                    _logger.LogWarning("UpdateChecklistVulnerability() Error with the System: {0}, Checklist: {1}, Vulnerability: {2} not a valid system Id or checklist Id",
                        systemGroupId, artifactId, vulnid);
                    return NotFound();
                }
                if (bulkUpdate)
                {
                    // go get ALL checklists in this System whose stigType and version are the same as the artifactId passed in
                    // we will update every single one of them with the same vulnerability information
                    if (checklist != null)
                    { // valid checklist so go get the type and generate a list
                        bulkChecklists = _artifactRepo.GetArtifactsByStigType(systemGroupId, checklist.stigType).GetAwaiter().GetResult().ToList();
                    }
                }

                // cycle through the list; only 1 for a !bulkUpdate and possibly more than 1 for a bulkUpdate
                foreach (Artifact checklistId in bulkChecklists)
                {
                    // reset the new raw checklist string
                    newChecklistString = "";

                    _logger.LogWarning("UpdateChecklistVulnerability() Saving Vulnerability with the System: {0}, Checklist: {1}, Vulnerability: {2}",
                        systemGroupId, checklistId.InternalId.ToString(), vulnid);

                    checklistId.updatedOn = DateTime.Now;
                    // grab the user/system ID from the token if it is there, which is *should* always be
                    if (claim != null)
                    { // get the value
                        checklistId.updatedBy = Guid.Parse(claim.Value);
                    }
                    else
                    {
                        checklistId.updatedBy = Guid.Empty;
                    }

                    // get the raw checklist, put into the classes, update the asset information, then save the checklist back to a string
                    chk = ChecklistLoader.LoadChecklist(checklistId.rawChecklist);
                    vulnerability = chk.STIGS.iSTIG.VULN.Where(y => vulnid == y.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault().ATTRIBUTE_DATA).FirstOrDefault();
                    if (vulnerability != null)
                    {
                        details = RecordGenerator.DecodeHTML(details);
                        comments = RecordGenerator.DecodeHTML(comments);
                        justification = RecordGenerator.DecodeHTML(justification);

                        if (!string.IsNullOrEmpty(details)) vulnerability.FINDING_DETAILS = details;
                        else vulnerability.FINDING_DETAILS = "";
                        if (!string.IsNullOrEmpty(status)) vulnerability.STATUS = status;
                        if (!string.IsNullOrEmpty(comments)) vulnerability.COMMENTS = comments;
                        else vulnerability.COMMENTS = "";
                        if (!string.IsNullOrEmpty(severityoverride)) vulnerability.SEVERITY_OVERRIDE = severityoverride;
                        if (!string.IsNullOrEmpty(justification)) vulnerability.SEVERITY_JUSTIFICATION = justification;
                        else vulnerability.SEVERITY_JUSTIFICATION = "";
                    }
                    else
                    { // record and keep going in this list 
                        _logger.LogWarning("UpdateChecklistVulnerability() No Vuln found Updating the System Checklist Data: {0}, Checklist: {1}, Vulnerability: {2}",
                            systemGroupId, checklistId.InternalId.ToString(), vulnid);
                        if (!bulkUpdate)
                        {
                            return NotFound();
                        }
                        continue; // go get the next one
                    }

                    // serialize into a string again
                    System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(chk.GetType());
                    using (StringWriter textWriter = new StringWriter())
                    {
                        xmlSerializer.Serialize(textWriter, chk);
                        newChecklistString = textWriter.ToString();
                    }
                    // strip out all the extra formatting crap and clean up the XML to be as simple as possible
                    System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(newChecklistString, System.Xml.Linq.LoadOptions.None);
                    // save the new serialized checklist record to the database
                    checklistId.rawChecklist = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);

                    // now save the new record
                    _logger.LogInformation("UpdateChecklistVulnerability() Saving the updated system: {0}, checklist: {1}, Vulnerability: {2}",
                        systemGroupId, checklistId.InternalId.ToString(), vulnid);
                    bool updateSuccess = await _artifactRepo.UpdateArtifact(checklistId.InternalId.ToString(), checklistId);
                    _logger.LogInformation("Called UpdateChecklistVulnerability(system:{0}, checklist:{1}, Vulnerability: {2}) successfully",
                        systemGroupId, checklistId.InternalId.ToString(), vulnid);

                    // update the vulnerability information for reporting with eventual consistency
                    // De/Serialize into a Dictionary object https://www.newtonsoft.com/json/help/html/DeserializeDictionary.htm
                    // this updates the reporting and the score (so far)
                    if (string.IsNullOrEmpty(details)) details = "";
                    if (string.IsNullOrEmpty(comments)) comments = "";
                    if (string.IsNullOrEmpty(severityoverride)) severityoverride = "";
                    if (string.IsNullOrEmpty(justification)) justification = "";
                    Dictionary<string, string> jsonVulnerabilityInfo = new Dictionary<string, string>();
                    jsonVulnerabilityInfo.Add("finding_details", details);
                    jsonVulnerabilityInfo.Add("status", status);
                    jsonVulnerabilityInfo.Add("comments", comments);
                    jsonVulnerabilityInfo.Add("severity_override", severityoverride);
                    jsonVulnerabilityInfo.Add("severity_justification", justification);
                    jsonVulnerabilityInfo.Add("artifactId", checklistId.InternalId.ToString());
                    jsonVulnerabilityInfo.Add("systemGroupId", checklistId.systemGroupId);
                    jsonVulnerabilityInfo.Add("stigType", checklistId.stigType);
                    jsonVulnerabilityInfo.Add("version", checklistId.version);
                    jsonVulnerabilityInfo.Add("vulnId", vulnid);
                    jsonVulnerabilityInfo.Add("updatedBy", checklistId.updatedBy.ToString());

                    _logger.LogInformation("UpdateChecklistVulnerability(system:{0}, checklist:{1}, Vulnerability: {2}) publishing the updated vulnerability for reporting",
                        systemGroupId, checklistId.InternalId.ToString(), vulnid);
                    _msgServer.Publish("openrmf.checklist.save.vulnerability.update",
                        Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(jsonVulnerabilityInfo))));
                    _msgServer.Flush();

                    // publish an audit event
                    _logger.LogInformation("UpdateChecklistVulnerability() publish an audit message on updating a system: {0}, checklist: {1}, Vulnerability: {2}.",
                        systemGroupId, checklistId.InternalId.ToString(), vulnid);
                    Audit newAudit = GenerateAuditMessage(claim, "update checklist vulnerability data");
                    newAudit.message = string.Format("UpdateChecklistVulnerability() update the system: {0}, checklist: {1}, Vulnerability: {2}.",
                        systemGroupId, checklistId.InternalId.ToString(), vulnid);
                    newAudit.url = string.Format("PUT /artifact/{0}/vulnid/{1}", checklistId.InternalId.ToString(), vulnid);
                    _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                    _msgServer.Flush();
                }

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "UpdateChecklistVulnerability() Error Updating the Checklist Data: {0}, checklist: {1}, Vulnerability: {2}",
                    systemGroupId, artifactId, vulnid);
                return BadRequest();
            }
        }

        /// <summary>
        /// Update the passed in checklist with the latest version and release information
        /// </summary>
        /// <param name="systemGroupId">The id of the system record for this checklist</param>
        /// <param name="artifactId">The id of the checklist</param>
        /// <returns>
        /// HTTP Status and the updated artifact/checklist record. Or a 404.
        /// </returns>
        /// <response code="200">Returns the updated artifact record</response>
        /// <response code="404">If the search did not work correctly</response>
        [HttpPost("upgradechecklist/system/{systemGroupId}/artifact/{artifactId}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> UpgradeChecklistRelease(string systemGroupId, string artifactId)
        {
            try
            {
                _logger.LogInformation("Calling GetLatestTemplate({0}, {1})", systemGroupId, artifactId);
                Artifact art = _artifactRepo.GetArtifactBySystem(systemGroupId, artifactId).GetAwaiter().GetResult();
                if (art != null)
                {
                    Artifact upgradeArtifact = new Artifact();
                    string stigType = "";
                    // make the large CHECKLIST record for the current as-is checklist
                    art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                    SI_DATA data = art.CHECKLIST.STIGS.iSTIG.STIG_INFO.SI_DATA.Where(x => x.SID_NAME == "title").FirstOrDefault();
                    if (data != null)
                    {
                        // get the artifact checklists's actual checklist type from DISA
                        stigType = data.SID_DATA;
                        string rawChecklistData = NATSClient.GetArtifactByTemplateTitle(stigType);
                        if (string.IsNullOrEmpty(rawChecklistData))
                        {
                            _logger.LogWarning("UpgradeChecklistRelease({0}, {1}) is not a valid ID", systemGroupId, artifactId);
                            return NotFound();
                        }
                        var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                        // grab the user/system ID from the token if there which is *should* always be
                        art.updatedOn = DateTime.Now;
                        if (claim != null)
                        { // get the value
                            art.updatedBy = Guid.Parse(claim.Value);
                        }
                        // setup the STIG type, release, and version information
                        upgradeArtifact = GetArtifactTypeReleaseVersion(rawChecklistData);
                        // clean up the data some
                        upgradeArtifact.rawChecklist = SanitizeData(rawChecklistData);
                        // get the updated data to copy into this artifact record
                        art.stigRelease = upgradeArtifact.stigRelease;
                        art.stigType = upgradeArtifact.stigType;
                        art.version = upgradeArtifact.version;

                        // make the large CHECKLIST record for the current as-is checklist
                        art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                        // make the empty CHECKLIST record for the new template
                        upgradeArtifact.CHECKLIST = ChecklistLoader.LoadChecklist(upgradeArtifact.rawChecklist);

                        // copy all asset from old to new, as we will generate a new rawChecklist for the existing record
                        // from the new one we are filling up
                        upgradeArtifact.CHECKLIST.ASSET.ROLE = art.CHECKLIST.ASSET.ROLE;
                        upgradeArtifact.CHECKLIST.ASSET.ASSET_TYPE = art.CHECKLIST.ASSET.ASSET_TYPE;
                        upgradeArtifact.CHECKLIST.ASSET.HOST_NAME = art.CHECKLIST.ASSET.HOST_NAME;
                        upgradeArtifact.CHECKLIST.ASSET.HOST_IP = art.CHECKLIST.ASSET.HOST_IP;
                        upgradeArtifact.CHECKLIST.ASSET.HOST_MAC = art.CHECKLIST.ASSET.HOST_MAC;
                        upgradeArtifact.CHECKLIST.ASSET.HOST_FQDN = art.CHECKLIST.ASSET.HOST_FQDN;
                        upgradeArtifact.CHECKLIST.ASSET.TECH_AREA = art.CHECKLIST.ASSET.TECH_AREA;
                        upgradeArtifact.CHECKLIST.ASSET.TARGET_KEY = art.CHECKLIST.ASSET.TARGET_KEY;
                        upgradeArtifact.CHECKLIST.ASSET.WEB_OR_DATABASE = art.CHECKLIST.ASSET.WEB_OR_DATABASE;
                        upgradeArtifact.CHECKLIST.ASSET.WEB_DB_SITE = art.CHECKLIST.ASSET.WEB_DB_SITE;
                        upgradeArtifact.CHECKLIST.ASSET.WEB_DB_INSTANCE = art.CHECKLIST.ASSET.WEB_DB_INSTANCE;

                        // now copy all the VULN information/listing
                        // 5 fields only: status, severity override, finding details, comments, severity override justification
                        // There may be newer ones in the new one, there may be dropped ones
                        // make sure the VulnNum matches then update
                        VULN vulnerability;
                        string vulnid = "";
                        foreach (VULN v in upgradeArtifact.CHECKLIST.STIGS.iSTIG.VULN)
                        {
                            vulnid = v.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Rule_Ver").FirstOrDefault().ATTRIBUTE_DATA;
                            // see if the updated checklist and current checklist have the same vulnerability
                            if (art.CHECKLIST.STIGS.iSTIG.VULN.Where(y => y.STIG_DATA.Where(q => q.VULN_ATTRIBUTE == "Rule_Ver").FirstOrDefault().ATTRIBUTE_DATA == vulnid).FirstOrDefault() != null)
                            {
                                vulnerability = art.CHECKLIST.STIGS.iSTIG.VULN.Where(y => v.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Rule_Ver").FirstOrDefault().ATTRIBUTE_DATA
                                        == y.STIG_DATA.Where(z => z.VULN_ATTRIBUTE == "Rule_Ver").FirstOrDefault().ATTRIBUTE_DATA).FirstOrDefault();
                                if (vulnerability != null)
                                {
                                    // copy the contents from the older checklist into the newer one
                                    if (!string.IsNullOrEmpty(vulnerability.FINDING_DETAILS)) v.FINDING_DETAILS = vulnerability.FINDING_DETAILS;
                                    if (!string.IsNullOrEmpty(vulnerability.STATUS)) v.STATUS = vulnerability.STATUS;
                                    if (!string.IsNullOrEmpty(vulnerability.COMMENTS)) v.COMMENTS = vulnerability.COMMENTS;
                                    if (!string.IsNullOrEmpty(vulnerability.SEVERITY_OVERRIDE)) v.SEVERITY_OVERRIDE = vulnerability.SEVERITY_OVERRIDE;
                                    if (!string.IsNullOrEmpty(vulnerability.SEVERITY_JUSTIFICATION)) v.SEVERITY_JUSTIFICATION = vulnerability.SEVERITY_JUSTIFICATION;
                                }
                            }
                        }

                        // copy the new setup into the existing record for serialization
                        _logger.LogInformation("UpgradeChecklistRelease(system:{0}, checklist:{1}) Copying the new checklist data into the existing record", systemGroupId, artifactId);
                        art.CHECKLIST = upgradeArtifact.CHECKLIST;
                        upgradeArtifact = null;

                        // serialize into a string again
                        _logger.LogInformation("UpgradeChecklistRelease(system:{0}, checklist:{1}) Making the raw checklist string from the CHECKLIST record", systemGroupId, artifactId);
                        string newChecklistString = "";
                        System.Xml.Serialization.XmlSerializer xmlSerializer = new System.Xml.Serialization.XmlSerializer(art.CHECKLIST.GetType());
                        using (StringWriter textWriter = new StringWriter())
                        {
                            xmlSerializer.Serialize(textWriter, art.CHECKLIST);
                            newChecklistString = textWriter.ToString();
                        }
                        // strip out all the extra formatting crap and clean up the XML to be as simple as possible
                        System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(newChecklistString, System.Xml.Linq.LoadOptions.None);
                        // save the new serialized checklist record to the database
                        art.rawChecklist = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                        // clear out the data we do not need now before saving this updated record
                        art.CHECKLIST = new CHECKLIST();

                        // save the data 
                        _logger.LogInformation("UpgradeChecklistRelease(system:{0}, checklist:{1}) Saving the upgraded checklist", systemGroupId, artifactId);
                        await _artifactRepo.UpdateArtifact(artifactId, art);
                        _logger.LogInformation("Called UpgradeChecklistRelease(system:{0}, checklist:{1}) successfully", systemGroupId, artifactId);

                        // update the score with eventual consistency
                        _logger.LogInformation("UpgradeChecklistRelease(system:{0}, checklist:{1}) publishing the updated checklist for scoring",
                            systemGroupId, artifactId);
                        _msgServer.Publish("openrmf.checklist.save.update", Encoding.UTF8.GetBytes(artifactId));
                        _msgServer.Flush();

                        // publish an audit event
                        _logger.LogInformation("UpgradeChecklistRelease() publish an audit message on upgrading a system: {0}, checklist: {1}.",
                            systemGroupId, artifactId);
                        Audit newAudit = GenerateAuditMessage(claim, "upgraded the checklist release");
                        newAudit.message = string.Format("UpgradeChecklistRelease() upgraded the system: {0}, checklist: {1}.",
                            systemGroupId, artifactId);
                        newAudit.url = string.Format("POST upgradechecklist/system/{0}/artifact/{1}", systemGroupId, artifactId);
                        _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                        _msgServer.Flush();

                        // we are good, send the record with updated data and update the screen
                        return Ok(art);
                    }
                    else
                    {
                        _logger.LogWarning("Called UpgradeChecklistRelease({0}, {1}) with an invalid systemId or artifactId checklist type", systemGroupId, artifactId);
                        return BadRequest("The checklist passed in had an Invalid System Artifact/Checklist Type");
                    }
                }
                else
                {
                    _logger.LogWarning("Called UpgradeChecklistRelease({0}, {1}) with an invalid systemId or artifactId", systemGroupId, artifactId);
                    return BadRequest("The checklist passed in had an Invalid System Artifact/Checklist");
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "UpgradeChecklistRelease({0}, {1}) Error Retrieving Latest Template", systemGroupId, artifactId);
                return BadRequest();
            }
        }

        #endregion

        #region Systems
        /// <summary>
        /// DELETE Called from the OpenRMF UI (or external access) to delete an entire system
        /// and all its checklists and scores by its ID.
        /// </summary>
        /// <param name="id">The ID of the system passed in</param>
        /// <returns>
        /// HTTP Status showing it was deleted or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not delete correctly</response>
        /// <response code="404">If the ID was not found</response>
        [HttpDelete("system/{id}")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> DeleteSystem(string id)
        {
            try
            {
                _logger.LogInformation("Calling DeleteSystem({0})", id);
                SystemGroup sys = _systemGroupRepo.GetSystemGroup(id).Result;
                if (sys != null)
                {
                    var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                    _logger.LogInformation("DeleteSystem() Deleting System {0} and all checklists", id);
                    var deleted = await _systemGroupRepo.DeleteSystemGroup(id);
                    if (deleted)
                    {
                        // publish an audit event
                        _logger.LogInformation("DeleteSystem() publish an audit message on a deleted system {0}.", id);
                        Audit newAudit = GenerateAuditMessage(claim, "delete system");
                        newAudit.message = string.Format("DeleteSystem() delete the entire system {0}.", id);
                        newAudit.url = string.Format("DELETE /system/{0}", id);
                        _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                        _msgServer.Flush();

                        _logger.LogInformation("DeleteSystem() Publishing the openrmf.system.delete message for {0}", id);
                        _msgServer.Publish("openrmf.system.delete", Encoding.UTF8.GetBytes(id));
                        _msgServer.Flush();

                        // get all checklists for this system and delete each one at a time, then run the publish on score delete
                        var checklists = await _artifactRepo.GetSystemArtifacts(id);
                        foreach (Artifact a in checklists)
                        {
                            _logger.LogInformation("DeleteSystem() Deleting Checklist {0} from System {1}", a.InternalId.ToString(), id);
                            var checklistDeleted = await _artifactRepo.DeleteArtifact(a.InternalId.ToString());
                            if (checklistDeleted)
                            {
                                // publish to the openrmf delete realm the new ID passed in to remove the score
                                _logger.LogInformation("DeleteSystem() Publishing the openrmf.checklist.delete message for {0}", a.InternalId.ToString());
                                _msgServer.Publish("openrmf.checklist.delete", Encoding.UTF8.GetBytes(a.InternalId.ToString()));
                                _msgServer.Flush();

                                // publish an audit event
                                _logger.LogInformation("DeleteSystem() publish an audit message on a deleted checklist {0} on system {1}.", a.InternalId.ToString(), id);
                                // reset the object
                                newAudit = new Audit();
                                newAudit = GenerateAuditMessage(claim, "delete checklist from a system delete");
                                newAudit.message = string.Format("DeleteSystem() delete the checklist {0} from the system {1}.", a.InternalId.ToString(), id);
                                newAudit.url = string.Format("DELETE /system/{0}", id);
                                _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                                _msgServer.Flush();
                            }
                        }
                        _logger.LogInformation("DeleteSystem() Finished deleting cleanup for System {0}", id);
                        _logger.LogInformation("Called DeleteSystem({0}) successfully", id);
                        return Ok();
                    }
                    else
                    {
                        _logger.LogWarning("DeleteSystem() System id {0} not deleted correctly", id);
                        return NotFound();
                    }
                }
                else
                {
                    _logger.LogWarning("DeleteSystem() System id {0} not found", id);
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DeleteSystem() Error Deleting System {0}", id);
                return BadRequest();
            }
        }

        /// <summary>
        /// DELETE Called from the OpenRMF UI (or external access) to delete all checklist found with 
        /// the system ID. Also deletes all their scores.
        /// </summary>
        /// <param name="id">The ID of the artifact passed in</param>
        /// <param name="checklistIds">The IDs in an array of all checklists to delete</param>
        /// <returns>
        /// HTTP Status showing it was deleted or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not delete correctly</response>
        /// <response code="404">If the ID was not found</response>
        [HttpDelete("system/{id}/artifacts")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> DeleteSystemChecklists(string id, [FromForm] string checklistIds)
        {
            try
            {
                _logger.LogInformation("Calling DeleteSystemChecklists({0})", id);
                SystemGroup sys = _systemGroupRepo.GetSystemGroup(id).Result;
                if (sys != null)
                {
                    string[] ids;
                    _logger.LogInformation("DeleteSystemChecklists() Deleting System {0} checklists only", id);
                    if (string.IsNullOrEmpty(checklistIds))
                    {
                        // get all checklists for this system and delete each one at a time, then run the publish on score delete
                        var checklists = await _artifactRepo.GetSystemArtifacts(id);
                        List<string> lstChecklistIds = new List<string>();
                        foreach (Artifact a in checklists)
                        {
                            // add the ID as a string to the list
                            lstChecklistIds.Add(a.InternalId.ToString());
                        }
                        // push the list to an array
                        ids = lstChecklistIds.ToArray();
                    }
                    else
                    {
                        // split on the command and get back an array to cycle through
                        ids = checklistIds.Split(",");
                    }

                    // now cycle through all the IDs and run with it
                    var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                    Audit newAudit;
                    foreach (string checklist in ids)
                    {
                        _logger.LogInformation("DeleteSystemChecklists() Deleting Checklist {0} from System {1}", checklist, id);
                        var checklistDeleted = await _artifactRepo.DeleteArtifact(checklist);
                        if (checklistDeleted)
                        {
                            // publish to the openrmf delete realm the new ID passed in to remove the score
                            _logger.LogInformation("DeleteSystemChecklists() Publishing the openrmf.checklist.delete message for {0}", checklist);
                            _msgServer.Publish("openrmf.checklist.delete", Encoding.UTF8.GetBytes(checklist));
                            _msgServer.Flush();
                            // decrement the system # of checklists by 1
                            _logger.LogInformation("DeleteSystemChecklists() Publishing the openrmf.system.count.delete message for {0}", id);
                            _msgServer.Publish("openrmf.system.count.delete", Encoding.UTF8.GetBytes(id));
                            _msgServer.Flush();

                            // publish an audit event
                            _logger.LogInformation("DeleteSystemChecklists() publish an audit message on a deleted checklist {0} on system {1}.", checklist, id);
                            newAudit = new Audit();
                            newAudit = GenerateAuditMessage(claim, "delete checklist from list of checklists selected");
                            newAudit.message = string.Format("DeleteSystemChecklists() delete the checklist {0} from the system {1}.", checklist, id);
                            newAudit.url = string.Format("DELETE system/{0}/artifacts", id);
                            _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                            _msgServer.Flush();
                        }
                    }

                    _logger.LogInformation("DeleteSystemChecklists() Finished deleting checklists for System {0}", id);
                    _logger.LogInformation("Called DeleteSystemChecklists({0}) successfully", id);
                    return Ok();
                }
                else
                {
                    _logger.LogWarning("DeleteSystemChecklists() System id {0} not found", id);
                    return NotFound();
                }
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DeleteSystemChecklists() Error Deleting System Checklists {0}", id);
                return BadRequest();
            }
        }

        /// <summary>
        /// POST Creating a system record from the UI or external that can set the title, description, 
        /// and attach a Nessus file.
        /// </summary>
        /// <param name="title">The title/name of the system</param>
        /// <param name="description">The description of the system</param>
        /// <param name="nessusFile">A Nessus scan file, if any</param>
        /// <returns>
        /// HTTP Status showing it was created or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not create correctly</response>
        /// <response code="404">If the system ID was not found</response>
        [HttpPost("system")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> CreateSystemGroup(string title, string description, IFormFile nessusFile)
        {
            try
            {
                _logger.LogInformation("Calling CreateSystemGroup({0})", title);
                string rawNessusFile = string.Empty;
                var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();

                // create the record to use
                SystemGroup sg = new SystemGroup();
                sg.created = DateTime.Now;

                if (!string.IsNullOrEmpty(title))
                {
                    sg.title = title;
                }
                else
                {
                    _logger.LogInformation("CreateSystemGroup() No title passed so returning a 404");
                    BadRequest("You must enter a title.");
                }

                // get the file for Nessus if there is one
                if (nessusFile != null)
                {
                    if (nessusFile.FileName.ToLower().EndsWith(".nessus"))
                    {
                        _logger.LogInformation("CreateSystemGroup() Reading the System {0} Nessus ACAS file", title);
                        using (var reader = new StreamReader(nessusFile.OpenReadStream()))
                        {
                            rawNessusFile = reader.ReadToEnd();
                        }
                        rawNessusFile = SanitizeData(rawNessusFile);
                        sg.nessusFilename = nessusFile.FileName; // keep a copy of the file name that was uploaded
                    }
                    else
                    {
                        // log this is a bad Nessus ACAS scan file
                        return BadRequest();
                    }
                }

                // add the information
                if (!string.IsNullOrEmpty(description))
                {
                    sg.description = description;
                }
                if (!string.IsNullOrEmpty(rawNessusFile))
                {
                    // save the XML to use later on
                    sg.rawNessusFile = rawNessusFile;
                }

                // grab the user/system ID from the token if there which is *should* always be
                if (claim != null)
                { // get the value
                    sg.createdBy = Guid.Parse(claim.Value);
                }

                // save the new record
                _logger.LogInformation("CreateSystemGroup() Saving the System {0}", title);
                var record = await _systemGroupRepo.AddSystemGroup(sg);
                _logger.LogInformation("Called CreateSystemGroup({0}) successfully", title);
                // we are finally done

                // publish an audit event
                _logger.LogInformation("CreateSystemGroup() publish an audit message on creating a new system {0}.", sg.title);
                Audit newAudit = GenerateAuditMessage(claim, "create system");
                newAudit.message = string.Format("CreateSystemGroup() create a new system {0}.", sg.title);
                newAudit.url = string.Format("POST /system");
                _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                if (!string.IsNullOrEmpty(rawNessusFile))
                    _msgServer.Publish("openrmf.system.patchscan", Encoding.UTF8.GetBytes(record.InternalId.ToString()));
                _msgServer.Flush();

                return Ok(record);
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "CreateSystemGroup() Error Creating the System {0}", title);
                return BadRequest();
            }
        }

        /// <summary>
        /// PUT Updating a system record from the UI or external that can update the title, description, 
        /// and attach a Nessus file.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system passed in</param>
        /// <param name="title">The title/name of the system</param>
        /// <param name="description">The description of the system</param>
        /// <param name="nessusFile">A Nessus scan file, if any</param>
        /// <returns>
        /// HTTP Status showing it was updated or that there is an error.
        /// </returns>
        /// <response code="200">Returns the newly created item</response>
        /// <response code="400">If the item did not create correctly</response>
        /// <response code="404">If the system ID was not found</response>
        [HttpPut("system/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> UpdateSystem(string systemGroupId, string title, string description, IFormFile nessusFile)
        {
            try
            {
                _logger.LogInformation("Calling UpdateSystem({0})", systemGroupId);
                // see if this is a valid system
                // update and fill in the same info
                SystemGroup sg = _systemGroupRepo.GetSystemGroup(systemGroupId).GetAwaiter().GetResult();
                if (sg == null)
                {
                    // not a valid system group ID passed in
                    _logger.LogWarning("UpdateSystem() Error with the System {0} not a valid system Id", systemGroupId);
                    return NotFound();
                }
                sg.updatedOn = DateTime.Now;

                string rawNessusFile = string.Empty;
                var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();

                // get the new / updated file for Nessus if there is one
                if (nessusFile != null)
                {
                    if (nessusFile.FileName.ToLower().EndsWith(".nessus"))
                    {
                        _logger.LogInformation("UpdateSystem() Reading the the System {0} Nessus ACAS file", systemGroupId);
                        using (var reader = new StreamReader(nessusFile.OpenReadStream()))
                        {
                            rawNessusFile = reader.ReadToEnd();
                        }
                        rawNessusFile = SanitizeData(rawNessusFile);
                        sg.nessusFilename = nessusFile.FileName; // keep a copy of the file name that was uploaded
                    }
                    else
                    {
                        // log this is a bad Nessus ACAS scan file
                        _logger.LogWarning("UpdateSystem() Error with the Nessus uploaded file for System {0}", systemGroupId);
                        return BadRequest("Invalid Nessus file");
                    }
                }

                // if it is update the information
                if (!string.IsNullOrEmpty(description))
                {
                    sg.description = description;
                }
                if (!string.IsNullOrEmpty(rawNessusFile))
                {
                    // save the XML to use later on
                    sg.rawNessusFile = rawNessusFile;
                }
                if (!string.IsNullOrEmpty(title))
                {
                    if (sg.title.Trim() != title.Trim())
                    {
                        // change in the title so update it
                        _logger.LogInformation("UpdateSystem() Updating the System Title for {0} to {1}", systemGroupId, title);
                        sg.title = title;
                        // if the title is different, it should change across all other checklist files
                        // publish to the openrmf update system realm the new title we can use it
                        _msgServer.Publish("openrmf.system.update." + systemGroupId.Trim(), Encoding.UTF8.GetBytes(title));
                        _msgServer.Flush();
                    }
                }
                // grab the user/system ID from the token if there which is *should* always be
                if (claim != null)
                { // get the value
                    sg.updatedBy = Guid.Parse(claim.Value);
                }
                // save the new record
                _logger.LogInformation("UpdateSystem() Saving the updated system {0}", systemGroupId);
                await _systemGroupRepo.UpdateSystemGroup(systemGroupId, sg);
                _logger.LogInformation("Called UpdateSystem({0}) successfully", systemGroupId);
                // we are finally done

                // publish an audit event
                _logger.LogInformation("UpdateSystem() publish an audit message on creating a new system {0}.", sg.title);
                Audit newAudit = GenerateAuditMessage(claim, "create system");
                newAudit.message = string.Format("UpdateSystem() update the system ({0}) {1}.", sg.InternalId.ToString(), sg.title);
                newAudit.url = string.Format("PUT /system/{0}", systemGroupId);
                _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                if (!string.IsNullOrEmpty(rawNessusFile))
                    _msgServer.Publish("openrmf.system.patchscan", Encoding.UTF8.GetBytes(systemGroupId));
                _msgServer.Flush();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "UpdateSystem() Error Updating the System {0}", systemGroupId);
                return BadRequest();
            }
        }

        /// <summary>
        /// DELETE Updating a system package record by removing the Nessus patch scan file.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system passed in</param>
        /// <returns>
        /// HTTP Status showing it was updated or that there is an error.
        /// </returns>
        /// <response code="200">Returns if the system package was updated</response>
        /// <response code="400">If the item did not update correctly</response>
        /// <response code="404">If the system ID was not found</response>
        [HttpDelete("system/{systemGroupId}/patchscan")]
        [Authorize(Roles = "Administrator,Editor")]
        public async Task<IActionResult> DeleteSystemPatchScanFile(string systemGroupId)
        {
            try
            {
                _logger.LogInformation("Calling DeleteSystemPatchScanFile({0})", systemGroupId);
                // see if this is a valid system
                // update and fill in the same info
                SystemGroup sg = _systemGroupRepo.GetSystemGroup(systemGroupId).GetAwaiter().GetResult();
                if (sg == null)
                {
                    // not a valid system group ID passed in
                    _logger.LogWarning("DeleteSystemPatchScanFile() Error with the System {0} not a valid system Id", systemGroupId);
                    return NotFound();
                }
                sg.updatedOn = DateTime.Now;
                sg.rawNessusFile = "";
                sg.nessusFilename = "";

                var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                // grab the user/system ID from the token if there which is *should* always be
                if (claim != null)
                { // get the value
                    sg.updatedBy = Guid.Parse(claim.Value);
                }
                // save the new record
                _logger.LogInformation("DeleteSystemPatchScanFile() Saving the updated system package and removed the patch scan file {0}", systemGroupId);
                await _systemGroupRepo.UpdateSystemGroup(systemGroupId, sg);
                _logger.LogInformation("Called UpdateSystem({0}) successfully", systemGroupId);
                // we are finally done

                // publish an audit event
                _logger.LogInformation("DeleteSystemPatchScanFile() publish an audit message on updating the system package {0}.", sg.title);
                Audit newAudit = GenerateAuditMessage(claim, "update system");
                newAudit.message = string.Format("DeleteSystemPatchScanFile() update the system and remove the system package patch scan file ({0}) {1}.", sg.InternalId.ToString(), sg.title);
                newAudit.url = string.Format("DELETE /system/{0}/patchscan", systemGroupId);
                _msgServer.Publish("openrmf.audit.save", Encoding.UTF8.GetBytes(Compression.CompressString(JsonConvert.SerializeObject(newAudit))));
                _msgServer.Flush();

                return Ok();
            }
            catch (Exception ex)
            {
                _logger.LogError(ex, "DeleteSystemPatchScanFile() Error Updating the System {0}", systemGroupId);
                return BadRequest();
            }
        }

        #endregion

        #region Private Functions
        private Artifact GetArtifactTypeReleaseVersion(string rawChecklist)
        {
            Artifact newArtifact = new Artifact();
            newArtifact.rawChecklist = rawChecklist;

            // parse the checklist and get the data needed
            rawChecklist = rawChecklist.Replace("\t", "");
            XmlDocument xmlDoc = new XmlDocument();
            xmlDoc.LoadXml(rawChecklist);

            // get the title and release which is a list of children of child nodes buried deeper :face-palm-emoji:
            XmlNodeList stiginfoList = xmlDoc.GetElementsByTagName("STIG_INFO");
            foreach (XmlElement child in stiginfoList.Item(0).ChildNodes)
            {
                if (child.FirstChild.InnerText == "releaseinfo")
                    newArtifact.stigRelease = child.LastChild.InnerText;
                else if (child.FirstChild.InnerText == "title")
                    newArtifact.stigType = child.LastChild.InnerText;
                else if (child.FirstChild.InnerText == "version")
                    newArtifact.version = child.LastChild.InnerText;
            }

            // shorten the names a bit
            if (newArtifact != null && !string.IsNullOrEmpty(newArtifact.stigType))
            {
                newArtifact.stigType = newArtifact.stigType.Replace("Security Technical Implementation Guide", "STIG");
                newArtifact.stigType = newArtifact.stigType.Replace("Windows", "WIN");
                newArtifact.stigType = newArtifact.stigType.Replace("Application Security and Development", "ASD");
                newArtifact.stigType = newArtifact.stigType.Replace("Microsoft Internet Explorer", "MSIE");
                newArtifact.stigType = newArtifact.stigType.Replace("Red Hat Enterprise Linux", "REL");
                newArtifact.stigType = newArtifact.stigType.Replace("MS SQL Server", "MSSQL");
                newArtifact.stigType = newArtifact.stigType.Replace("Server", "SVR");
                newArtifact.stigType = newArtifact.stigType.Replace("Workstation", "WRK");
            }
            if (newArtifact != null && !string.IsNullOrEmpty(newArtifact.stigRelease))
            {
                newArtifact.stigRelease = newArtifact.stigRelease.Replace("Release: ", "R"); // i.e. R11, R2 for the release number
                newArtifact.stigRelease = newArtifact.stigRelease.Replace("Benchmark Date:", "dated");
            }
            return newArtifact;
        }

        private string SanitizeData(string rawdata)
        {
            return rawdata.Replace("\t", "");
        }

        private Audit GenerateAuditMessage(System.Security.Claims.Claim claim, string action)
        {
            Audit audit = new Audit();
            audit.program = "Save API";
            audit.created = DateTime.Now;
            audit.action = action;
            if (claim != null)
            {
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

        #endregion
    }
}
