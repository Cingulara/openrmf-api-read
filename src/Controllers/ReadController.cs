// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using openrmf_read_api.Classes;
using openrmf_read_api.Models;
using System.IO;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Logging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using openrmf_read_api.Data;

namespace openrmf_read_api.Controllers
{
    [Route("/")]
    public class ReadController : Controller
    {
	    private readonly IArtifactRepository _artifactRepo;
        private readonly ISystemGroupRepository _systemGroupRepo;
        private readonly ILogger<ReadController> _logger;

        public ReadController(IArtifactRepository artifactRepo, ISystemGroupRepository systemGroupRepo, ILogger<ReadController> logger)
        {
            _logger = logger;
            _artifactRepo = artifactRepo;
            _systemGroupRepo = systemGroupRepo;
        }

        #region System Functions and API calls

        /// <summary>
        /// GET The list of checklists for the given System ID
        /// </summary>
        /// <param name="system">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the list of system records 
        /// exported to an XLSX file to download.
        /// </returns>
        /// <response code="200">Returns the Artifact List of records for the passed in system in XLSX format</response>
        /// <response code="400">If the item did not query correctly</response>
        [HttpGet("export")]
        [Authorize(Roles = "Administrator,Reader,Assessor")]
        public async Task<IActionResult> ExportChecklistListing(string system = null)
        {
            try {
                _logger.LogInformation("Calling ExportChecklistListing({0})", system);
                IEnumerable<Artifact> artifacts;
                // if they pass in a system, get all for that system
                if (string.IsNullOrEmpty(system))
                {
                    _logger.LogInformation("ExportChecklistListing() Getting a listing of all checklists to export to XLSX");
                    artifacts = await _artifactRepo.GetAllArtifacts();
                }
                else {
                    _logger.LogInformation(string.Format("ExportChecklistListing() Getting a listing of all {0} checklists to export to XLSX", system));
                    artifacts = await _artifactRepo.GetSystemArtifacts(system);
                }
                if (artifacts != null && artifacts.Count() > 0) {
                    // starting row number for data
                    uint rowNumber = 6;

                    // create the XLSX in memory and send it out
                    var memory = new MemoryStream();
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                    {
                        // Add a WorkbookPart to the document.
                        WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                        workbookpart.Workbook = new Workbook();
                        
                        // add styles to workbook
                        WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();

                        // Add a WorksheetPart to the WorkbookPart.
                        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        // add stylesheet to use cell formats 1 - 4
                        wbsp.Stylesheet = ExcelStyleSheet.GenerateStylesheet();

                        DocumentFormat.OpenXml.Spreadsheet.Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
                        if (lstColumns == null) { // generate the column listings we need with custom widths
                            lstColumns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 30, CustomWidth = true }); // col System
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 100, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 20, CustomWidth = true }); // NaF
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 5, Max = 5, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 20, CustomWidth = true }); // N/R
                            worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                        }

                        // Add Sheets to the Workbook.
                        Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                        // Append a new worksheet and associate it with the workbook.
                        Sheet sheet = new Sheet() { Id = spreadSheet.WorkbookPart.
                            GetIdOfPart(worksheetPart), SheetId = 1, Name = "ChecklistListing" };
                        sheets.Append(sheet);
                        // Get the sheetData cell table.
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                        DocumentFormat.OpenXml.Spreadsheet.Cell newCell = null;

                        DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("OpenRMF by Cingulara and Tutela");
                        sheetData.Append(row);
                        if (artifacts != null && artifacts.Count() > 0) {
                            row = MakeChecklistInfoRow("Checklist Listing", artifacts.FirstOrDefault().systemTitle, 2);
                        } else {
                            row = MakeChecklistInfoRow("Checklist Listing", system, 2);
                        }
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Printed Date", DateTime.Now.ToString("MM/dd/yy hh:mm"),3);
                        sheetData.Append(row);
                        row = MakeChecklistListingHeaderRows(rowNumber);
                        sheetData.Append(row);

                        uint styleIndex = 0; // use this for 4, 5, 6, or 7 for status
                        Score checklistScore;

                        // cycle through the checklists and grab the score for each individually
                        _logger.LogInformation("ExportChecklistListing({0}) cycling through checklists to list", system);
                        foreach (Artifact art in artifacts.OrderBy(x => x.title).OrderBy(y => y.systemTitle).ToList()) {
                            //art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                            try {
                                checklistScore = NATSClient.GetChecklistScore(art.InternalId.ToString());
                            }
                            catch (Exception ex) {
                                _logger.LogWarning(ex, "No score found for artifact {0}", art.InternalId.ToString());
                                checklistScore = new Score();
                            }
                            rowNumber++;

                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", art.systemTitle.Trim().ToLower() != "none"? art.systemTitle : "", styleIndex);
                            // now cycle through the rest of the items
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(art.title);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalNotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalNotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalOpen.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalNotReviewed.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            sheetData.Append(row);

                            // now add the cat 1, 2, and 3 findings
                            rowNumber++; // CAT 1
                            row = MakeDataRow(rowNumber, "A", "", styleIndex);
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("CAT 1");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1NotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1NotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1NotReviewed.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            sheetData.Append(row);
                            rowNumber++; // CAT 2
                            row = MakeDataRow(rowNumber, "A", "", styleIndex);
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("CAT 2");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2NotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2NotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2NotReviewed.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            sheetData.Append(row);
                            rowNumber++; // CAT 3
                            row = MakeDataRow(rowNumber, "A", "", styleIndex);
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("CAT 3");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3NotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3NotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3NotReviewed.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            sheetData.Append(row);
                        }

                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        spreadSheet.Close();
                        memory.Seek(0, SeekOrigin.Begin);
                        _logger.LogInformation("Called ExportChecklistListing({0}) successfully", system);
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ChecklistListing.xlsx");
                    }
                }
                else {
                    _logger.LogInformation("Calling ExportChecklistListing({0}) but had no checklists to show", system);
                    return NotFound();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "ExportChecklistListing({0}) Error Retrieving Artifacts for Exporting", system);
                return NotFound();
            }
        } 
        
        /// <summary>
        /// GET The list of systems in the database
        /// </summary>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the list of system records.
        /// </returns>
        /// <response code="200">Returns the System List of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If there are no systems yet</response>
        [HttpGet("systems")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ListArtifactSystems()
        {
            try {
                _logger.LogInformation("Calling ListArtifactSystems()");
                IEnumerable<SystemGroup> systems;
                systems = await _systemGroupRepo.GetAllSystemGroups();
                if (systems == null) {
                    _logger.LogWarning("Calling ListArtifactSystems() returned 0 systems");
                    return NotFound();
                }
                _logger.LogInformation("Called ListArtifactSystems() successfully");
                foreach(SystemGroup sys in systems) {
                    // remove the Nessus file as we don't need it
                    sys.rawNessusFile = "";
                }
                return Ok(systems);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "ListArtifactSystems() Error listing all checklist systems");
                return BadRequest();
            }
        }
        
        /// <summary>
        /// GET The list of checklists for the given System ID
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the list of checklists records.
        /// </returns>
        /// <response code="200">Returns the Artifact List of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("systems/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ListArtifactsBySystem(string systemGroupId)
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling ListArtifactsBySystem({0})", systemGroupId);
                    IEnumerable<Artifact> systemChecklists;
                    systemChecklists = await _artifactRepo.GetSystemArtifacts(systemGroupId);
                    if (systemChecklists == null) {
                        _logger.LogWarning("Calling ListArtifactsBySystem({0}) returned no checklists", systemGroupId);
                        return NotFound();
                    }
                    // we do not need all the data for the raw checklist in the listing, too bloated
                    foreach(Artifact a in systemChecklists) {
                        a.rawChecklist = "";
                    }
                    _logger.LogInformation("Called ListArtifactsBySystem({0}) successfully", systemGroupId);
                    return Ok(systemChecklists);
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "ListArtifactsBySystem() Error listing all checklists for system {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called ListArtifactsBySystem() with no system ID");
                return BadRequest(); // no systemGroupId entered
            }
        }
        
        /// <summary>
        /// GET The system record based on the ID.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the system record data.
        /// </returns>
        /// <response code="200">Returns the SystemGroup record</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("system/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetSystem(string systemGroupId)
        {
            _logger.LogInformation("Calling GetSystem({0})", systemGroupId);
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    SystemGroup systemRecord;
                    systemRecord = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                    if (systemRecord == null) {                        
                        _logger.LogWarning("Calling GetSystem({0}) with an invalid system ID", systemGroupId);
                        return NotFound();
                    }
                    _logger.LogInformation("Called GetSystem({0}) successfully", systemGroupId);
                    return Ok(systemRecord);
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "GetSystem() Error getting the system record for {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Calling GetSystem() with no System ID");
                return BadRequest(); // no systemGroupId entered
            }
        }

        /// <summary>
        /// GET Download the Nessus file for the system, if any
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the nessus file downloaded as a 
        /// .nessus text file, basically XML.
        /// </returns>
        /// <response code="200">Returns the Nessus file in XML format</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid or the Nessus file is not there</response>
        [HttpGet("/system/{systemGroupId}/downloadnessus")]
        [Authorize(Roles = "Administrator,Editor,Download")]
        public async Task<IActionResult> DownloadSystemNessus(string systemGroupId)
        {
            try {
                _logger.LogInformation("Calling DownloadSystemNessus({0})", systemGroupId);
                SystemGroup sg = new SystemGroup();
                sg = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                if (sg != null) {
                    if (!string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogInformation("Called DownloadSystemNessus({0}) successfully", systemGroupId);
                        return Ok(sg.rawNessusFile);
                    }
                    else {
                        _logger.LogWarning("Calling DownloadSystemNessus({0}) returned an empty Nessus file", systemGroupId);
                        return NotFound();
                    }
                }
                else {
                    _logger.LogWarning("Calling DownloadSystemNessus({0}) with an invalid System ID", systemGroupId);
                    return NotFound();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "DownloadSystemNessus() Error Retrieving System Nessus file for Download");
                return NotFound();
            }
        }

        /// <summary>
        /// GET The count of patch information from the attached Nessus scan file, if there.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the count of 
        /// critical, high, medium, low, and info items from the Nessus ACAS Patch scan.
        /// </returns>
        /// <response code="200">Returns the count of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in does not have a valid Nessus file</response>
        [HttpGet("system/{systemGroupId}/nessuspatchsummary")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetNessusPatchSummary(string systemGroupId)
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling GetNessusPatchSummary({0})", systemGroupId);
                    SystemGroup sg = new SystemGroup();
                    sg = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                    if (sg == null) {
                        _logger.LogWarning("GetNessusPatchSummary({0}) an invalid system record");
                        return NotFound();
                    }
                    if (string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogWarning("GetNessusPatchSummary({0}) system record has no Nessus patch file to use");
                        return NotFound();
                    }
                    // load the NessusPatch XML into a List
                    // do a count of Critical, High, and Medium and Low items
                    // return the class of numbers for this
                    _logger.LogInformation("GetNessusPatchSummary({0}) loading Nessus patch data file", systemGroupId);
                    NessusPatchData patchData = NessusPatchLoader.LoadPatchData(sg.rawNessusFile);
                    _logger.LogInformation("GetNessusPatchSummary({0}) querying Nessus patch data file for counts", systemGroupId);

                    NessusPatchCount patchCount = new NessusPatchCount();
                    if (patchData.summary.Count > 0) {
                        patchCount.totalCriticalOpen = patchData.summary.Where(x => x.severity == 4).Count();
                        patchCount.totalHighOpen = patchData.summary.Where(x => x.severity == 3).Count();
                        patchCount.totalMediumOpen = patchData.summary.Where(x => x.severity == 2).Count();
                        patchCount.totalLowOpen = patchData.summary.Where(x => x.severity == 1).Count();;
                        patchCount.totalInfoOpen = patchData.summary.Where(x => x.severity == 0).Count();
                    }

                    _logger.LogInformation("Called GetNessusPatchSummary({0}) successfully", systemGroupId);
                    return Ok(patchCount);
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "GetNessusPatchSummary() Error getting the Nessus patch summary data for system {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called GetNessusPatchSummary() with no system ID");
                return BadRequest(); // no systemGroupId entered
            }
        }
        
        /// <summary>
        /// GET The XLSX export of patch information from the attached Nessus scan file, if there.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <param name="summaryOnly">Show only a summary view? or include the hosts as well?</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the XLSX summary of 
        /// critical, high, medium, low, and info items from the Nessus ACAS Patch scan.
        /// </returns>
        /// <response code="200">Returns the count of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in does not have a valid Nessus file</response>
        [HttpGet("system/{systemGroupId}/exportnessus")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ExportNessusPatchSummary(string systemGroupId, bool summaryOnly)
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling ExportNessusPatchSummary({0})", systemGroupId);
                    SystemGroup sg = new SystemGroup();
                    sg = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                    if (sg == null) {
                        _logger.LogWarning("ExportNessusPatchSummary({0}) an invalid system record");
                        return NotFound();
                    }
                    if (string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogWarning("ExportNessusPatchSummary({0}) system record has no Nessus patch file to use");
                        return NotFound();
                    }
                    // load the NessusPatch XML into a List
                    // do a count of Critical, High, and Medium and Low items
                    // return the class of numbers for this
                    _logger.LogInformation("ExportNessusPatchSummary({0}) loading Nessus patch data file", systemGroupId);
                    NessusPatchData patchData = NessusPatchLoader.LoadPatchData(sg.rawNessusFile);
                    _logger.LogInformation("ExportNessusPatchSummary({0}) querying Nessus patch data file for counts", systemGroupId);
                    
                    // generate the XLSX file from this
                    
                    // starting row number for data
                    uint rowNumber = 7;

                    // create the XLSX in memory and send it out
                    var memory = new MemoryStream();
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                    {
                        // Add a WorkbookPart to the document.
                        WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                        workbookpart.Workbook = new Workbook();
                        
                        // add styles to workbook
                        WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();

                        // Add a WorksheetPart to the WorkbookPart.
                        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        // add stylesheet to use cell formats 1 - 4
                        wbsp.Stylesheet = ExcelStyleSheet.GenerateStylesheet();

                        // generate the column listings we need with custom widths
                        DocumentFormat.OpenXml.Spreadsheet.Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
                        if (lstColumns == null) {
                            lstColumns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 20, CustomWidth = true }); // col A
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 80, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 40, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 30, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 5, Max = 5, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 7, Max = 7, Width = 60, CustomWidth = true });
                            worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                        }

                        // Add Sheets to the Workbook.
                        Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                        // Append a new worksheet and associate it with the workbook.
                        Sheet sheet = new Sheet() { Id = spreadSheet.WorkbookPart.
                            GetIdOfPart(worksheetPart), SheetId = 1, Name = "Nessus-Patch-Summary" };
                        sheets.Append(sheet);
                        // Get the sheetData cell table.
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                        DocumentFormat.OpenXml.Spreadsheet.Cell newCell = null;

                        DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("OpenRMF by Cingulara and Tutela");
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Nessus Patch Summary Export", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"),2);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("System Name", sg.title,3);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Report Name", patchData.reportName,4);
                        sheetData.Append(row);
                        row = MakeNessusSummaryHeaderRows(rowNumber, summaryOnly);
                        sheetData.Append(row);

                        uint styleIndex = 0; // use this for 4, 5, 6, or 7 for status

                        // cycle through the vulnerabilities to export into columns
                        _logger.LogInformation("ExportNessusPatchSummary() grouping the patch information by host");
                        // group all the results by hostname, pluginId, pluginName, Family, and severity and count the totals per hostname
                        List<NessusPatchSummary> reportSummaryFinal = patchData.summary.GroupBy(x => new {x.hostname, x.pluginId, x.pluginName, x.family, x.severity})
                            .Select(g => new NessusPatchSummary {
                                hostname = g.Key.hostname, 
                                pluginId = g.Key.pluginId, 
                                pluginName = g.Key.pluginName, 
                                family = g.Key.family, 
                                severity = g.Key.severity, 
                                total = g.Count()}).ToList();
                        if (summaryOnly) {
                            _logger.LogInformation("ExportNessusPatchSummary() grouping the patch information for summary only (no hosts)");
                            // now sum the total, but get all the other data minus hostname. Count the hostnames used though per plugin grouping
                            reportSummaryFinal = reportSummaryFinal.GroupBy(x => new {x.pluginId, x.pluginName, x.family, x.severity})
                                .Select(g => new NessusPatchSummary {
                                    pluginId = g.Key.pluginId, 
                                    pluginName = g.Key.pluginName, 
                                    family = g.Key.family, 
                                    severity = g.Key.severity, 
                                    hostTotal = g.Count(),
                                    total = g.Sum(z => z.total)}).ToList();
                        }

                        _logger.LogInformation("ExportNessusPatchSummary() making the XLSX Summary for the patch information");
                        // make this go in reverse order of severity, 4 to 0
                        foreach (NessusPatchSummary summary in reportSummaryFinal.OrderBy(y => y.pluginIdSort).OrderByDescending(x => x.severity).ToList()) {
                            // if this is a regular checklist, make sure the filter for VULN ID is checked before we add this to the list
                            // if this is from a compliance listing, only add the VULN IDs from the control to the listing
                            // the VULN has a CCI_REF field where the attribute would be the value in the cciList if it is valid
                            rowNumber++;                            
                            styleIndex = GetPatchScanStatus(summary.severity);
                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", summary.pluginId, styleIndex);
                            // now cycle through the rest of the items
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(summary.pluginName);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(summary.family);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(summary.severityName);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            if (summaryOnly) 
                                newCell.CellValue = new CellValue(summary.hostTotal.ToString());
                            else 
                                newCell.CellValue = new CellValue("1"); // we are doing this by host, so host total is 1
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(summary.total.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = styleIndex;
                            // only print the hostname if not just a summary
                            if (!summaryOnly) {
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(summary.hostname);
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                            }
                            sheetData.Append(row);
                        }

                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        spreadSheet.Close();
                        // set the filename
                        string filename = sg.title + "-NessusScanSummary";
                        memory.Seek(0, SeekOrigin.Begin);
                        _logger.LogInformation("Called ExportNessusPatchSummary({0}) successfully", systemGroupId);
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(filename));
                    }
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "ExportNessusPatchSummary() Error getting the Nessus scan XLSX export for system {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called ExportNessusPatchSummary() with no system ID");
                return BadRequest(); // no systemGroupId entered
            }
        }        


        /// <summary>
        /// GET The XLSX export of the system test plan.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the XLSX test plan file.
        /// </returns>
        /// <response code="200">Returns the count of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in does not have a valid system record</response>
        [HttpGet("system/{systemGroupId}/testplanexport")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ExportSystemTestPlan(string systemGroupId)
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling ExportSystemTestPlan({0})", systemGroupId);
                    SystemGroup sg = new SystemGroup();
                    sg = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                    if (sg == null) {
                        _logger.LogWarning("ExportSystemTestPlan({0}) an invalid system record");
                        return NotFound();
                    }
                    // load the NessusPatch XML into a List
                    // do a count of Critical, High, and Medium and Low items
                    // return the class of numbers for this
                    NessusPatchData patchData = new NessusPatchData();
                    if (!string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogInformation("ExportSystemTestPlan({0}) loading Nessus patch data file", systemGroupId);
                        patchData = NessusPatchLoader.LoadPatchData(sg.rawNessusFile);
                        _logger.LogInformation("ExportSystemTestPlan({0}) Nessus patch data file loaded", systemGroupId);
                    }
                    
                    // generate the XLSX file from this
                    
                    // starting row number for data
                    uint rowNumber = 7;

                    // create the XLSX in memory and send it out
                    var memory = new MemoryStream();
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                    {
                        _logger.LogInformation("ExportSystemTestPlan({0}) setting up XLSX file", systemGroupId);
                        // Add a WorkbookPart to the document.
                        WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                        workbookpart.Workbook = new Workbook();
                        
                        // add styles to workbook
                        WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();

                        // Add a WorksheetPart to the WorkbookPart.
                        WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                        worksheetPart.Worksheet = new Worksheet(new SheetData());

                        // add stylesheet to use cell formats 1 - 4
                        wbsp.Stylesheet = ExcelStyleSheet.GenerateStylesheet();

                        // generate the column listings we need with custom widths
                        DocumentFormat.OpenXml.Spreadsheet.Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
                        if (lstColumns == null) {
                            lstColumns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 40, CustomWidth = true }); // col A
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 30, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 40, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 40, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 5, Max = 5, Width = 80, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 10, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 7, Max = 7, Width = 10, CustomWidth = true }); // col G
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 8, Max = 8, Width = 10, CustomWidth = true }); // col H
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 9, Max = 9, Width = 10, CustomWidth = true }); // col I
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 10, Max = 10, Width = 10, CustomWidth = true }); // col J
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 11, Max = 11, Width = 30, CustomWidth = true }); // col K
                            worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                        }

                        // Add Sheets to the Workbook.
                        Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                        // Append a new worksheet and associate it with the workbook.
                        Sheet sheet = new Sheet() { Id = spreadSheet.WorkbookPart.
                            GetIdOfPart(worksheetPart), SheetId = 1, Name = "System-Test-Plan" };
                        sheets.Append(sheet);
                        // Get the sheetData cell table.
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                        DocumentFormat.OpenXml.Spreadsheet.Cell newCell = null;

                        _logger.LogInformation("ExportSystemTestPlan({0}) setting title XLSX information", systemGroupId);
                        DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("OpenRMF by Cingulara and Tutela");
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("System Test Plan", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"),2);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("System Name", sg.title,3);
                        sheetData.Append(row);
                        row = MakeTestPlanHeaderRows(rowNumber, true);
                        sheetData.Append(row);

                        uint styleIndex = 0; // use this for 4, 5, 6, or 7 for status
                        Score checklistScore;

                        // cycle through the vulnerabilities to export into columns
                        _logger.LogInformation("ExportSystemTestPlan() grouping the patch information by host");
                        _logger.LogInformation("ExportSystemTestPlan() making the XLSX Summary for the patch information");
                        //     styleIndex = GetPatchScanStatus(summary.severity);
                        
                        // get the list of hosts to use
                        _logger.LogInformation("ExportSystemTestPlan({0}) getting Hosts from Nessus patch data file", systemGroupId);
                        List<string> hostnames = patchData.summary.Select(x => x.hostname).Distinct().ToList();
                        int patchCount = 0;
                        int patchTotal = 0;
                        // for each host, cycle through the # of items and print out
                        foreach (string host in hostnames) {
                            _logger.LogInformation("ExportSystemTestPlan({0}) adding Nessus patch data file for {1}", systemGroupId, host);
                            rowNumber++;
                            // reset numbers just in case
                            patchTotal = 0;
                            patchCount = 0;

                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", host, styleIndex);
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(!string.IsNullOrEmpty(sg.nessusFilename)? sg.nessusFilename : "Latest Scan file");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 0;
                            // now cycle through the rest of the items
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 4).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 3).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 2).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 1).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(patchTotal.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 0;
                            sheetData.Append(row);
                        }

                        // generate the checklist files
                        rowNumber = rowNumber + 5;                        
                        _logger.LogInformation("ExportSystemTestPlan({0}) getting checklist data and scores", systemGroupId);
                        row = MakeTestPlanHeaderRows(rowNumber, false);
                        sheetData.Append(row);
                        IEnumerable<Artifact> artifacts;
                        artifacts = await _artifactRepo.GetSystemArtifacts(systemGroupId);
                        foreach (Artifact art in artifacts.OrderBy(x => x.title).OrderBy(y => y.systemTitle).ToList()) {
                            try {
                                _logger.LogInformation("ExportSystemTestPlan({0}) getting checklist data and scores for {1}", systemGroupId, art.title);
                                checklistScore = NATSClient.GetChecklistScore(art.InternalId.ToString());
                            }
                            catch (Exception ex) {
                                _logger.LogWarning(ex, "No score found for artifact {0}", art.InternalId.ToString());
                                checklistScore = new Score();
                            }
                            rowNumber++;

                            _logger.LogInformation("ExportSystemTestPlan({0}) making checklist row for {1}", systemGroupId, art.title);
                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", art.hostName.Trim().ToLower() != ""? art.hostName : "", styleIndex);
                            // now cycle through the rest of the items
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(art.title+ ".ckl");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("0");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalOpen.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 0;
                            sheetData.Append(row);
                        }

                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        _logger.LogInformation("ExportSystemTestPlan({0}) closing the XLSX test plan", systemGroupId);
                        spreadSheet.Close();
                        // set the filename
                        string filename = sg.title + "-SystemTestPlan";
                        memory.Seek(0, SeekOrigin.Begin);
                        _logger.LogInformation("Called ExportSystemTestPlan({0}) successfully", systemGroupId);
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(filename));
                    }
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "ExportSystemTestPlan() Error getting the test plan XLSX export for system {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called ExportSystemTestPlan() with no system ID");
                return BadRequest(); // no systemGroupId entered
            }
        }        
        #endregion

        #region Artifacts and Checklists

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) to return a checklist record . 
        /// </summary>
        /// <param name="id">The ID of the checklist to use</param>
        /// <returns>
        /// HTTP Status showing it was generated or that there is an error. And the checklist full record
        /// with all metadata.
        /// </returns>
        /// <response code="200">Returns the checklist data in CKL/XML format</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("artifact/{id}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetArtifact(string id)
        {
            try {
                _logger.LogInformation("Calling GetArtifact({0})", id);
                Artifact art = new Artifact();
                art = await _artifactRepo.GetArtifact(id);
                if (art == null) {
                    _logger.LogWarning("Called GetArtifact({0}) with an invalid Artifact ID", id);
                    return NotFound();
                }
                art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist.Replace("\t","").Replace(">\n<","><"));
                art.rawChecklist = string.Empty;
                _logger.LogInformation("Called GetArtifact({0}) successfully", id);
                return Ok(art);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "GetArtifact({0}) Error Retrieving Artifact", id);
                return BadRequest();
            }
        }

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) to export a checklist to its native CKL format. 
        /// </summary>
        /// <param name="id">The ID of the checklist to use</param>
        /// <returns>
        /// HTTP Status showing it was generated or that there is an error. And the checklist in a valid 
        /// CKL file downloaded to the user. This file can be used elsewhere as in the DISA Java STIG viewer.
        /// </returns>
        /// <response code="200">Returns the checklist data in CKL/XML format</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("download/{id}")]
        [Authorize(Roles = "Administrator,Editor,Assessor,Reader")]
        public async Task<IActionResult> DownloadChecklist(string id)
        {
            try {
                _logger.LogInformation("Calling DownloadChecklist({0})", id);
                Artifact art = new Artifact();
                art = await _artifactRepo.GetArtifact(id);
                if (art == null) {
                    _logger.LogWarning("Called DownloadChecklist({0}) with an invalid Artifact ID", id);
                    return NotFound();
                }
                _logger.LogInformation("Called DownloadChecklist({0}) successfully", id);
                return Ok(art.rawChecklist);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "DownloadChecklist({0}) Error Retrieving Artifact for Download", id);
                return BadRequest();
            }
        }

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) to export a checklist and the valid vulnerabilities to 
        /// MS Excel format. 
        /// </summary>
        /// <param name="id">The ID of the checklist to use</param>
        /// <param name="nf">Include Not a Finding Vulnerabilities</param>
        /// <param name="open">Include Open Vulnerabilities</param>
        /// <param name="na">Include Not Applicable Vulnerabilities</param>
        /// <param name="nr">Include Not Reviewed Vulnerabilities</param>
        /// <param name="ctrl">Include Vulnerabilities only linked to a specific control</param>
        /// <returns>
        /// HTTP Status showing it was generated or that there is an error. And the checklist with all relevant
        /// vulnerabilities in a valid XLSX file downloaded to the user.
        /// </returns>
        /// <response code="200">Returns the checklist data in XLSX format</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpPost("export/{id}")]
        [Authorize(Roles = "Administrator,Editor,Assessor,Reader")]
        public async Task<IActionResult> ExportChecklist(string id, bool nf, bool open, bool na, bool nr, string ctrl)
        {
            try {
                _logger.LogInformation("Calling ExportChecklist({0}, {1}, {2}, {3}, {4}, {5})", id, nf.ToString(), open.ToString(), na.ToString(), nr.ToString(), ctrl);
                if (!string.IsNullOrEmpty(id)) {
                    Artifact art = new Artifact();
                    art = await _artifactRepo.GetArtifact(id);
                    if (art != null && art.CHECKLIST != null) {
                        List<string> cciList = new List<string>();
                        _logger.LogInformation("ExportChecklist({0}....) formatting the checklist to XML from the raw string format");
                        art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                        _logger.LogInformation("ExportChecklist({0}....) checklist formatted");
                        // starting row number for data
                        uint rowNumber = 10;

                        // create the XLSX in memory and send it out
                        var memory = new MemoryStream();
                        using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                        {
                            // Add a WorkbookPart to the document.
                            WorkbookPart workbookpart = spreadSheet.AddWorkbookPart();
                            workbookpart.Workbook = new Workbook();
                            
                            // add styles to workbook
                            WorkbookStylesPart wbsp = workbookpart.AddNewPart<WorkbookStylesPart>();

                            // Add a WorksheetPart to the WorkbookPart.
                            WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                            worksheetPart.Worksheet = new Worksheet(new SheetData());

                            // add stylesheet to use cell formats 1 - 4
                            wbsp.Stylesheet = ExcelStyleSheet.GenerateStylesheet();

                            DocumentFormat.OpenXml.Spreadsheet.Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
                            if (lstColumns == null) { // generate the column listings we need with custom widths
                                lstColumns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 20, CustomWidth = true }); // col A
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 40, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 5, Max = 5, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 7, Max = 7, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 8, Max = 8, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 9, Max = 9, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 10, Max = 10, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 11, Max = 11, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 12, Max = 12, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 13, Max = 13, Width = 20, CustomWidth = true }); // col M
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 14, Max = 14, Width = 75, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 15, Max = 15, Width = 75, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 16, Max = 16, Width = 75, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 17, Max = 17, Width = 75, CustomWidth = true }); // col Q
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 18, Max = 18, Width = 75, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 19, Max = 19, Width = 75, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 20, Max = 20, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 21, Max = 21, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 22, Max = 22, Width = 30, CustomWidth = true }); // col V
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 23, Max = 23, Width = 50, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 24, Max = 24, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 25, Max = 25, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 26, Max = 26, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 27, Max = 27, Width = 75, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 28, Max = 28, Width = 20, CustomWidth = true }); // col AB
                                worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                            }

                            // Add Sheets to the Workbook.
                            Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                            // Append a new worksheet and associate it with the workbook.
                            Sheet sheet = new Sheet() { Id = spreadSheet.WorkbookPart.
                                GetIdOfPart(worksheetPart), SheetId = 1, Name = "STIG-Checklist" };
                            sheets.Append(sheet);
                            // Get the sheetData cell table.
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = null;

                            DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("OpenRMF by Cingulara and Tutela");
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("System Name", art.systemTitle,2);
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("Checklist Name", art.title,3);
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("Host Name", art.hostName,4);
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("Type", art.stigType,5);
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("Release", art.stigRelease,6);
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("Last Updated", art.updatedOn.Value.ToString("MM/dd/yy hh:mm tt"),7);
                            sheetData.Append(row);
                            row = MakeChecklistHeaderRows(rowNumber);
                            sheetData.Append(row);

                            uint styleIndex = 0; // use this for 4, 5, 6, or 7 for status
                            // if this is from a compliance generated listing link to a checklist, go grab all the CCIs for that control
                            // as we are only exporting through VULN IDs that are related to that CCI
                            if (!string.IsNullOrEmpty(ctrl)) {
                                _logger.LogInformation("ExportChecklist() generating the CCI Listing for {0}", ctrl);
                                cciList = NATSClient.GetCCIListing(ctrl);
                            }

                            // cycle through the vulnerabilities to export into columns                            
                            _logger.LogInformation("ExportChecklist() cycling through all the vulnerabilities");

                            string cciReferences = "";
                            foreach (VULN v in art.CHECKLIST.STIGS.iSTIG.VULN) {
                                // if this is a regular checklist, make sure the filter for VULN ID is checked before we add this to the list
                                // if this is from a compliance listing, only add the VULN IDs from the control to the listing
                                // the VULN has a CCI_REF field where the attribute would be the value in the cciList if it is valid
                                if (ShowVulnerabilityInExcel(v.STATUS, nf, nr, open, na) || 
                                        (!string.IsNullOrEmpty(ctrl) && 
                                        v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "CCI_REF" && cciList.Contains(x.ATTRIBUTE_DATA)).FirstOrDefault() != null))  {
                                    rowNumber++;
                                    styleIndex = GetVulnerabilityStatus(v.STATUS);
                                    // make a new row for this set of items
                                    row = MakeDataRow(rowNumber, "A", v.STIG_DATA[0].ATTRIBUTE_DATA, styleIndex);
                                    // now cycle through the rest of the items
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[1].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[2].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[3].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[4].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[5].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[6].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[7].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[8].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[9].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[10].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[11].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[12].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[13].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "O" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[14].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[15].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[16].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[17].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[18].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[19].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[20].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STIG_DATA[21].ATTRIBUTE_DATA);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.STATUS);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.COMMENTS);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.FINDING_DETAILS);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.SEVERITY_OVERRIDE);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AA" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(v.SEVERITY_JUSTIFICATION);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AB" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);

                                    // collect all CCI references
                                    cciReferences = "";
                                    for (int i = 24; i < v.STIG_DATA.Count; i++) { 
                                        if (v.STIG_DATA[i].VULN_ATTRIBUTE == "CCI_REF") 
                                            cciReferences += v.STIG_DATA[i].ATTRIBUTE_DATA + ", ";
                                    }
                                    // take off the ", " at the end
                                    if (!string.IsNullOrEmpty(cciReferences) && cciReferences.Length > 2)
                                        cciReferences = cciReferences.Substring(0, cciReferences.Length-2);
                                    else 
                                        cciReferences = "";

                                    newCell.CellValue = new CellValue(cciReferences);
                                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                    newCell.StyleIndex = styleIndex;
                                    sheetData.Append(row);
                                }
                            }

                            // Save the new worksheet.
                            workbookpart.Workbook.Save();
                            // Close the document.
                            spreadSheet.Close();
                            // set the filename
                            string filename = art.title;
                            if (!string.IsNullOrEmpty(art.systemTitle) && art.systemTitle.ToLower().Trim() == "none")
                                filename = art.systemTitle.Trim() + "-" + filename; // add the system onto the front
                            // return the file
                            memory.Seek(0, SeekOrigin.Begin);
                            _logger.LogInformation("Called ExportChecklist({0}, {1}, {2}, {3}, {4}, {5}) successfully", id, nf.ToString(), open.ToString(), na.ToString(), nr.ToString(), ctrl);
                            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(filename));
                        }
                    }
                    else {                        
                        _logger.LogWarning("Calling ExportChecklist({0}, {1}, {2}, {3}, {4}, {5}) with an invalid Artifact ID", id, nf.ToString(), open.ToString(), na.ToString(), nr.ToString(), ctrl);
                        return NotFound();
                    }
                }
                else { // did not pass in an id
                    _logger.LogWarning("ExportChecklist() Did not pass in an Artifact ID in the Export of a single checklist.");
                    return BadRequest();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "ExportChecklist({0} Error Retrieving Artifact for Exporting", id);
                return BadRequest();
            }
        } 

        private string CreateXLSXFilename(string title) {
            return title.Trim().Replace(" ", "_") + ".xlsx";
        }

        private bool ShowVulnerabilityInExcel(string status, bool NotAFinding, bool NotReviewed, 
                bool Open, bool NotApplicable) {
            if (status.ToLower() == "not_reviewed" && NotReviewed)
                return true;
            if (status.ToLower() == "open" && Open)
                return true;
            if (status.ToLower() == "not_applicable" && NotApplicable)
                return true;
            if (status.ToLower() == "notafinding" && NotAFinding) 
                return true;

            // catchall 
            return false;
        }

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) to return the list of vulnerability IDs in 
        /// a checklist filtered by the control they are linked to.
        /// </summary>
        /// <param name="id">The ID of the checklist to use</param>
        /// <param name="control">The control to filter the vulnerabilities on</param>
        /// <returns>
        /// HTTP Status showing it was generated or that there is an error. And the list of VULN IDs for
        /// this control on this checklist. Usually called from clicking an individual checklist on the 
        /// Compliance page in the UI.
        /// </returns>
        /// <response code="200">Returns the list of vulnerability IDs for the control</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("{id}/control/{control}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetArtifactVulnIdsByControl(string id, string control)
        {
            try {
                _logger.LogInformation("Calling GetArtifactVulnIdsByControl({0}, {1})", id, control);
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(control)) {
                    _logger.LogInformation("GetArtifactVulnIdsByControl() Invalid Artifact Id {0} or Control {1}", id, control);
                    Artifact art = new Artifact();
                    art = await _artifactRepo.GetArtifact(id);
                    art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                    // go get the list of CCIs to look for
                    List<string> cciList = NATSClient.GetCCIListing(control);
                    if (cciList != null) {
                        List<string> vulnIds = new List<string>();
                        // for each string in the listing, find all VULN Ids where you have the CCI listed
                        foreach (VULN v in art.CHECKLIST.STIGS.iSTIG.VULN){
                            // see if the CCI_REF is in the 
                            if (v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "CCI_REF" && cciList.Contains(x.ATTRIBUTE_DATA)).FirstOrDefault() != null) {
                                // the CCI is in this VULN so pull the VULN_ID and add to the list
                                // the Vuln_Num is required so it will be there, otherwise this checklist is invalid
                                vulnIds.Add(v.STIG_DATA.Where(y => y.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault().ATTRIBUTE_DATA); // add the V-xxxx number
                            }
                        }
                        _logger.LogInformation("Called GetArtifactVulnIdsByControl({0}, {1}) successfully", id, control);
                        return Ok(vulnIds.Distinct().OrderBy(z => z).ToList());
                    }
                    else {
                        _logger.LogWarning("Called GetArtifactVulnIdsByControl({0}, {1}) but returned an empty CCI Listing", id, control);
                        return NotFound();
                    }
                }
                else {
                    // log the values passed in
                    _logger.LogWarning("GetArtifactVulnIdsByControl() Invalid Artifact Id {0} or Control {1}", 
                        !string.IsNullOrEmpty(id)? id : "null", !string.IsNullOrEmpty(control)? control : "null");
                    return NotFound();    
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Called GetArtifactVulnIdsByControl({0}, {1}) Error Retrieving Artifact", id, control);
                return BadRequest();
            }
        }

        #endregion

        #region XLSX Formatting
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeTitleRow(string title) {
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = 1 };
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A1"};
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(title.Trim());
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 2;
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeChecklistInfoRow(string title, string value, uint rowindex) {
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString()};
            row.InsertBefore(newCell, refCell);
            string cellValue = title + ": ";
            if (!String.IsNullOrEmpty(value))
                cellValue += value;
            else 
                cellValue += "N/A";
            newCell.CellValue = new CellValue(cellValue);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeChecklistHeaderRows(uint rowindex) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            uint styleIndex = 3;

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Vuln ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Group Title");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Rule ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("STIG ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Rule Title");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Discussion");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("IA Controls");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Check Content");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Fix Text");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("False Positives");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("False Negatives");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Documentable");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Mitigations");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "O" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Potential Impact");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Third Party Tools");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Mitigation Control");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Responsibility");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity Override Guidance");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Check Content Reference");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Classification");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("STIG");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Status");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Comments");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Finding Details");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity Override");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AA" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity Override Justification");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AB" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("CCI");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeNessusSummaryHeaderRows(uint rowindex, bool summaryOnly) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            uint styleIndex = 3;

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Plugin ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Plugin Name");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Family");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Host Total");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Total");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            if (!summaryOnly) {
                // next column
                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowindex.ToString() };
                row.InsertBefore(newCell, refCell);
                newCell.CellValue = new CellValue("Host");
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                newCell.StyleIndex = styleIndex;
            }
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeTestPlanHeaderRows(uint rowindex, bool credentialed) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            uint styleIndex = 3;

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Host Name");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("IP Address");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Group Name");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Operating System");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Filename");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("CAT I");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("CAT II");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("CAT III");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("CAT IV");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Total");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            if (credentialed) {
                // next column
                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowindex.ToString() };
                row.InsertBefore(newCell, refCell);
                newCell.CellValue = new CellValue("Credentialed");
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                newCell.StyleIndex = styleIndex;
            }
            return row;
        }

        private DocumentFormat.OpenXml.Spreadsheet.Row MakeDataRow(uint rowNumber, string cellReference, string value, uint styleIndex) {
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowNumber };
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = cellReference + rowNumber.ToString()};
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(value);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeChecklistListingHeaderRows(uint rowindex) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            uint styleIndex = 3;

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("System");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Title");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Not a Finding");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 10;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Not Applicable");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 9;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Open");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 8;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Not Reviewed");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 11;
            return row;
        }
        private uint GetVulnerabilityStatus(string status) {
            // open = 4, N/A = 5, NotAFinding = 6, Not Reviewed = 7
            if (status.ToLower() == "not_reviewed")
                return 7U;
            if (status.ToLower() == "open")
                return 4U;
            if (status.ToLower() == "not_applicable")
                return 5U;
            // catch all
            return 6U;
        }
        private uint GetPatchScanStatus(int severity) {
            // critical or high = 3 or 4, medium = 2, low = 1, informational = 0
            if (severity > 2)
                return 4U;
            // catch all informational
            return 7U;
        }
        #endregion

        /******************************************
        * Dashboard Specific API calls
        ******************************************/
        #region Dashboard APIs

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) dashboard to return the number of 
        /// total artifacts/checklists within the database.
        /// </summary>
        /// <returns>
        /// HTTP Status showing it was searched correctly and the count of the artifacts. Or that there is an error.
        /// </returns>
        /// <response code="200">Returns the list of checklists and their count</response>
        /// <response code="400">If the item did not search correctly</response>
        [HttpGet("count/artifacts")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> CountArtifacts()
        {
            try {
                _logger.LogInformation("Calling CountArtifacts()");
                long result = await _artifactRepo.CountChecklists();
                _logger.LogInformation("Called CountArtifacts() successfully for {0} items", result.ToString());
                return Ok(result);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "CountArtifacts() Error Retrieving Artifact Count in MongoDB");
                return BadRequest();
            }
        }

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) dashboard to return the number of 
        /// systems within the database.
        /// </summary>
        /// <returns>
        /// HTTP Status showing it was searched correctly and the count of the systems. Or that there is an error.
        /// </returns>
        /// <response code="200">Returns the list of checklist types and their count</response>
        /// <response code="400">If the item did not search correctly</response>
        [HttpGet("count/systems")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> CountSystems()
        {
            try {
                _logger.LogInformation("Calling CountSystems()");
                long result = await _systemGroupRepo.CountSystems();
                _logger.LogInformation("Called CountSystems() successfully for {0} items", result.ToString());
                return Ok(result);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "CountSystems() Error Retrieving System Count in MongoDB");
                return BadRequest();
            }
        }

        /// <summary>
        /// GET Called from the OpenRMF UI (or external access) to return a count of each type of checklist. 
        /// If you pass in the system Id it does this per system.
        /// </summary>
        /// <param name="system">The ID of the system for generating the count</param>
        /// <returns>
        /// HTTP Status showing it was searched correctly and the count per STIG type or that there is an error.
        /// </returns>
        /// <response code="200">Returns the list of checklist types and their count</response>
        /// <response code="400">If the item did not search correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("counttype")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetCountByType(string system)
        {
            try {
                _logger.LogInformation("Calling GetCountByType('{0}')", system);
                IEnumerable<Object> artifacts;
                artifacts = await _artifactRepo.GetCountByType(system);
                _logger.LogInformation("Called GetCountByType('{0}')", system);
                if (artifacts == null) {                    
                    _logger.LogWarning("Calling GetCountByType('{0}') returned null", system);
                    NotFound();
                }
                return Ok(artifacts);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "GetCountByType() Error getting the counts by type for the Reports page");
                return BadRequest();
            }
        }
        #endregion

        /******************************************
        * Reports Specific API calls
        ******************************************/
        #region Report APIs

        /// <summary>
        /// GET The raw data of a Nessus file for the reports.
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the raw data of 
        /// the Nessus ACAS Patch scan.
        /// </returns>
        /// <response code="200">Returns the count of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in does not have a valid Nessus file</response>
        [HttpGet("report/nessus/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetNessusPatchDataForReport(string systemGroupId)
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling GetNessusPatchDataForReport({0})", systemGroupId);
                    SystemGroup sg = new SystemGroup();
                    sg = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                    if (sg == null) {
                        _logger.LogWarning("GetNessusPatchDataForReport({0}) an invalid system record");
                        return NotFound();
                    }
                    if (string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogWarning("GetNessusPatchDataForReport({0}) system record has no Nessus patch file to use");
                        return NotFound();
                    }
                    // load the NessusPatch XML into a List
                    // do a count of Critical, High, and Medium and Low items
                    // return the class of numbers for this
                    _logger.LogInformation("GetNessusPatchDataForReport({0}) loading Nessus patch data file", systemGroupId);
                    NessusPatchData patchData = NessusPatchLoader.LoadPatchData(sg.rawNessusFile);
                    _logger.LogInformation("GetNessusPatchDataForReport({0}) querying Nessus patch data file for report", systemGroupId);

                    _logger.LogInformation("Called GetNessusPatchDataForReport({0}) successfully", systemGroupId);
                    return Ok(patchData);
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "GetNessusPatchDataForReport() Error getting data for the Nessus Data report for system {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called GetNessusPatchDataForReport() with no system ID");
                return BadRequest(); // no systemGroupId entered
            }
        }
        
        #endregion
    }
}
