﻿// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using openrmf_read_api.Classes;
using openrmf_read_api.Models;
using System.IO;
using System.IO.Compression;
using Microsoft.AspNetCore.Authorization;
using Microsoft.Extensions.Logging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using openrmf_read_api.Data;

namespace openrmf_read_api.Controllers
{
    [Route("/artifact")]
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
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <param name="naf">True/False include Not a Finding vulnerabilities</param>
        /// <param name="open">True/False include Open vulnerabilities</param>
        /// <param name="na">True/False include Not Applicable vulnerabilities</param>
        /// <param name="nr">True/False include Not Reviewed vulnerabilities</param>
        /// <param name="cat1">True/False include CAT 1 / High vulnerabilities</param>
        /// <param name="cat2">True/False include CAT 2 / Medium vulnerabilities</param>
        /// <param name="cat3">True/False include CAT 3 / Low vulnerabilities</param>
        /// <param name="hostname">The hostname of the checklist to filter on</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the list of system records 
        /// exported to an XLSX file to download.
        /// </returns>
        /// <response code="200">Returns the Artifact List of records for the passed in system in XLSX format</response>
        /// <response code="400">If the item did not query correctly</response>
        [HttpGet("system/export/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Reader,Assessor")]
        public async Task<IActionResult> ExportChecklistListing(string systemGroupId, bool naf = true, bool open = true, bool na = true, 
            bool nr = true, bool cat1 = true, bool cat2 = true, bool cat3 = true, string hostname = "")
        {
            try {
                _logger.LogInformation("Calling ExportChecklistListing({0})", systemGroupId);
                IEnumerable<Artifact> artifacts;
                // if they pass in a system, get all for that system
                if (string.IsNullOrEmpty(systemGroupId))
                {
                    _logger.LogWarning("ExportChecklistListing() Getting a listing of all checklists to export to XLSX");
                    return BadRequest("You must specify a system to export");
                }
                _logger.LogInformation(string.Format("ExportChecklistListing() Getting a listing of all {0} checklists to export to XLSX", systemGroupId));
                artifacts = await _artifactRepo.GetSystemArtifacts(systemGroupId);
                if (!naf || !open || !na || !nr || !cat1 || !cat2 || !cat3) {
                    // used to store the allowed checklists
                    List<string> allowedChecklists = new List<string>();
                    bool allowChecklist = false;
                    List<Score> scoreListing = NATSClient.GetSystemScores(systemGroupId);
                    // get the list of scores for all of them
                    if (scoreListing != null && scoreListing.Count > 0) {
                        // see for each if they should be included
                        foreach (Score s in scoreListing) {
                            allowChecklist = false;
                            // now we can check status and boolean
                            if (s.totalCat1Open > 0 && cat1 && open)
                                allowChecklist = true;
                            else if (s.totalCat2Open > 0 && cat2 && open)
                                allowChecklist = true;
                            else if (s.totalCat3Open > 0 && cat3 && open)
                                allowChecklist = true;
                            else if (s.totalCat1NotReviewed > 0 && cat1 && nr)
                                allowChecklist = true;
                            else if (s.totalCat2NotReviewed > 0 && cat2 && nr)
                                allowChecklist = true;
                            else if (s.totalCat3NotReviewed > 0 && cat3 && nr)
                                allowChecklist = true;
                            else if (s.totalCat1NotApplicable > 0 && cat1 && na)
                                allowChecklist = true;
                            else if (s.totalCat2NotApplicable > 0 && cat2 && na)
                                allowChecklist = true;
                            else if (s.totalCat3NotApplicable > 0 && cat3 && na)
                                allowChecklist = true;
                            else if (s.totalCat1NotAFinding > 0 && cat1 && naf)
                                allowChecklist = true;
                            else if (s.totalCat2NotAFinding > 0 && cat2 && naf)
                                allowChecklist = true;
                            else if (s.totalCat3NotAFinding > 0 && cat3 && naf)
                                allowChecklist = true;
                            else
                                allowChecklist = false; // only if no severity is checked, which is not smart

                            if (allowChecklist) // then add it, we will do a distinct later
                                allowedChecklists.Add(s.artifactId);
                        }
                        if (allowedChecklists.Count == 0)
                            artifacts = new List<Artifact>(); // there are not allowed, so empty it
                        else {
                            // otherwise remove from the systemChecklists before going on
                            artifacts = artifacts.Where(o => allowedChecklists.Contains(o.InternalId.ToString()));
                        }
                    }
                } 
                // check for hostname being used
                if (!string.IsNullOrEmpty(hostname)) {
                    artifacts = artifacts.Where(z => z.hostName.Contains(hostname));
                }

                // now use the listing
                if (artifacts != null && artifacts.Count() > 0) {
                    // starting row number for data
                    uint rowNumber = 5;

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
                            row = MakeXLSXInfoRow("Checklist Listing", artifacts.FirstOrDefault().systemTitle, 2);
                        } else {
                            row = MakeXLSXInfoRow("Checklist Listing", systemGroupId, 2);
                        }
                        sheetData.Append(row);
                        row = MakeXLSXInfoRow("Printed Date", DateTime.Now.ToString("MM/dd/yy hh:mm"),3);
                        sheetData.Append(row);
                        row = MakeChecklistListingHeaderRows(rowNumber);
                        sheetData.Append(row);

                        MergeCells mergeCells = new MergeCells();
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:F1") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:F2") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:F3") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A4:F4") });
                        
                        uint styleIndex = 18; // use this for 4, 5, 6, or 7 for status
                        Score checklistScore;

                        // cycle through the checklists and grab the score for each individually
                        _logger.LogInformation("ExportChecklistListing({0}) cycling through checklists to list", systemGroupId);
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
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalOpen.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalNotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalNotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
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
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1NotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1NotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
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
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 14;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2NotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2NotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
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
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 15;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3NotAFinding.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 10;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3NotApplicable.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 9;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3NotReviewed.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            sheetData.Append(row);
                        }

                        // save the merged cells
                        worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        spreadSheet.Dispose();
                        memory.Seek(0, SeekOrigin.Begin);
                        _logger.LogInformation("Called ExportChecklistListing({0}) successfully", systemGroupId);
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ChecklistListing.xlsx");
                    }
                }
                else {
                    _logger.LogInformation("Calling ExportChecklistListing({0}) but had no checklists to show", systemGroupId);
                    return NotFound();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "ExportChecklistListing({0}) Error Retrieving Artifacts for Exporting", systemGroupId);
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
        /// <param name="naf">True/False include Not a Finding vulnerabilities</param>
        /// <param name="open">True/False include Open vulnerabilities</param>
        /// <param name="na">True/False include Not Applicable vulnerabilities</param>
        /// <param name="nr">True/False include Not Reviewed vulnerabilities</param>
        /// <param name="cat1">True/False include CAT 1 / High vulnerabilities</param>
        /// <param name="cat2">True/False include CAT 2 / Medium vulnerabilities</param>
        /// <param name="cat3">True/False include CAT 3 / Low vulnerabilities</param>
        /// <param name="hostname">The hostname to filter the list of checklists on</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the list of checklists records.
        /// </returns>
        /// <response code="200">Returns the Artifact List of records</response>
        /// <response code="400">If the item did not query correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("systems/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ListArtifactsBySystem(string systemGroupId, bool naf = true, bool open = true, bool na = true, 
            bool nr = true, bool cat1 = true, bool cat2 = true, bool cat3 = true, string hostname="")
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling ListArtifactsBySystem({0})", systemGroupId);
                    // store the list of all system checklists
                    IEnumerable<Artifact> systemChecklists;
                    // used to store the allowed checklists
                    List<string> allowedChecklists = new List<string>();
                    bool allowChecklist = false;
                    // get the checklists
                    systemChecklists = await _artifactRepo.GetSystemArtifacts(systemGroupId);
                    // If we are filtering at all then use the filter
                    // Otherwise send the whole list
                    if (!naf || !open || !na || !nr || !cat1 || !cat2 || !cat3) {
                        List<Score> scoreListing = NATSClient.GetSystemScores(systemGroupId);
                        // get the list of scores for all of them
                        if (scoreListing != null && scoreListing.Count > 0) {
                            // see for each if they should be included
                            foreach (Score s in scoreListing) {
                                allowChecklist = false;
                                // now we can check status and boolean
                                if (s.totalCat1Open > 0 && cat1 && open)
                                    allowChecklist = true;
                                else if (s.totalCat2Open > 0 && cat2 && open)
                                    allowChecklist = true;
                                else if (s.totalCat3Open > 0 && cat3 && open)
                                    allowChecklist = true;
                                else if (s.totalCat1NotReviewed > 0 && cat1 && nr)
                                    allowChecklist = true;
                                else if (s.totalCat2NotReviewed > 0 && cat2 && nr)
                                    allowChecklist = true;
                                else if (s.totalCat3NotReviewed > 0 && cat3 && nr)
                                    allowChecklist = true;
                                else if (s.totalCat1NotApplicable > 0 && cat1 && na)
                                    allowChecklist = true;
                                else if (s.totalCat2NotApplicable > 0 && cat2 && na)
                                    allowChecklist = true;
                                else if (s.totalCat3NotApplicable > 0 && cat3 && na)
                                    allowChecklist = true;
                                else if (s.totalCat1NotAFinding > 0 && cat1 && naf)
                                    allowChecklist = true;
                                else if (s.totalCat2NotAFinding > 0 && cat2 && naf)
                                    allowChecklist = true;
                                else if (s.totalCat3NotAFinding > 0 && cat3 && naf)
                                    allowChecklist = true;
                                else
                                    allowChecklist = false; // only if no severity is checked, which is not smart

                                if (allowChecklist) // then add it, we will do a distinct later
                                    allowedChecklists.Add(s.artifactId);
                            }
                            if (allowedChecklists.Count == 0)
                                systemChecklists = new List<Artifact>(); // there are not allowed, so empty it
                            else {
                                // otherwise remove from the systemChecklists before going on
                                systemChecklists = systemChecklists.Where(o => allowedChecklists.Contains(o.InternalId.ToString()));
                            }
                        }
                    } 
                    
                    // check for hostname being used
                    if (!string.IsNullOrEmpty(hostname)) {
                        systemChecklists = systemChecklists.Where(z => z.hostName.ToLower().Contains(hostname.ToLower()));
                    }

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
        [HttpGet("system/{systemGroupId}/downloadnessus")]
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
                        _logger.LogWarning("GetNessusPatchSummary({0}) an invalid system record", systemGroupId);
                        return NotFound();
                    }
                    if (string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogWarning("GetNessusPatchSummary({0}) system record has no Nessus patch file to use", systemGroupId);
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
                        _logger.LogWarning("ExportNessusPatchSummary({0}) an invalid system record", systemGroupId);
                        return NotFound();
                    }
                    if (string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogWarning("ExportNessusPatchSummary({0}) system record has no Nessus patch file to use", systemGroupId);
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

                        // generate the column listings we need with custom widths
                        DocumentFormat.OpenXml.Spreadsheet.Columns lstColumns = worksheetPart.Worksheet.GetFirstChild<DocumentFormat.OpenXml.Spreadsheet.Columns>();
                        if (lstColumns == null) {
                            lstColumns = new DocumentFormat.OpenXml.Spreadsheet.Columns();
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 20, CustomWidth = true }); // col A
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 80, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 40, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 5, Max = 5, Width = 15, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 15, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 7, Max = 7, Width = 60, CustomWidth = true }); // col G host if shown
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
                        row = MakeXLSXInfoRow("Nessus Patch Summary Export", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"),2);
                        sheetData.Append(row);
                        row = MakeXLSXInfoRow("System Name", sg.title,3);
                        sheetData.Append(row);
                        row = MakeXLSXInfoRow("Report Name", patchData.reportName,4);
                        sheetData.Append(row);
                        row = MakeNessusSummaryHeaderRows(rowNumber, summaryOnly);
                        sheetData.Append(row);

                        MergeCells mergeCells = new MergeCells();
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:F1") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:F2") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:F3") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A4:F4") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A5:F5") });
                        
                        uint styleIndex = 18; // use this for 4, 5, 6, or 7 for status

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

                        // save the merged cells
                        worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                        
                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        spreadSheet.Dispose();
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
                        _logger.LogWarning("ExportSystemTestPlan({0}) an invalid system record", systemGroupId);
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
                    uint rowNumber = 5;

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
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 30, CustomWidth = true }); // col A
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 30, CustomWidth = true });
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
                        row = MakeXLSXInfoRow("System Test Plan Summary", DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"),2);
                        sheetData.Append(row);
                        row = MakeXLSXInfoRow("System Name", sg.title,3);
                        sheetData.Append(row);
                        row = MakeTestPlanHeaderRows(rowNumber, true);
                        sheetData.Append(row);

                        MergeCells mergeCells = new MergeCells();
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:E1") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:E2") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:E3") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A4:E4") });

                        uint styleIndex = 18; // use this for 4, 5, 6, or 7 for status
                        Score checklistScore;

                        //     styleIndex = GetPatchScanStatus(summary.severity);
                        
                        // get the list of hosts to use
                        _logger.LogInformation("ExportSystemTestPlan({0}) getting Hosts from Nessus patch data file", systemGroupId);
                        List<string> hostnames = patchData.summary.Select(x => x.hostname).Distinct().ToList();
                        int patchCount = 0;
                        int patchTotal = 0;
                        string ipAddress = "";
                        // for each host, cycle through the # of items and print out
                        foreach (string host in hostnames) {
                            _logger.LogInformation("ExportSystemTestPlan({0}) adding Nessus patch data file for {1}", systemGroupId, host);
                            rowNumber++;
                            // reset numbers just in case
                            patchTotal = 0;
                            patchCount = 0;

                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", host, styleIndex);
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(patchData.summary.Where(x => x.hostname == host).Select(y => y.ipAddress).FirstOrDefault());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(patchData.summary.Where(x => x.hostname == host).Select(y => y.operatingSystem).FirstOrDefault());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(!string.IsNullOrEmpty(sg.nessusFilename)? sg.nessusFilename : "Latest Scan file");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            // now cycle through the rest of the items
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity > 2).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 2).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 14;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 1).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 15;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            patchCount = patchData.summary.Where(x => x.hostname == host && x.severity == 0).Count();
                            patchTotal += patchCount;
                            newCell.CellValue = new CellValue(patchCount.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(patchTotal.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 11;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            if (patchData.summary.Where(x => x.hostname == host).Select(y => y.credentialed).FirstOrDefault())
                                newCell.CellValue = new CellValue("Yes");
                            else
                                newCell.CellValue = new CellValue("No");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
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
                            art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
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
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            ipAddress = art.CHECKLIST.ASSET.HOST_IP;
                            if (!string.IsNullOrEmpty(ipAddress)) {
                                ipAddress = NessusPatchLoader.SanitizeHostname(ipAddress);
                            } else
                                ipAddress = "";
                            newCell.CellValue = new CellValue(ipAddress);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(art.title+ ".ckl");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat1Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 8;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat2Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 14;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalCat3Open.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = 15;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("0");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(checklistScore.totalOpen.ToString());
                            newCell.DataType = new EnumValue<CellValues>(CellValues.Number);
                            newCell.StyleIndex = styleIndex;
                            sheetData.Append(row);
                        }

                        // save the merged cells
                        worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        _logger.LogInformation("ExportSystemTestPlan({0}) closing the XLSX test plan", systemGroupId);
                        spreadSheet.Dispose();
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
        [HttpGet("system/{systemGroupId}/poamexport")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ExportSystemPOAM(string systemGroupId)
        {
            if (!string.IsNullOrEmpty(systemGroupId)) {
                try {
                    _logger.LogInformation("Calling ExportSystemPOAM({0})", systemGroupId);
                    SystemGroup sg = new SystemGroup();
                    sg = await _systemGroupRepo.GetSystemGroup(systemGroupId);
                    if (sg == null) {
                        _logger.LogWarning("ExportSystemPOAM({0}) an invalid system record", systemGroupId);
                        return NotFound();
                    }
                    // load the NessusPatch XML into a List
                    // do a count of Critical, High, and Medium and Low items
                    // return the class of numbers for this
                    NessusPatchData patchData = new NessusPatchData();
                    if (!string.IsNullOrEmpty(sg.rawNessusFile)) {
                        _logger.LogInformation("ExportSystemPOAM({0}) loading Nessus patch data file", systemGroupId);
                        patchData = NessusPatchLoader.LoadPatchData(sg.rawNessusFile);
                        _logger.LogInformation("ExportSystemPOAM({0}) Nessus patch data file loaded", systemGroupId);
                    }
                    
                    // starting row number for data
                    uint rowNumber = 11;

                    // create the XLSX in memory and send it out
                    var memory = new MemoryStream();
                    using (SpreadsheetDocument spreadSheet = SpreadsheetDocument.Create(memory, SpreadsheetDocumentType.Workbook))
                    {
                        _logger.LogInformation("ExportSystemPOAM({0}) setting up XLSX file", systemGroupId);
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
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 2, Width = 20, CustomWidth = true }); // col A B
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 15, CustomWidth = true }); 
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 5, Width = 10, CustomWidth = true }); // D E
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 20, CustomWidth = true }); // F
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 7, Max = 7, Width = 10, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 8, Max = 8, Width = 15, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 9, Max = 9, Width = 20, CustomWidth = true }); // col I
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 10, Max = 10, Width = 10, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 11, Max = 11, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 12, Max = 12, Width = 15, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 13, Max = 13, Width = 10, CustomWidth = true }); // col M
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 14, Max = 14, Width = 20, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 15, Max = 15, Width = 15, CustomWidth = true }); // col O
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 16, Max = 16, Width = 25, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 17, Max = 17, Width = 20, CustomWidth = true }); // col Q
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 18, Max = 18, Width = 30, CustomWidth = true }); // col R
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 19, Max = 19, Width = 20, CustomWidth = true }); // col S
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 20, Max = 20, Width = 40, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 21, Max = 21, Width = 20, CustomWidth = true }); // col U
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 22, Max = 22, Width = 25, CustomWidth = true });
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 23, Max = 26, Width = 15, CustomWidth = true }); // col W - Z
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 27, Max = 27, Width = 25, CustomWidth = true }); // col AA
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 28, Max = 28, Width = 20, CustomWidth = true }); 
                            lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 29, Max = 29, Width = 25, CustomWidth = true }); // col AC
                            worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                        }

                        // Add Sheets to the Workbook.
                        Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                        // a placeholder for merged cells later
                        MergeCells mergeCells = new MergeCells();

                        // Append a new worksheet and associate it with the workbook.
                        Sheet sheet = new Sheet() { Id = spreadSheet.WorkbookPart.
                            GetIdOfPart(worksheetPart), SheetId = 1, Name = "System-POAM" };
                        sheets.Append(sheet);
                        // Get the sheetData cell table.
                        SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                        DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                        DocumentFormat.OpenXml.Spreadsheet.Cell newCell = null;
                        var claim = this.User.Claims.Where(x => x.Type == System.Security.Claims.ClaimTypes.NameIdentifier).FirstOrDefault();
                        string strClaimName = "Unknown";
                        if (claim != null) {
                            strClaimName = claim.Value; // userid
                            if (claim.Subject.Claims.Where(x => x.Type == "name").FirstOrDefault() != null) 
                                strClaimName = claim.Subject.Claims.Where(x => x.Type == "name").FirstOrDefault().Value;
                            else if (claim.Subject.Claims.Where(x => x.Type == "preferred_username").FirstOrDefault() != null)
                                strClaimName = claim.Subject.Claims.Where(x => x.Type == "preferred_username").FirstOrDefault().Value;
                            else if (claim.Subject.Claims.Where(x => x.Type.Contains("emailaddress")).FirstOrDefault() != null)
                                strClaimName = claim.Subject.Claims.Where(x => x.Type.Contains("emailaddress")).FirstOrDefault().Value;
                        }
                        _logger.LogInformation("ExportSystemPOAM({0}) setting title XLSX information", systemGroupId);
                        // DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("OpenRMF by Cingulara and Tutela");
                        // sheetData.Append(row);
                        // row = MakePOAMInfoRow("Date Created:",DateTime.Now.ToString("MM/dd/yyyy hh:mm:ss tt"),"System Type:","","OMB Project ID:","",2);
                        // sheetData.Append(row);
                        // row = MakePOAMInfoRow("Created By:",strClaimName,"","","","",3);
                        // sheetData.Append(row);
                        // row = MakePOAMInfoRow("DoD Component:","","POC Name:","","","",4);
                        // sheetData.Append(row);
                        // row = MakePOAMInfoRow("System Project Name:", sg.title, "POC Phone:", "", "Security Costs:","",5);
                        // sheetData.Append(row);
                        // row = MakePOAMInfoRow("DoD IT Registration Number:","","POC Email:","","","",6);
                        // sheetData.Append(row);
                        // row = MakePOAMHeaderRows(rowNumber, true);
                        // sheetData.Append(row);
                        DocumentFormat.OpenXml.Spreadsheet.Row row = MakeDataRow(1, "A", "CONTROLLED UNCLASSIFIED INFORMATION", 20);
                        sheetData.Append(row);
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:AC1") });
                        row = MakeDataRow(2, "A", "Plan of Action and Milestones (POA&M)", 20);
                        sheetData.Append(row);
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:AC2") });

                        row = MakePOAMInfoRow("Date Initiated:", "","System Type:","","OMB Project ID:","",3);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("", "","","","","",4);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("Date Last Updated:", "","","","","",5);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("", "","","","","",6);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("", "","","","","",7);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("DoD Component:", "","POC Name:","","","",8);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("System/Project Name:", sg.title,"POC Phone:","","Security Costs:","",9);
                        sheetData.Append(row);
                        row = MakePOAMInfoRow("System Identification:", "","POC Email:","","","",10);
                        sheetData.Append(row);
                        
                        // merge the data fields up top
                        // first section
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:A4") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B3:D4") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("E3:G4") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("H3:J4") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("K3:L7") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("M3:N7") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("O3:AC7") });
                        // second row sections
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A5:A7") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B5:D7") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("E5:G7") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("H5:J7") });
                        // third row sections
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B8:D8") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("E8:G8") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("H8:J8") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("K8:N8") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("O8:AC8") });

                        // fourth row sections
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B9:D9") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("E9:G9") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("H9:J9") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("K9:L9") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("M9:N9") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("O9:AC9") });

                        // fifth row sections
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("B10:D10") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("E10:G10") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("H10:J10") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("K10:N10") });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("O10:AC10") });

                        row = MakeHeaderRows(new List<string>(){"Vulnerability Description", "", "Security Control Acronym",
                            "AP Acronym", "", "Office / Org", "Security Checks", "", 
                            "Resources Required", "Scheduled Completion Date", "", "", "", "Milestones with Completion Date", "", 
                            "Milestone Changes", "Artifacts", "Source Identifying Vulnerability", 
                            "Status", "Comments", "Raw Severity", "Mitigations","Severity","Relevance of Threat", "Likelihood",
                            "Impact","Impact Description","Residual Risk Level","Recommendations" }, rowNumber, 17);
                        sheetData.Append(row);
                        // column header sections
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("A" + rowNumber + ":B" + rowNumber) });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("D" + rowNumber + ":E" + rowNumber) });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("G" + rowNumber + ":H" + rowNumber) });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("J" + rowNumber + ":M" + rowNumber) });
                        mergeCells.Append(new MergeCell() { Reference = new StringValue("N" + rowNumber + ":O" + rowNumber) });

                        uint styleIndex = 18; // use this for 4, 5, 6, or 7 for status
                        
                        // get the list of hosts to use
                        _logger.LogInformation("ExportSystemPOAM({0}) getting Hosts from Nessus patch data file", systemGroupId);
                        string reportName = patchData.reportName;
                        string severityName = "";
                        // for each patch issue, cycle through and put a row -- order by severity and then hostname if in the correct order!!
                        // only severity 1 - 4 items
                        patchData.summary = patchData.summary.Where(z => z.severity > 0).OrderBy(y => y.hostname).OrderByDescending(x => x.severity).ToList();
                        foreach (NessusPatchSummary p in patchData.summary) {
                            _logger.LogInformation("ExportSystemPOAM({0}) adding Nessus patch summary row for {1}", systemGroupId, p.pluginId);
                            rowNumber++;

                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", "Title:\n" + p.pluginName + "\n\nDescription:\n" + p.description, styleIndex);

                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(p.pluginId);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "0" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            if (!string.IsNullOrEmpty(p.scanVersion))
                                newCell.CellValue = new CellValue("Assured Compliance Assessment Solution (ACAS) Nessus Scanner :: " + p.scanVersion);
                            else 
                                newCell.CellValue = new CellValue("Assured Compliance Assessment Solution (ACAS) Nessus Scanner :: (version)");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("Ongoing");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            if (p.severity == 4) severityName = "Critical";
                            else if (p.severity == 3) severityName = "High";
                            else if (p.severity == 2) severityName = "Medium";
                            else if (p.severity == 1) severityName = "Low";
                            newCell.CellValue = new CellValue(severityName);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            // color code the information
                            newCell.StyleIndex = GetPatchScanStatus(p.severity);
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AA" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AB" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AC" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue("");
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = styleIndex;

                            // add the row to the sheet now
                            sheetData.Append(row);
                                
                            // merge appropriate data cells    
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A" + rowNumber + ":B" + rowNumber) });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("D" + rowNumber + ":E" + rowNumber) });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("G" + rowNumber + ":H" + rowNumber) });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("J" + rowNumber + ":M" + rowNumber) });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("N" + rowNumber + ":O" + rowNumber) });
                        }

                        // get the list of checklists
                        _logger.LogInformation("ExportSystemPOAM({0}) getting all system checklists.", systemGroupId);
                        IEnumerable<Artifact> checklists = await _artifactRepo.GetSystemArtifacts(systemGroupId);
                        if (checklists != null) {
                            // put all VULNs in here to order later
                            List<VulnerabilityReport> vulnerabilities = new List<VulnerabilityReport>();
                            _logger.LogInformation("ExportSystemPOAM({0}) reading the list of all CCI to NIST controls.", systemGroupId);
                            List<CciItem> cciList = NATSClient.GetCCIListing();
                            _logger.LogInformation("ExportSystemPOAM({0}) cycling through the checklists to get all Open or N/R vulnerabilities.", systemGroupId);
                            VulnerabilityReport vulnReport;
                            string hostname;
                            string checklistType;
                            string checklistVersion;
                            string checklistRelease;
                            string cciReferences;
                            foreach (Artifact a in checklists) {
                                // for each checklist, cycle through the vulnerabilities that are Open or N/R and add them to a listing of VULN
                                a.CHECKLIST = ChecklistLoader.LoadChecklist(a.rawChecklist);
                                // get the main info every single VULN will need
                                hostname = a.CHECKLIST.ASSET.HOST_NAME;
                                checklistType = a.CHECKLIST.STIGS.iSTIG.STIG_INFO.SI_DATA.Where(x => x.SID_NAME == "title").FirstOrDefault().SID_DATA;
                                checklistVersion = a.CHECKLIST.STIGS.iSTIG.STIG_INFO.SI_DATA.Where(x => x.SID_NAME == "version").FirstOrDefault().SID_DATA;
                                checklistRelease = a.CHECKLIST.STIGS.iSTIG.STIG_INFO.SI_DATA.Where(x => x.SID_NAME == "releaseinfo").FirstOrDefault().SID_DATA;
                                foreach (VULN v in a.CHECKLIST.STIGS.iSTIG.VULN){
                                    if (v.STATUS.ToLower() == "open" || v.STATUS.ToLower() == "not_reviewed"){
                                        vulnReport = new VulnerabilityReport();
                                        vulnReport.hostname = hostname;
                                        vulnReport.checklistType = checklistType;
                                        vulnReport.checklistVersion = checklistVersion;
                                        vulnReport.checklistRelease = checklistRelease;
                                        // get the VULN information in here now
                                        vulnReport.vulnid = v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "Vuln_Num").FirstOrDefault().ATTRIBUTE_DATA;
                                        vulnReport.ruleTitle = v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "Rule_Title").FirstOrDefault().ATTRIBUTE_DATA;
                                        vulnReport.comments = v.COMMENTS;
                                        vulnReport.checkContent = v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "Check_Content").FirstOrDefault().ATTRIBUTE_DATA;
                                        vulnReport.fixText = v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "Fix_Text").FirstOrDefault().ATTRIBUTE_DATA;
                                        vulnReport.severity = v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "Severity").FirstOrDefault().ATTRIBUTE_DATA;
                                        vulnReport.severityJustification = v.SEVERITY_JUSTIFICATION;
                                        vulnReport.severityOverride = v.SEVERITY_OVERRIDE;
                                        vulnReport.status = v.STATUS;
                                        
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
    
                                        if (!string.IsNullOrEmpty(cciReferences)) { // split it, put into the listing
                                            List<CciReference> cciRefs = new List<CciReference>();

                                            // now, go get the actual NIST controls for the CCI-REF numbers we collected
                                            // do a search for each of the CCI-REF and then get back all of the "index" field values
                                            // there can be more than 1 per CCI-REF, and more than 1 CCI-REF per VULN record :)
                                            foreach (string cci in cciReferences.Split(",").ToList()) {
                                                if (cciList.Where(z => z.cciId == cci).Select(y => y.references).FirstOrDefault() != null)
                                                    cciRefs.AddRange(cciList.Where(z => z.cciId == cci).Select(y => y.references).FirstOrDefault());
                                            }
                                            if (cciRefs != null && cciRefs.Count > 0) { // get the indexes
                                                foreach (string cciRef in cciRefs.Select(y => y.index).Distinct().ToList()) { // put the distinct strings together for this
                                                    vulnReport.securityControlNumbers += cciRef + "\n";
                                                }
                                            }
                                        }
                                        // add the Vulnerability Report record to the listing we cycle through
                                        vulnerabilities.Add(vulnReport);
                                    }
                                }
                            }

                            // Order by severity, and then VULN ID
                            vulnerabilities = vulnerabilities.OrderByDescending(c => c.severityCategory).ThenBy(d => d.vulnid).ToList();

                            // cycle through the listing and add each VULN ID for a row
                            foreach (VulnerabilityReport vuln in vulnerabilities) {
                                _logger.LogInformation("ExportSystemPOAM({0}) adding Vulnerabilities for {1}", systemGroupId, vuln.vulnid);
                                rowNumber++;

                                // make a new row for this set of items
                                row = MakeDataRow(rowNumber, "A", vuln.ruleTitle, styleIndex);

                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(vuln.securityControlNumbers);
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(vuln.vulnid);
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "0" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(vuln.checklistType + " V" + vuln.checklistVersion + " " + vuln.checklistRelease);
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("Ongoing");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(vuln.comments);
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(vuln.severity);
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = GetVulnerabilityStatus(vuln.status, vuln.severity);
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue(!string.IsNullOrEmpty(vuln.severityJustification)? vuln.severityJustification : "");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AA" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AB" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;
                                newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AC" + rowNumber.ToString() };
                                row.InsertBefore(newCell, refCell);
                                newCell.CellValue = new CellValue("");
                                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                newCell.StyleIndex = styleIndex;

                                // add the row to the sheet now
                                sheetData.Append(row);

                                // merge appropriate data cells    
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("A" + rowNumber + ":B" + rowNumber) });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("D" + rowNumber + ":E" + rowNumber) });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("G" + rowNumber + ":H" + rowNumber) });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("J" + rowNumber + ":M" + rowNumber) });
                                mergeCells.Append(new MergeCell() { Reference = new StringValue("N" + rowNumber + ":O" + rowNumber) });
                            }                        

                        }

                        // save the merged cells
                        worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        _logger.LogInformation("ExportSystemPOAM({0}) closing the XLSX test plan", systemGroupId);
                        spreadSheet.Dispose();
                        // set the filename
                        string filename = sg.title + "-SystemPOAM";
                        memory.Seek(0, SeekOrigin.Begin);
                        _logger.LogInformation("Called ExportSystemPOAM({0}) successfully", systemGroupId);
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(filename));
                    }
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "ExportSystemPOAM() Error getting the test plan XLSX export for system {0}", systemGroupId);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called ExportSystemPOAM() with no system ID");
                return BadRequest(); // no systemGroupId entered
            }
        }

        /// <summary>
        /// GET The list of checklists for the given System ID and download as a zip file
        /// </summary>
        /// <param name="systemGroupId">The ID of the system to use</param>
        /// <param name="naf">True/False include Not a Finding vulnerabilities</param>
        /// <param name="open">True/False include Open vulnerabilities</param>
        /// <param name="na">True/False include Not Applicable vulnerabilities</param>
        /// <param name="nr">True/False include Not Reviewed vulnerabilities</param>
        /// <param name="cat1">True/False include CAT 1 / High vulnerabilities</param>
        /// <param name="cat2">True/False include CAT 2 / Medium vulnerabilities</param>
        /// <param name="cat3">True/False include CAT 3 / Low vulnerabilities</param>
        /// <param name="hostname">The hostname of the checklist to filter on</param>
        /// <returns>
        /// HTTP Status showing it was found or that there is an error. And the list of system records 
        /// exported to an XLSX file to download.
        /// </returns>
        /// <response code="200">Returns the Artifact List of records for the passed in system in XLSX format</response>
        /// <response code="400">If the item did not query correctly</response>
        [HttpGet("system/download/{systemGroupId}")]
        [Authorize(Roles = "Administrator,Reader,Assessor")]
        public async Task<IActionResult> DownloadChecklistListingToZip(string systemGroupId, bool naf = true, bool open = true, bool na = true, 
            bool nr = true, bool cat1 = true, bool cat2 = true, bool cat3 = true, string hostname = "")
        {
            try {
                _logger.LogInformation("Calling DownloadChecklistListingToZip({0})", systemGroupId);
                IEnumerable<Artifact> artifacts;
                // if they pass in a system, get all for that system
                if (string.IsNullOrEmpty(systemGroupId))
                {
                    _logger.LogWarning("DownloadChecklistListingToZip() Getting a listing of all checklists to export to XLSX");
                    return BadRequest("You must specify a system to export");
                }
                _logger.LogInformation(string.Format("DownloadChecklistListingToZip() Getting a listing of all {0} checklists to export to XLSX", systemGroupId));
                artifacts = await _artifactRepo.GetSystemArtifacts(systemGroupId);
                if (!naf || !open || !na || !nr || !cat1 || !cat2 || !cat3) {
                    // used to store the allowed checklists
                    List<string> allowedChecklists = new List<string>();
                    bool allowChecklist = false;
                    List<Score> scoreListing = NATSClient.GetSystemScores(systemGroupId);
                    // get the list of scores for all of them
                    if (scoreListing != null && scoreListing.Count > 0) {
                        // see for each if they should be included
                        foreach (Score s in scoreListing) {
                            allowChecklist = false;
                            // now we can check status and boolean
                            if (s.totalCat1Open > 0 && cat1 && open)
                                allowChecklist = true;
                            else if (s.totalCat2Open > 0 && cat2 && open)
                                allowChecklist = true;
                            else if (s.totalCat3Open > 0 && cat3 && open)
                                allowChecklist = true;
                            else if (s.totalCat1NotReviewed > 0 && cat1 && nr)
                                allowChecklist = true;
                            else if (s.totalCat2NotReviewed > 0 && cat2 && nr)
                                allowChecklist = true;
                            else if (s.totalCat3NotReviewed > 0 && cat3 && nr)
                                allowChecklist = true;
                            else if (s.totalCat1NotApplicable > 0 && cat1 && na)
                                allowChecklist = true;
                            else if (s.totalCat2NotApplicable > 0 && cat2 && na)
                                allowChecklist = true;
                            else if (s.totalCat3NotApplicable > 0 && cat3 && na)
                                allowChecklist = true;
                            else if (s.totalCat1NotAFinding > 0 && cat1 && naf)
                                allowChecklist = true;
                            else if (s.totalCat2NotAFinding > 0 && cat2 && naf)
                                allowChecklist = true;
                            else if (s.totalCat3NotAFinding > 0 && cat3 && naf)
                                allowChecklist = true;
                            else
                                allowChecklist = false; // only if no severity is checked, which is not smart

                            if (allowChecklist) // then add it, we will do a distinct later
                                allowedChecklists.Add(s.artifactId);
                        }
                        if (allowedChecklists.Count == 0)
                            artifacts = new List<Artifact>(); // there are not allowed, so empty it
                        else {
                            // otherwise remove from the systemChecklists before going on
                            artifacts = artifacts.Where(o => allowedChecklists.Contains(o.InternalId.ToString()));
                        }
                    }
                } 
                // check for hostname being used
                if (!string.IsNullOrEmpty(hostname)) {
                    artifacts = artifacts.Where(z => z.hostName.Contains(hostname));
                }

                // now use the listing
                if (artifacts != null && artifacts.Count() > 0) {
                    using (var ms = new MemoryStream())
                    {
                        using (var zipArchive = new ZipArchive(ms, ZipArchiveMode.Create, true))
                        {

                            _logger.LogInformation("DownloadChecklistListingToZip({0}) cycling through checklists to list", systemGroupId);
                            foreach (Artifact art in artifacts.OrderBy(x => x.title).ToList()) {
                                // add to the ZIP from the rawData field
                                ZipArchiveEntry checklistEntry = zipArchive.CreateEntry(art.title.Replace(" ","-") + ".ckl", CompressionLevel.Fastest);
                                // add the TEXT and then close it up
                                using (StreamWriter writer = new StreamWriter(checklistEntry.Open()))
                                {                           
                                    art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.ROLE)) art.CHECKLIST.ASSET.ROLE = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.ASSET_TYPE)) art.CHECKLIST.ASSET.ASSET_TYPE = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_NAME)) art.CHECKLIST.ASSET.HOST_NAME = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_IP)) art.CHECKLIST.ASSET.HOST_IP = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_MAC)) art.CHECKLIST.ASSET.HOST_MAC = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_FQDN)) art.CHECKLIST.ASSET.HOST_FQDN = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.TECH_AREA)) art.CHECKLIST.ASSET.TECH_AREA = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.TARGET_KEY)) art.CHECKLIST.ASSET.TARGET_KEY = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.WEB_OR_DATABASE)) art.CHECKLIST.ASSET.WEB_OR_DATABASE = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.WEB_DB_SITE)) art.CHECKLIST.ASSET.WEB_DB_SITE = "";
                                    if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.WEB_DB_INSTANCE)) art.CHECKLIST.ASSET.WEB_DB_INSTANCE = "";

                                    System.Xml.Serialization.XmlSerializer checklistWriter = new System.Xml.Serialization.XmlSerializer(typeof(CHECKLIST)); 
                                    using(StringWriter textWriter = new StringWriter())                
                                    {
                                        checklistWriter.Serialize(textWriter, art.CHECKLIST);
                                        art.rawChecklist = textWriter.ToString();
                                    }
                                    // strip out all the extra formatting crap and clean up the XML to be as simple as possible
                                    System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(art.rawChecklist, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                                    // save the new serialized checklist record to the database
                                    art.rawChecklist = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                                    string rawChecklist = art.rawChecklist.Substring(art.rawChecklist.IndexOf("<STIGS>")); // save the rest but redo the top part
                                    rawChecklist = string.Format("<CHECKLIST><ASSET><ROLE>{0}</ROLE><ASSET_TYPE>{1}</ASSET_TYPE><HOST_NAME>{2}</HOST_NAME><HOST_IP>{3}</HOST_IP><HOST_MAC>{4}</HOST_MAC><HOST_FQDN>{5}</HOST_FQDN><TECH_AREA>{6}</TECH_AREA><TARGET_KEY>{7}</TARGET_KEY><WEB_OR_DATABASE>{8}</WEB_OR_DATABASE><WEB_DB_SITE>{9}</WEB_DB_SITE><WEB_DB_INSTANCE>{10}</WEB_DB_INSTANCE></ASSET>",
                                        art.CHECKLIST.ASSET.ROLE,art.CHECKLIST.ASSET.ASSET_TYPE,art.CHECKLIST.ASSET.HOST_NAME,art.CHECKLIST.ASSET.HOST_IP,
                                        art.CHECKLIST.ASSET.HOST_MAC,art.CHECKLIST.ASSET.HOST_FQDN,art.CHECKLIST.ASSET.TECH_AREA,
                                        art.CHECKLIST.ASSET.TARGET_KEY,art.CHECKLIST.ASSET.WEB_OR_DATABASE,art.CHECKLIST.ASSET.WEB_DB_SITE,
                                        art.CHECKLIST.ASSET.WEB_DB_INSTANCE) + rawChecklist;
                                    
                                    // take the final cleanup and write it out to the ZIP memory
                                    writer.WriteLine(CleanupData(rawChecklist.Trim()));
                                }
                            }
                        }
                        
                        ms.Position = 0;
                        ms.Seek(0, SeekOrigin.Begin);
                        _logger.LogInformation("Called DownloadChecklistListingToZip({0}) successfully", systemGroupId);
                        return File(ms.ToArray(), "application/zip", "checklistfiles.zip");
                    }
                }
                else {
                    _logger.LogInformation("Calling DownloadChecklistListingToZip({0}) but had no checklists to show", systemGroupId);
                    return NotFound();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "DownloadChecklistListingToZip({0}) Error Retrieving Artifacts for Exporting", systemGroupId);
                return NotFound();
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
                art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist.Replace("\t",""));
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

                // cleanup the checklist data so it loads correctly                        
                _logger.LogInformation("Called DownloadChecklist({0}) cleaning up checklist data", id);
                art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.ROLE)) art.CHECKLIST.ASSET.ROLE = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.ASSET_TYPE)) art.CHECKLIST.ASSET.ASSET_TYPE = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_NAME)) art.CHECKLIST.ASSET.HOST_NAME = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_IP)) art.CHECKLIST.ASSET.HOST_IP = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_MAC)) art.CHECKLIST.ASSET.HOST_MAC = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.HOST_FQDN)) art.CHECKLIST.ASSET.HOST_FQDN = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.TECH_AREA)) art.CHECKLIST.ASSET.TECH_AREA = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.TARGET_KEY)) art.CHECKLIST.ASSET.TARGET_KEY = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.WEB_OR_DATABASE)) art.CHECKLIST.ASSET.WEB_OR_DATABASE = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.WEB_DB_SITE)) art.CHECKLIST.ASSET.WEB_DB_SITE = "";
                if (string.IsNullOrEmpty(art.CHECKLIST.ASSET.WEB_DB_INSTANCE)) art.CHECKLIST.ASSET.WEB_DB_INSTANCE = "";

                System.Xml.Serialization.XmlSerializer checklistWriter = new System.Xml.Serialization.XmlSerializer(typeof(CHECKLIST)); 
                using(StringWriter textWriter = new StringWriter())                
                {
                    checklistWriter.Serialize(textWriter, art.CHECKLIST);
                    art.rawChecklist = textWriter.ToString();
                }
                // strip out all the extra formatting crap and clean up the XML to be as simple as possible
                System.Xml.Linq.XDocument xDoc = System.Xml.Linq.XDocument.Parse(art.rawChecklist, System.Xml.Linq.LoadOptions.PreserveWhitespace);
                // save the new serialized checklist record to the database
                art.rawChecklist = xDoc.ToString(System.Xml.Linq.SaveOptions.DisableFormatting);
                string rawChecklist = art.rawChecklist.Substring(art.rawChecklist.IndexOf("<STIGS>")); // save the rest but redo the top part
                rawChecklist = string.Format("<CHECKLIST><ASSET><ROLE>{0}</ROLE><ASSET_TYPE>{1}</ASSET_TYPE><HOST_NAME>{2}</HOST_NAME><HOST_IP>{3}</HOST_IP><HOST_MAC>{4}</HOST_MAC><HOST_FQDN>{5}</HOST_FQDN><TECH_AREA>{6}</TECH_AREA><TARGET_KEY>{7}</TARGET_KEY><WEB_OR_DATABASE>{8}</WEB_OR_DATABASE><WEB_DB_SITE>{9}</WEB_DB_SITE><WEB_DB_INSTANCE>{10}</WEB_DB_INSTANCE></ASSET>",
                    art.CHECKLIST.ASSET.ROLE,art.CHECKLIST.ASSET.ASSET_TYPE,art.CHECKLIST.ASSET.HOST_NAME,art.CHECKLIST.ASSET.HOST_IP,
                    art.CHECKLIST.ASSET.HOST_MAC,art.CHECKLIST.ASSET.HOST_FQDN,art.CHECKLIST.ASSET.TECH_AREA,
                    art.CHECKLIST.ASSET.TARGET_KEY,art.CHECKLIST.ASSET.WEB_OR_DATABASE,art.CHECKLIST.ASSET.WEB_DB_SITE,
                    art.CHECKLIST.ASSET.WEB_DB_INSTANCE) + rawChecklist;

                _logger.LogInformation("Called DownloadChecklist({0}) successfully", id);
                return Ok(CleanupData(rawChecklist));
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
                    if (art != null && art.CHECKLIST == null && !string.IsNullOrEmpty(art.rawChecklist)) // load the checklist data
                        art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                    if (art != null && art.CHECKLIST != null) {
                        List<string> cciList = new List<string>();
                        _logger.LogInformation("ExportChecklist({0}) formatting the checklist to XML from the raw string format", id);
                        art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                        _logger.LogInformation("ExportChecklist({0}) checklist formatted", id);
                        // starting row number for data
                        uint rowNumber = 9;

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
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 15, CustomWidth = true }); // col A
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 15, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 5, Max = 5, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 6, Max = 6, Width = 60, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 7, Max = 7, Width = 100, CustomWidth = true }); // col G Discussion
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 8, Max = 8, Width = 20, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 9, Max = 9, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 10, Max = 10, Width = 100, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 11, Max = 11, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 12, Max = 12, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 13, Max = 13, Width = 20, CustomWidth = true }); // col M
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 14, Max = 14, Width = 45, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 15, Max = 15, Width = 45, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 16, Max = 16, Width = 45, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 17, Max = 17, Width = 45, CustomWidth = true }); // col Q
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 18, Max = 18, Width = 45, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 19, Max = 19, Width = 45, CustomWidth = true }); // col S Severity Override Guidance
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 20, Max = 20, Width = 30, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 21, Max = 21, Width = 25, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 22, Max = 22, Width = 15, CustomWidth = true }); // col V
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 23, Max = 23, Width = 15, CustomWidth = true }); // col W Status
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 24, Max = 24, Width = 100, CustomWidth = true }); // col X Comments
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
                            row = MakeXLSXInfoRow("System Name", art.systemTitle,2);
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("Checklist Name", art.title,3);
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("Host Name", art.hostName,4);
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("Type", art.stigType,5);
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("Release", art.stigRelease,6);
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("Last Updated", art.updatedOn.Value.ToString("MM/dd/yy hh:mm tt"),7);
                            sheetData.Append(row);
                            row = MakeChecklistHeaderRows(rowNumber);
                            sheetData.Append(row);

                            MergeCells mergeCells = new MergeCells();
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:F1") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:F2") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:F3") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A4:F4") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A5:F5") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A6:F6") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A7:F7") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A8:F8") });

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
                                    styleIndex = GetVulnerabilityStatus(v.STATUS, v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "Severity").Select(y => y.ATTRIBUTE_DATA).FirstOrDefault());
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

                            // save the merged cells
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());

                            // Save the new worksheet.
                            workbookpart.Workbook.Save();
                            // Close the document.
                            spreadSheet.Dispose();
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
        private static string[] headerColumns = {"A", "B", "C", "D", "E", "F", "G", "H", "I", "J", "K", "L", "M", "N", "O", "P", "Q", "R", "S",
            "T", "U", "V", "W", "X", "Y", "Z","AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", 
            "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ"};
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
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeXLSXInfoRow(string title, string value, uint rowindex) {
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
            newCell.StyleIndex = 19;
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
            newCell.CellValue = new CellValue("File Name");
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

        public static DocumentFormat.OpenXml.Spreadsheet.Row MakeHeaderRows(List<string> headers, uint rowindex, uint styleIndex = 3) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };

            if (headers.Count > 0) {
                DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
                row.InsertBefore(newCell, refCell);
                newCell.CellValue = new CellValue(headers[0]);
                newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                newCell.StyleIndex = styleIndex;
                // cycle through all other columns
                for (int i = 1; i < headers.Count; i++)
                {
                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = headerColumns[i] + rowindex.ToString() };
                    row.InsertBefore(newCell, refCell);
                    newCell.CellValue = new CellValue(headers[i]);
                    newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                    newCell.StyleIndex = styleIndex;
                }
            }
            
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakePOAMInfoRow(string colAlabel, string colBvalue, string colIlabel, string colJvalue, string colLlabel, string colMvalue, uint rowindex) {
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString()};
            uint styleIndex = 16; // background color grey

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(colAlabel);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(colBvalue);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;

            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(colIlabel);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(colJvalue);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;

            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(colLlabel);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(colMvalue);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 18;

            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakePOAMHeaderRows(uint rowindex, bool credentialed) {
                        DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            uint styleIndex = 3;

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("POA&M Item ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Control Vulnerability Description");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Security Control Number");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Office/Org");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Security Checks");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Resources Required");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Scheduled Completion Date");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Milestone with Completion Dates");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Milestone Changes");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Source Identifying Vulnerability");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Status");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Comments");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Raw Severity");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "O" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Devices Affected");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Mitigations");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Predisposing Conditions");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Relevance of Threat");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Threat Description");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Likelihood");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Impact");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Impact Description");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Residual Risk Level");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Recommendations");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Resulting Residual Risk after Proposed Mitigations");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;

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
            newCell.CellValue = new CellValue("Open");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 8;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Not a Finding");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 10;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Not Applicable");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 9;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Not Reviewed");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 11;
            return row;
        }
        private uint GetVulnerabilityStatus(string status, string severity) {
            // open = 4, N/A = 5, NotAFinding = 6, Not Reviewed = 7
            if (status.ToLower() == "not_reviewed")
                return 7U;
            if (status.ToLower() == "open") { // need to know the category as well
                if (severity == "high")
                    return 4U;
                if (severity == "medium")
                    return 12U;
                return 13U; // catch all for Open CAT 3 Low
            }
            if (status.ToLower() == "not_applicable")
                return 5U;
            // catch all
            return 6U;
        }
        private uint GetPatchScanStatus(int severity) {
            // critical or high = 3 or 4, medium = 2, low = 1, informational = 0
            if (severity > 2)
                return 4U;
            if (severity == 2)
                return 12U;
            if (severity == 1)
                return 13U;
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

        private string CleanupData (string rawdata) {
            return rawdata.Replace("\t","");
        }
    }
}
