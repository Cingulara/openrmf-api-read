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
        private readonly ILogger<ReadController> _logger;

        public ReadController(IArtifactRepository artifactRepo, ILogger<ReadController> logger)
        {
            _logger = logger;
            _artifactRepo = artifactRepo;
        }

        // GET the listing with Ids of the Checklist artifacts, but without all the extra XML
        [HttpGet]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ListArtifacts()
        {
            try {
                IEnumerable<Artifact> artifacts;
                artifacts = await _artifactRepo.GetAllArtifacts();
                foreach (Artifact a in artifacts) {
                    a.rawChecklist = string.Empty;
                }
                return Ok(artifacts.OrderBy(x => x.title).OrderBy(y => y.system).ToList());
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error listing all artifacts and deserializing the checklist XML");
                return BadRequest();
            }
        }

        // GET /export
        [HttpGet("export")]
        [Authorize(Roles = "Administrator,Reader,Assessor")]
        public async Task<IActionResult> ExportChecklistListing(string system = null)
        {
            try {
                IEnumerable<Artifact> artifacts;
                // if they pass in a system, get all for that system
                if (string.IsNullOrEmpty(system))
                {
                    _logger.LogInformation("Getting a listing of all checklists to export to XLSX");
                    artifacts = await _artifactRepo.GetAllArtifacts();
                }
                else {
                    _logger.LogInformation(string.Format("Getting a listing of all {0} checklists to export to XLSX", system));
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

                        DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("openRMF by Cingulara");
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Checklist Listing", system, 2);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Printed Date", DateTime.Now.ToString("MM/dd/yy hh:mm"),3);
                        sheetData.Append(row);
                        row = MakeChecklistListingHeaderRows(rowNumber);
                        sheetData.Append(row);

                        uint styleIndex = 0; // use this for 4, 5, 6, or 7 for status
                        Score checklistScore;

                        // cycle through the checklists and grab the score for each individually
                        foreach (Artifact art in artifacts.OrderBy(x => x.title).OrderBy(y => y.system).ToList()) {
                            art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                            checklistScore = WebClient.GetChecklistScore(art.InternalId.ToString()).GetAwaiter().GetResult();
                            rowNumber++;

                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", art.system.Trim().ToLower() != "none"? art.system : "", styleIndex);
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
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", "ChecklistListing.xlsx");
                    }
                }
                else {
                    return NotFound();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifacts for Exporting");
                return NotFound();
            }
        } 

        // GET the distinct list of systems
        [HttpGet("systems")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ListArtifactSystems()
        {
            try {
                IEnumerable<ChecklistSystem> systems;
                systems = await _artifactRepo.GetAllSystems();
                return Ok(systems);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error listing all checklist systems");
                return BadRequest();
            }
        }
        
        // GET the list of checklist records for a systems
        [HttpGet("systems/{term}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> ListArtifactsBySystem(string term)
        {
            if (!string.IsNullOrEmpty(term)) {
                try {
                    IEnumerable<Artifact> systemChecklists;
                    systemChecklists = await _artifactRepo.GetSystemArtifacts(term);
                    return Ok(systemChecklists);
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "Error listing all checklists for system {0}", term);
                    return BadRequest();
                }
            }
            else
                return BadRequest(); // no term entered
        }
        
        // GET /value
        [HttpGet("{id}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetArtifact(string id)
        {
            try {
                Artifact art = new Artifact();
                art = await _artifactRepo.GetArtifact(id);
                art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                art.rawChecklist = string.Empty;
                return Ok(art);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact");
                return NotFound();
            }
        }
        
        // GET /download/value
        [HttpGet("download/{id}")]
        [Authorize(Roles = "Administrator,Editor,Assessor,Reader")]
        public async Task<IActionResult> DownloadChecklist(string id)
        {
            try {
                Artifact art = new Artifact();
                art = await _artifactRepo.GetArtifact(id);
                return Ok(art.rawChecklist);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact for Download");
                return NotFound();
            }
        }
        
        // GET /export/value
        [HttpGet("export/{id}")]
        [Authorize(Roles = "Administrator,Editor,Assessor,Reader")]
        public async Task<IActionResult> ExportChecklist(string id, bool nf, bool open, bool na, bool nr, string ctrl)
        {
            try {
                if (!string.IsNullOrEmpty(id)) {
                    Artifact art = new Artifact();
                    art = await _artifactRepo.GetArtifact(id);
                    if (art != null && art.CHECKLIST != null) {
                        List<string> cciList = new List<string>();
                        art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);

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

                            DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("openRMF by Cingulara");
                            sheetData.Append(row);
                            row = MakeChecklistInfoRow("System Name", art.system,2);
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
                            if (!string.IsNullOrEmpty(ctrl))
                                cciList = WebClient.GetCCIListing(ctrl).GetAwaiter().GetResult();

                            // cycle through the vulnerabilities to export into columns
                            foreach (VULN v in art.CHECKLIST.STIGS.iSTIG.VULN) {
                                // if this is a regular checklist, make sure the filter for VULN ID is checked before we add this to the list
                                // if this is from a compliance listing, only add the VULN IDs from the control to the listing
                                // the VULN has a CCI_REF field where the attribute would be the value in the cciList if it is valid
                                if (ShowVulnerabilityInExcel(v.STATUS, nf, nr, open, na) || 
                                        (!string.IsNullOrEmpty(ctrl) && 
                                        v.STIG_DATA.Where(x => x.VULN_ATTRIBUTE == "CCI_REF" && cciList.Contains(x.ATTRIBUTE_DATA)).FirstOrDefault() != null))  {
                                    rowNumber++;
                                    styleIndex = GetVulnerabilitiStatus(v.STATUS);
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
                                    newCell.CellValue = new CellValue(v.STIG_DATA[24].ATTRIBUTE_DATA);
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
                            if (!string.IsNullOrEmpty(art.system) && art.system.ToLower().Trim() == "none")
                                filename = art.system.Trim() + "-" + filename; // add the system onto the front
                            // return the file
                            memory.Seek(0, SeekOrigin.Begin);
                            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(filename));
                        }
                    }
                    else {
                        return NotFound();
                    }
                }
                else { // did not pass in an id
                    _logger.LogInformation("Did not pass in an id in the Export of a single checklist.");
                    return BadRequest();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact for Exporting");
                return NotFound();
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

        // GET /value
        [HttpGet("{id}/control/{control}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetArtifactVulnIdsByControl(string id, string control)
        {
            try {
                if (!string.IsNullOrEmpty(id) && !string.IsNullOrEmpty(control)) {
                    _logger.LogInformation("Invalid Artifact Id {0} or Control {1}", id, control);
                    Artifact art = new Artifact();
                    art = await _artifactRepo.GetArtifact(id);
                    art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);
                    // go get the list of CCIs to look for
                    List<string> cciList = WebClient.GetCCIListing(control).GetAwaiter().GetResult();
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
                        return Ok(vulnIds.Distinct().OrderBy(z => z).ToList());
                    }
                    else
                        return BadRequest();
                }
                else {
                    // log the values passed in
                    _logger.LogWarning("Invalid Artifact Id {0} or Control {1}", 
                        !string.IsNullOrEmpty(id)? id : "null", !string.IsNullOrEmpty(control)? control : "null");
                    return BadRequest();    
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact");
                return NotFound();
            }
        }


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
        private uint GetVulnerabilitiStatus(string status) {
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
        #endregion

        /******************************************
        * Dashboard Specific API calls
        */
        #region Dashboard APIs
        // GET /count
        [HttpGet("count")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> CountArtifacts(string id)
        {
            try {
                long result = await _artifactRepo.CountChecklists();
                return Ok(result);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact Count in MongoDB");
                return NotFound();
            }
        }
        // GET /latest
        [HttpGet("latest/{number}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetLatestArtifacts(int number)
        {
            try {
                IEnumerable<Artifact> artifacts;
                artifacts = await _artifactRepo.GetLatestArtifacts(number);
                foreach (Artifact a in artifacts) {
                    a.rawChecklist = string.Empty;
                }
                return Ok(artifacts);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error listing latest {0} artifacts and deserializing the checklist XML", number.ToString());
                return BadRequest();
            }
        }

        
        // GET /latest
        [HttpGet("counttype")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetCountByType(string system)
        {
            try {
                IEnumerable<Object> artifacts;
                artifacts = await _artifactRepo.GetCountByType(system);
                return Ok(artifacts);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error getting the counts by type for the Reports page");
                return BadRequest();
            }
        }
        #endregion
    }
}
