// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using openrmf_read_api.Classes;
using Microsoft.Extensions.Logging;
using Microsoft.AspNetCore.Authorization;
using System.IO;
using System.Linq;
using openrmf_read_api.Models;
using openrmf_read_api.Models.Compliance;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using openrmf_read_api.Data;

namespace openrmf_read_api.Controllers
{
    [Route("/compliance")]
    public class ComplianceController : Controller
    {
        private readonly ILogger<ComplianceController> _logger;
	    private readonly IArtifactRepository _artifactRepo;
        private readonly ISystemGroupRepository _systemGroupRepo;

        public ComplianceController(IArtifactRepository artifactRepo, ISystemGroupRepository systemGroupRepo, ILogger<ComplianceController> logger)
        {
            _artifactRepo = artifactRepo;
            _systemGroupRepo = systemGroupRepo;
            _logger = logger;
        }

        /// <summary>
        /// GET The compliance of a system based on the system ID passed in, if it has PII, 
        /// and whatever filter we have such as the impact level.
        /// </summary>
        /// <param name="id">The ID of the system to generate compliance</param>
        /// <param name="filter">The filter to show impact level of Low, Moderate, High</param>
        /// <param name="pii">A boolean to say if this has PII or not.  There are 
        ///        specific things to check compliance if this has PII.
        /// <param name="majorcontrol">The filter to show only show compliance with the major control passed in</param>
        /// </param>
        /// <returns>
        /// HTTP Status showing it was generated as well as the list of compliance records and status. 
        /// Also shows the status of each checklist per major control.
        /// </returns>
        /// <response code="200">Returns the generated compliance listing</response>
        /// <response code="400">If the item did not generate it correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("system/{id}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetCompliancBySystem(string id, string filter, bool pii, string majorcontrol = "")
        {
            if (!string.IsNullOrEmpty(id)) {
                try {
                    _logger.LogInformation("Calling GetCompliancBySystem({0}, {1}, {2})", id, filter, pii.ToString());
                    IEnumerable<Artifact> checklists = await _artifactRepo.GetSystemArtifacts(id);
                    var result = ComplianceGenerator.GetSystemControls(id, checklists.ToList(), filter, pii, majorcontrol);
                    if (result != null && result.Result != null && result.Result.Count > 0) {
                        _logger.LogInformation("Called GetCompliancBySystem({0}, {1}, {2}) successfully", id, filter, pii.ToString());
                        return Ok(result);
                    }
                    else {
                        _logger.LogWarning("Called GetCompliancBySystem({0}, {1}, {2}) but had no returned data", id, filter, pii.ToString());
                        return NotFound(); // bad system reference
                    }
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "GetCompliancBySystem() Error listing all checklists for system {0}", id);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called GetCompliancBySystem() but with an invalid or empty Id", id);
                return BadRequest(); // no term entered
            }
        }
        
        /// <summary>
        /// GET The compliance of a system based on the system ID passed in, if it has PII, 
        /// and whatever filter we have such as the impact level. Export as XLSX.
        /// </summary>
        /// <param name="id">The ID of the system to generate compliance</param>
        /// <param name="filter">The filter to show impact level of Low, Moderate, High</param>
        /// <param name="pii">A boolean to say if this has PII or not.  There are 
        ///        specific things to check compliance if this has PII.
        /// <param name="majorcontrol">The filter to show only show compliance with the major control passed in</param>
        /// </param>
        /// <returns>
        /// HTTP Status showing it was generated as well as the list of compliance records and status. 
        /// Also shows the status of each checklist per major control.
        /// </returns>
        /// <response code="200">Returns the generated compliance listing</response>
        /// <response code="400">If the item did not generate it correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("system/{id}/export")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        public async Task<IActionResult> GetCompliancBySystemExport(string id, string filter, bool pii, string majorcontrol = "")
        {
            if (!string.IsNullOrEmpty(id)) {
                try {
                    _logger.LogInformation("Calling GetCompliancBySystemExport({0}, {1}, {2})", id, filter, pii.ToString());
                    // verify system information
                    SystemGroup sg = await _systemGroupRepo.GetSystemGroup(id);
                    if (sg == null) {
                        _logger.LogInformation("Called GetCompliancBySystemExport({0}, {1}, {2}) invalid System Group", id, filter, pii.ToString());
                        return NotFound();
                    }

                    IEnumerable<Artifact> checklists = await _artifactRepo.GetSystemArtifacts(id);
                    var result = ComplianceGenerator.GetSystemControls(id, checklists.ToList(), filter, pii, majorcontrol);
                    if (result != null && result.Result != null && result.Result.Count > 0) {
                        _logger.LogInformation("Called GetCompliancBySystemExport({0}, {1}, {2}) successfully. Putting into XLSX.", id, filter, pii.ToString());

                        // starting row
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
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 1, Max = 1, Width = 20, CustomWidth = true }); // col A
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 2, Max = 2, Width = 60, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 3, Max = 3, Width = 50, CustomWidth = true });
                                lstColumns.Append(new DocumentFormat.OpenXml.Spreadsheet.Column() { Min = 4, Max = 4, Width = 25, CustomWidth = true });
                                worksheetPart.Worksheet.InsertAt(lstColumns, 0);
                            }

                            // Add Sheets to the Workbook.
                            Sheets sheets = spreadSheet.WorkbookPart.Workbook.AppendChild<Sheets>(new Sheets());

                            // Append a new worksheet and associate it with the workbook.
                            Sheet sheet = new Sheet() { Id = spreadSheet.WorkbookPart.
                                GetIdOfPart(worksheetPart), SheetId = 1, Name = "System-Compliance" };
                            sheets.Append(sheet);
                            // Get the sheetData cell table.
                            SheetData sheetData = worksheetPart.Worksheet.GetFirstChild<SheetData>();
                            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
                            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = null;

                            DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("OpenRMF by Cingulara and Tutela");
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("System Package", sg.title,2);
                            sheetData.Append(row);
                            row = MakeXLSXInfoRow("Generated", DateTime.Now.ToString("MM/dd/yy hh:mm tt"),3);
                            sheetData.Append(row);
                            row = MakeComplianceHeaderRows(rowNumber);
                            sheetData.Append(row);

                            MergeCells mergeCells = new MergeCells();
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A1:D1") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A2:D2") });
                            mergeCells.Append(new MergeCell() { Reference = new StringValue("A3:D3") });
                            
                            uint styleIndex = 0; // use this for 4, 5, 6, or 7 for status

                            _logger.LogInformation("GetCompliancBySystemExport() cycling through all the vulnerabilities");

                            foreach (NISTCompliance nist in result.Result) {
                                if (nist.complianceRecords.Count > 0) {
                                    foreach( ComplianceRecord rec in nist.complianceRecords) {
                                        rowNumber++;
                                        styleIndex = GetVulnerabilityStatus(rec.status, "high");
                                        // make a new row for this set of items
                                        row = MakeDataRow(rowNumber, "A", nist.control, styleIndex);
                                        // now cycle through the rest of the items
                                        newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                                        row.InsertBefore(newCell, refCell);
                                        newCell.CellValue = new CellValue(nist.title);
                                        newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                        newCell.StyleIndex = styleIndex;
                                        newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                                        row.InsertBefore(newCell, refCell);
                                        newCell.CellValue = new CellValue(rec.title);
                                        newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                        newCell.StyleIndex = styleIndex;
                                        newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                                        row.InsertBefore(newCell, refCell);
                                        // print out status, if N/A or NAF then just NAF
                                        if (rec.status.ToLower() == "open")
                                            newCell.CellValue = new CellValue("Open");
                                        else if (rec.status.ToLower() == "not_reviewed")
                                            newCell.CellValue = new CellValue("Not Reviewed");
                                        else
                                            newCell.CellValue = new CellValue("Not a Finding");
                                        newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                                        newCell.StyleIndex = styleIndex;
                                        sheetData.Append(row);
                                    }
                                } else {
                                    rowNumber++;
                                    styleIndex = 18;
                                    // make a new row for this set of items
                                    row = MakeDataRow(rowNumber, "A", nist.control, styleIndex);
                                    // now cycle through the rest of the items
                                    newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                                    row.InsertBefore(newCell, refCell);
                                    newCell.CellValue = new CellValue(nist.title);
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
                                    sheetData.Append(row);
                                }
                            }

                            // save the merged cells
                            worksheetPart.Worksheet.InsertAfter(mergeCells, worksheetPart.Worksheet.Elements<SheetData>().First());
                            // Save the new worksheet.
                            workbookpart.Workbook.Save();
                            // Close the document.
                            spreadSheet.Close();
                            // set the filename
                            string filename = sg.title;
                            if (!string.IsNullOrEmpty(sg.title) && sg.title.ToLower().Trim() == "none")
                                filename = sg.title.Trim() + "-" + filename; // add the system onto the front
                            // return the file
                            memory.Seek(0, SeekOrigin.Begin);
                            _logger.LogInformation("Called GetCompliancBySystemExport({0}, {1}, {2}) successfully", id, filter, pii.ToString());
                            return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(filename));
                        } // end of using statement
                    }
                    else {
                        _logger.LogWarning("Called GetCompliancBySystemExport({0}, {1}, {2}) but had no returned data", id, filter, pii.ToString());
                        return NotFound(); // bad system reference
                    }
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "GetCompliancBySystemExport() Error exporting Compliance for system {0}", id);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called GetCompliancBySystemExport() but with an invalid or empty system group Id", id);
                return BadRequest(); // no term entered
            }
        }

        /// <summary>
        /// GET The CCI item information and references based on the CCI ID passed in
        /// </summary>
        /// <param name="cciid">The CCI Number/ID of the system to generate compliance</param>
        /// <returns>
        /// HTTP Status showing the CCI item is here and the CCI item record with references.
        /// </returns>
        /// <response code="200">Returns the CCI item record</response>
        /// <response code="400">If the item did not generate it correctly</response>
        /// <response code="404">If the ID passed in is not valid</response>
        [HttpGet("cci/{cciid}")]
        [Authorize(Roles = "Administrator,Reader,Editor,Assessor")]
        [ResponseCache(Duration = 60, Location = ResponseCacheLocation.Any, VaryByQueryKeys = new [] {"cciid"})]
        public async Task<IActionResult> GetCCIItem(string cciid)
        {
            if (!string.IsNullOrEmpty(cciid)) {
                try {
                    _logger.LogInformation("Calling GetCCIItem({0})", cciid);
                    var result = NATSClient.GetCCIItemReferences(cciid);
                    if (result != null) {
                        _logger.LogInformation("Called GetCCIItem({0}) successfully", cciid);
                        return Ok(result);
                    }
                    else {
                        _logger.LogWarning("Called GetCCIItem({0}) but had no returned data", cciid);
                        return NotFound(); // bad system reference
                    }
                }
                catch (Exception ex) {
                    _logger.LogError(ex, "GetCCIItem() Error getting CCI Item information for {0}", cciid);
                    return BadRequest();
                }
            }
            else {
                _logger.LogWarning("Called GetCCIItem() but with an invalid or empty Id", cciid);
                return BadRequest(); // no CCI Id entered
            }
        }

        private uint GetVulnerabilityStatus(string status, string severity = "high") {
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

        #region Excel Formatting
        private string CreateXLSXFilename(string title) {
            return title.Trim().Replace(" ", "_") + ".xlsx";
        }

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
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeComplianceHeaderRows(uint rowindex) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            uint styleIndex = 3;

            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Control");
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
            newCell.CellValue = new CellValue("Checklist");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = styleIndex;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Status");
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

        #endregion
    }
}
