using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using Microsoft.AspNetCore.Mvc;
using openstig_read_api.Classes;
using openstig_read_api.Models;
using System.IO;
using System.Text;
using Microsoft.AspNetCore;
using Microsoft.AspNetCore.Hosting;
using System.Xml.Serialization;
using System.Xml;
using Microsoft.AspNetCore.Cors.Infrastructure;
using Newtonsoft.Json;
using Microsoft.Extensions.Options;
using Microsoft.AspNetCore.Http;
using Microsoft.EntityFrameworkCore;
using Microsoft.Extensions.Logging;

using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using openstig_read_api.Data;

namespace openstig_read_api.Controllers
{
    //[Route("[controller]")]
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
        public async Task<IActionResult> ListArtifacts()
        {
            try {
                IEnumerable<Artifact> artifacts;
                artifacts = await _artifactRepo.GetAllArtifacts();
                foreach (Artifact a in artifacts) {
                    a.rawChecklist = string.Empty;
                }
                return Ok(artifacts);
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error listing all artifacts and deserializing the checklist XML");
                return BadRequest();
            }
        }

        // GET /value
        [HttpGet("{id}")]
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
        public async Task<IActionResult> ExportChecklist(string id)
        {
            try {
                Artifact art = new Artifact();
                art = await _artifactRepo.GetArtifact(id);
                if (art != null && art.CHECKLIST != null) {
                    art.CHECKLIST = ChecklistLoader.LoadChecklist(art.rawChecklist);

                    // starting row number for data
                    uint rowNumber = 8;

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

                        DocumentFormat.OpenXml.Spreadsheet.Row row = MakeTitleRow("openSTIG by Cingulara");
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Checklist Name", art.title,2);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Description", art.description,3);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Type", art.typeTitle,4);
                        sheetData.Append(row);
                        row = MakeChecklistInfoRow("Last Updated", art.updatedOn.Value.ToString("MM/dd/yy hh:mm tt"),5);
                        sheetData.Append(row);
                        row = MakeHeaderRows(rowNumber);
                        sheetData.Append(row);

                        // cycle through the vulnerabilities
                        foreach (VULN v in art.CHECKLIST.STIGS.iSTIG.VULN) {
                            rowNumber++;
                            // make a new row for this set of items
                            row = MakeDataRow(rowNumber, "A", v.STIG_DATA[1].ATTRIBUTE_DATA);
                            // now cycle through the rest of the items
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[1].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[2].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[3].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[4].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[5].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[6].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[7].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[8].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[9].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[10].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[11].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[12].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[13].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "O" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[14].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[15].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[16].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[17].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[18].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[19].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[20].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[21].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STATUS);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.COMMENTS);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.FINDING_DETAILS);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.SEVERITY_OVERRIDE);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AA" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.SEVERITY_JUSTIFICATION);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AB" + rowNumber.ToString() };
                            row.InsertBefore(newCell, refCell);
                            newCell.CellValue = new CellValue(v.STIG_DATA[24].ATTRIBUTE_DATA);
                            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
                            newCell.StyleIndex = 0;
                            sheetData.Append(row);
                        }

                        // Save the new worksheet.
                        workbookpart.Workbook.Save();
                        // Close the document.
                        spreadSheet.Close();

                        memory.Seek(0, SeekOrigin.Begin);
                        //return new FileStreamResult(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
                        return File(memory, "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet", CreateXLSXFilename(art.title));
                    }
                }
                else {
                    return NotFound();
                }
            }
            catch (Exception ex) {
                _logger.LogError(ex, "Error Retrieving Artifact for Download");
                return NotFound();
            }
        } 

        private string CreateXLSXFilename(string title) {
            return title.Replace(" ", "-") + ".xlsx";
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
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeHeaderRows(uint rowindex) {
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowindex };
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "A" + rowindex.ToString() };
            //row.Height = 25;
            //row.CustomHeight = true;
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Vuln ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "B" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "C" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Group Title");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "D" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Rule ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "E" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("STIG ID");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "F" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Rule Title");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "G" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Discussion");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "H" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("IA Controls");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "I" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Check Content");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "J" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Fix Text");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "K" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("False Positives");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "L" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("False Negatives");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "M" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Documentable");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "N" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Mitigations");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "O" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Potential Impact");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "P" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Third Party Tools");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Q" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Mitigation Control");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "R" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Responsibility");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "S" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity Override Guidance");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "T" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Check Content Reference");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "U" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Classification");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "V" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("STIG");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "W" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Status");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "X" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Comments");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Y" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Finding Details");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "Z" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity Override");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AA" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("Severity Override Justification");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            // next column
            newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = "AB" + rowindex.ToString() };
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue("CCI");
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 3;
            return row;
        }
        private DocumentFormat.OpenXml.Spreadsheet.Row MakeDataRow(uint rowNumber, string cellReference, string value) {
            DocumentFormat.OpenXml.Spreadsheet.Row row = new DocumentFormat.OpenXml.Spreadsheet.Row() { RowIndex = rowNumber };
            DocumentFormat.OpenXml.Spreadsheet.Cell refCell = null;
            DocumentFormat.OpenXml.Spreadsheet.Cell newCell = new DocumentFormat.OpenXml.Spreadsheet.Cell() { CellReference = cellReference + rowNumber.ToString()};
            row.InsertBefore(newCell, refCell);
            newCell.CellValue = new CellValue(value);
            newCell.DataType = new EnumValue<CellValues>(CellValues.String);
            newCell.StyleIndex = 0;
            return row;
        }
        #endregion

        /******************************************
        * Dashboard Specific API calls
        */
        #region Dashboard APIs
        // GET /count
        [HttpGet("count")]
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
        public async Task<IActionResult> GetCountByType()
        {
            try {
                IEnumerable<Object> artifacts;
                artifacts = await _artifactRepo.GetCountByType();
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
