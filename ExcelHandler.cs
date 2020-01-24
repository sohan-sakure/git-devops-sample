using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using BulkUpdateCore.Utitlities;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using log4net;
using SharedModels;

namespace BulkUpdateCore.Helpers
{
    public static class ExcelHandler
    {
        private static readonly ILog Log
            = LogManager.GetLogger(MethodBase.GetCurrentMethod().DeclaringType);

        private static string ReplaceHexadecimalSymbols(string txt)
        {
            const string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
            return Regex.Replace(txt, r, "", RegexOptions.Compiled);
        }

        private static int CellReferenceToIndex(Cell cell)
        {
            var index = 0;
            var reference = cell.CellReference?.ToString().ToUpper();
            foreach (var ch in reference)
                if (char.IsLetter(ch))
                {
                    var value = ch - 'A';
                    index = index == 0 ? value : (index + 1) * 26 + value;
                }
                else
                {
                    return index;
                }

            return index;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart;
            string value;
            try
            {
                stringTablePart = document.WorkbookPart.SharedStringTablePart;
                value = cell.CellValue?.InnerXml;

                if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                    return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
                return value;
            }
            catch (Exception ex)
            {
                Log.Error($"Error in getting cell value: {cell.CellValue} " + ex.Message);
                return null;
            }
        }

        #region Public Methods

        public static void GenerateExcel(DataTable table, string excelfileName, string sheetName = "Sheet 1")
        {
            try
            {
                using (var document = SpreadsheetDocument.Create(excelfileName, SpreadsheetDocumentType.Workbook))
                {
                    var workbookPart = document.AddWorkbookPart();
                    workbookPart.Workbook = new Workbook();

                    var worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var sheetData = new SheetData();
                    worksheetPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbookPart.Workbook.AppendChild(new Sheets());
                    var sheet = new Sheet { Id = workbookPart.GetIdOfPart(worksheetPart), SheetId = 1, Name = sheetName };

                    sheets.Append(sheet);

                    var headerRow = new Row();

                    var columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                    {
                        columns.Add(column.ColumnName);

                        var cell = new Cell
                        {
                            DataType = CellValues.String,
                            CellValue = new CellValue(ReplaceHexadecimalSymbols(column.ColumnName))
                        };
                        headerRow.AppendChild(cell);
                    }

                    sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        var newRow = new Row();
                        foreach (var col in columns)
                        {
                            var cell = new Cell
                            {
                                DataType = CellValues.String,
                                CellValue = new CellValue(ReplaceHexadecimalSymbols(dsrow[col].ToString()))
                            };
                            newRow.AppendChild(cell);
                        }

                        sheetData.AppendChild(newRow);
                    }

                    workbookPart.Workbook.Save();
                }
            }
            catch (Exception exception)
            {
                Log.Error(exception.Message);
            }
        }

        public static void GenerateExcelFromTemplate(DataTable table, string excelfileName,
            string sheetName = "Sheet 2")
        {
            try
            {
                var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                if (outPutDirectory != null)
                {
                    var sourceFilePath = Path.Combine(outPutDirectory, "Template\\Template.xlsx");
                    var outputFilePath = excelfileName;
                    File.Copy(sourceFilePath, outputFilePath, true);
                }
                else
                {
                    Log.Error("Failed to get directory path of executing assembly");
                    return;
                }

                SheetData sheetData = null;
                using (var document = SpreadsheetDocument.Open(excelfileName, true))
                {
                    var workbookPart = document.WorkbookPart;
                    //  workbookPart.Workbook = new Workbook();

                    //WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
                    var currWorksheet = workbookPart.Workbook
                        .GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
                    if (currWorksheet != null)
                    {
                        var currWorksheetPart = workbookPart.GetPartById(currWorksheet.Id.Value) as WorksheetPart;
                        if (currWorksheetPart != null)
                            sheetData = currWorksheetPart.Worksheet.GetFirstChild<SheetData>();
                    }

                    //workbookPart.Worksheet = new Worksheet(sheetData);

                    var sheets = workbookPart.Workbook.Sheets;
                    // Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(currWorksheetPart), SheetId = 1, Name = sheetName };

                    // if (currWorksheet != null) sheets.Append(currWorksheet);

                    var headerRow = new Row();

                    var columns = new List<string>();
                    foreach (DataColumn column in table.Columns)
                        columns.Add(column.ColumnName);

                    //Cell cell = new Cell
                    //{
                    //    DataType = CellValues.String,
                    //    CellValue = new CellValue(ReplaceHexadecimalSymbols(column.ColumnName))
                    //};
                    //headerRow.AppendChild(cell);

                    //  sheetData.AppendChild(headerRow);

                    foreach (DataRow dsrow in table.Rows)
                    {
                        var newRow = new Row();
                        foreach (var col in columns)
                        {
                            Cell cell;
                            if (col.Contains("Id"))
                                cell = new Cell
                                {
                                    DataType = CellValues.Number,
                                    CellValue = new CellValue(dsrow[col].ToString())
                                };
                            else
                                cell = new Cell
                                {
                                    DataType = CellValues.String,
                                    CellValue = new CellValue(ReplaceHexadecimalSymbols(dsrow[col].ToString()))
                                };

                            newRow.AppendChild(cell);
                        }

                        if (sheetData != null) sheetData.AppendChild(newRow);
                    }

                    workbookPart.Workbook.Save();
                }
            }
            catch (Exception exception)
            {
                Log.Error(exception.Message);
            }
        }

        public static DataTable GenerateDataTableFromExcel(string excelfileName)
        {
            var dt = new DataTable();

            using (var spreadSheetDocument = SpreadsheetDocument.Open(excelfileName, false))
            {
                var workbookPart = spreadSheetDocument.WorkbookPart;
                var sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                var relationshipId = sheets.First().Id.Value;
                var worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                var workSheet = worksheetPart.Worksheet;
                var sheetData = workSheet.GetFirstChild<SheetData>();
                var rows = sheetData.Descendants<Row>();

                if (rows == null) return dt;

                var enumerable = rows.ToList();
                //  dt.Columns.Add("Blank");
                Global.ClearColumnCellReferenceMaps();
                foreach (var openXmlElement in enumerable.ElementAt(0))
                {
                    var cell = (Cell)openXmlElement;
                    var columnName = GetCellValue(spreadSheetDocument, cell);
                    dt.Columns.Add(columnName);
                    var columnIndex = Regex.Replace(cell.CellReference.Value, @"[\d-]", string.Empty);

                    Global.ColumnNameCellReferenceMap.Add(columnName, columnIndex);
                }
                Global.ClearRowIndexes();
                foreach (var row in enumerable) //this will also include your header row...
                {

                    var tempRow = dt.NewRow();
                    int i;
                    if (row.ToString() == String.Empty)
                    {

                    }
                    for (i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        //tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        var cell = row.Descendants<Cell>().ElementAt(i);
                        var actualCellIndex = CellReferenceToIndex(cell);
                        tempRow[actualCellIndex] = GetCellValue(spreadSheetDocument, cell);
                        
                    }
                    try
                    {
                        if (string.IsNullOrEmpty(tempRow[0].ToString()))
                        {
                            //last row of excel
                        }

                        if (row.RowIndex > 1 && !string.IsNullOrEmpty(tempRow[0].ToString()))
                            Global.MetaIdRowIndexMap.Add(int.Parse(tempRow[0].ToString()), (int)row.RowIndex.Value);

                        dt.Rows.Add(tempRow);
                    }
                    catch (Exception exception)
                    {

                        RConsole.WriteLineRed(exception.StackTrace + exception.InnerException?.StackTrace);
                    }

                }

                dt.Rows.RemoveAt(0); //...so i'm taking it out here.
            }

            return dt;
        }


        //TODO: Needs re-work... 
        public static void UpdateExcelRecordStatus(List<MediaInfo> medialist, string filename)
        {
            if (medialist?.Count == 0)
            {
                Log.Error("Media list for updating the excel sheet is either empty or null");
                return;
            }
            if (!File.Exists(filename))
            {
                Log.Error("File doesn't exist for exporting report" + filename);
                return;
            }

            List<Row> rows;
            try
            {
                using (var spreadSheetDocument = SpreadsheetDocument.Open(filename, true))
                {
                    var workbookPart = spreadSheetDocument.WorkbookPart;
                    //  var sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                    //    var relationshipId = sheets.First().Id.Value;
                    var worksheetPart = GetWorksheetPartByName(spreadSheetDocument, "Sheet 2");
                    var workSheet = worksheetPart.Worksheet;
                    var sheetData = workSheet.GetFirstChild<SheetData>();
                    rows = sheetData.Descendants<Row>().ToList();

                    if (rows.Count < 0) throw new Exception("The excel provided doesn't contain any rows");

                    foreach (var mediaInfo in medialist)
                    {
                        var index = Global.MetaIdRowIndexMap.ContainsKey(mediaInfo.MetaId)
                            ? Global.MetaIdRowIndexMap[mediaInfo.MetaId]
                            : 0;

                        if (index != 0)
                        {
                            var rowIndex = unchecked((uint)index);
                            var processingNotesColumnIndex =
                                Global.ColumnNameCellReferenceMap.ContainsKey("Import Notes (RO - For IDT use)")
                                    ? Global.ColumnNameCellReferenceMap["Import Notes (RO - For IDT use)"]
                                    : "X";
                            var successColumnIndex = Global.ColumnNameCellReferenceMap.ContainsKey("Update this media item?")
                                ? Global.ColumnNameCellReferenceMap["Update this media item?"]
                                : "C";

                            var updateNotes = UpdateCell(worksheetPart, mediaInfo.ProcessingNotes, rowIndex,
                                processingNotesColumnIndex);
                            if (!updateNotes)
                                Log.Error("Updating of excel cell " + processingNotesColumnIndex + rowIndex +
                                          " failed. Message to update: " + mediaInfo.ProcessingNotes);

                            var updateSuccess = UpdateCell(worksheetPart, mediaInfo.AreChangesRequired ? "Yes" : "No", rowIndex,
                                successColumnIndex);
                            if (!updateSuccess)
                                Log.Error("Updating of excel cell " + successColumnIndex + rowIndex +
                                          " failed. Changing the column \"Are changes required\" to : " + mediaInfo.AreChangesRequired + " failed.");

                            // Save the worksheet.
                            worksheetPart.Worksheet.Save();
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Log.Error("Something failed while updating the excel sheet" + ex.Message + "\n" + ex.StackTrace);
            }

        }

        #endregion

        #region Helper methods

        public static bool UpdateCell(WorksheetPart worksheetPart, string text,
            uint rowIndex, string columnName)
        {
            if (worksheetPart != null)
            {
                var cell = GetCell(worksheetPart.Worksheet,
                    columnName, rowIndex);

                if (cell != null)
                {
                    cell.CellValue = new CellValue(text);
                    cell.DataType =
                        new EnumValue<CellValues>(CellValues.String);
                    return true;
                }
            }

            return false;
        }

        //public static void UpdateCell(WorksheetPart worksheetPart, bool flag,
        //    uint rowIndex, string columnName)
        //{

        //    if (worksheetPart != null)
        //    {
        //        Cell cell = GetCell(worksheetPart.Worksheet,
        //            columnName, rowIndex);

        //        cell.CellValue = new CellValue(flag.ToString());
        //        cell.DataType =
        //            new EnumValue<CellValues>(CellValues.Boolean);

        //        //// Save the worksheet.
        //        //worksheetPart.Worksheet.Save();
        //    }


        //  }

        #region Private Methods

        private static WorksheetPart
            GetWorksheetPartByName(SpreadsheetDocument document,
                string sheetName)
        {
            var sheets =
                document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>()
                    .Where(s => s.Name == sheetName);

            if (sheets.Count() == 0) return null;

            var relationshipId = sheets.First().Id.Value;
            var worksheetPart = (WorksheetPart)
                document.WorkbookPart.GetPartById(relationshipId);
            return worksheetPart;
        }

        // Given a worksheet, a column name, and a row index, 
        // gets the cell at the specified column and 
        private static Cell GetCell(Worksheet worksheet,
            string columnName, uint rowIndex)
        {
            var row = GetRow(worksheet, rowIndex);

            var cell = row?.Elements<Cell>().FirstOrDefault(c => string.Compare
                                                                 (c.CellReference.Value, columnName +
                                                                                         rowIndex,
                                                                     StringComparison.InvariantCultureIgnoreCase) == 0);

            return cell;
        }


        // Given a worksheet and a row index, return the row.
        private static Row GetRow(Worksheet worksheet, uint rowIndex)
        {
            return worksheet.GetFirstChild<SheetData>().Elements<Row>().FirstOrDefault(r => r.RowIndex == rowIndex);
        }

        #endregion

        #endregion
    }
}