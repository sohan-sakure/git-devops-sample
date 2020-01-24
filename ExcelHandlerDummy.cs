using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;

using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using BulkUpdateCore.Utitlities;

namespace BulkUpdateCore.Helpers
{
    //public static class ExcelHandlerDummy
    //{
    //    private const string V = "..\\..\\Template\\Template.xlsx";

    //    private static readonly log4net.ILog Log
    //        = log4net.LogManager.GetLogger(System.Reflection.MethodBase.GetCurrentMethod().DeclaringType);

    //    static string ReplaceHexadecimalSymbols(string txt)
    //    {
    //        string r = "[\x00-\x08\x0B\x0C\x0E-\x1F\x26]";
    //        return Regex.Replace(txt, r, "", RegexOptions.Compiled);
    //    }
    //    public static void GenerateExcelFromTemplate(DataTable table, string excelfileName, string sheetName = "Sheet 2")
    //    {
    //        try
    //        {
    //            var outPutDirectory = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
    //            string sourceFilePath = Path.Combine(outPutDirectory, "Template\\Template.xlsx");
    //            string outputFilePath = excelfileName;
    //            File.Copy(sourceFilePath, outputFilePath, true);
    //            SheetData sheetData = null;
    //            using (SpreadsheetDocument document = SpreadsheetDocument.Open(excelfileName,true))
    //            {
    //                WorkbookPart workbookPart = document.WorkbookPart;
    //              //  workbookPart.Workbook = new Workbook();

    //                //WorksheetPart worksheetPart = workbookPart.AddNewPart<WorksheetPart>();
    //                var currWorksheet = workbookPart.Workbook
    //                    .GetFirstChild<Sheets>().Elements<Sheet>().FirstOrDefault(s => s.Name == sheetName);
    //                if (currWorksheet != null)
    //                {
    //                    var currWorksheetPart = workbookPart.GetPartById(currWorksheet.Id.Value) as WorksheetPart;
    //                    if (currWorksheetPart != null)
    //                    {
    //                        sheetData = currWorksheetPart.Worksheet.GetFirstChild<SheetData>();
    //                    }
    //                }

    //                //workbookPart.Worksheet = new Worksheet(sheetData);

    //                Sheets sheets = workbookPart.Workbook.Sheets;
    //               // Sheet sheet = new Sheet() { Id = workbookPart.GetIdOfPart(currWorksheetPart), SheetId = 1, Name = sheetName };

    //               // if (currWorksheet != null) sheets.Append(currWorksheet);

    //                Row headerRow = new Row();

    //                List<string> columns = new List<string>();
    //                foreach (System.Data.DataColumn column in table.Columns)
    //                {
    //                    columns.Add(column.ColumnName);

    //                    //Cell cell = new Cell
    //                    //{
    //                    //    DataType = CellValues.String,
    //                    //    CellValue = new CellValue(ReplaceHexadecimalSymbols(column.ColumnName))
    //                    //};
    //                    //headerRow.AppendChild(cell);
    //                }

    //              //  sheetData.AppendChild(headerRow);

    //                foreach (DataRow dsrow in table.Rows)
    //                {
    //                    Row newRow = new Row();
    //                    foreach (var col in columns)
    //                    {
    //                        Cell cell;
    //                        if (col.Contains("Id"))
    //                        {
    //                            cell = new Cell
    //                            {
    //                                DataType = CellValues.Number,
    //                                CellValue = new CellValue(dsrow[col].ToString()), 
    //                            };
    //                        }
    //                        else
    //                        {
    //                            cell = new Cell
    //                            {
    //                                DataType = CellValues.String,
    //                                CellValue = new CellValue(ReplaceHexadecimalSymbols(dsrow[col].ToString()))
    //                            };
    //                        }

    //                        newRow.AppendChild(cell);
    //                    }

    //                    if (sheetData != null) sheetData.AppendChild(newRow);
                        
    //                }
                    
    //                workbookPart.Workbook.Save();
    //            }

    //        }
    //        catch (Exception exception)
    //        {
    //            Log.Error(exception.Message);
    //        }
    //    }

    //    public static DataTable GenerateDataTableFromExcel(string excelfileName)
    //    {
    //        var dt = new DataTable();

    //        using (var spreadSheetDocument = SpreadsheetDocument.Open(excelfileName, false))
    //        {

    //            var workbookPart = spreadSheetDocument.WorkbookPart;
    //            var sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
    //            var relationshipId = sheets.First().Id.Value;
    //            var worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
    //            var workSheet = worksheetPart.Worksheet;
    //            var sheetData = workSheet.GetFirstChild<SheetData>();
    //            var rows = sheetData.Descendants<Row>();

    //            if (rows == null) return dt;

    //            var enumerable = rows.ToList();
    //            //  dt.Columns.Add("Blank");
    //            foreach (var openXmlElement in enumerable.ElementAt(0))
    //            {
    //                var cell = (Cell)openXmlElement;
    //                var columnName = GetCellValue(spreadSheetDocument, cell);
    //                dt.Columns.Add(columnName);
    //                var columnIndex = Regex.Replace(cell.CellReference.Value, @"[\d-]", string.Empty);

    //                Global.ColumnNameCellReferenceMap.Add(columnName, columnIndex);
    //            }

    //            foreach (var row in enumerable) //this will also include your header row...
    //            {
    //                var tempRow = dt.NewRow();

    //                for (var i = 0; i < row.Descendants<Cell>().Count(); i++)
    //                {
    //                    //tempRow[i] = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
    //                    Cell cell = row.Descendants<Cell>().ElementAt(i);
    //                    int actualCellIndex = CellReferenceToIndex(cell);
    //                    tempRow[actualCellIndex] = GetCellValue(spreadSheetDocument, cell);
    //                }

    //                dt.Rows.Add(tempRow);
    //            }
    //            dt.Rows.RemoveAt(0); //...so i'm taking it out here.
    //        }

    //        return dt;
    //    }

    //    private static string GetCellValue(SpreadsheetDocument document, Cell cell)
    //    {
    //        var stringTablePart = document.WorkbookPart.SharedStringTablePart;
    //        string value = cell.CellValue.InnerXml;

    //        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
    //        {
    //            return stringTablePart.SharedStringTable.ChildElements[int.Parse(value)].InnerText;
    //        }

    //        return value;
    //    }

    //    private static int CellReferenceToIndex(Cell cell)
    //    {
    //        int index = 0;
    //        string reference = cell.CellReference.ToString().ToUpper();
    //        foreach (char ch in reference)
    //        {
    //            if (Char.IsLetter(ch))
    //            {
    //                int value = (int)ch - (int)'A';
    //                index = (index == 0) ? value : ((index + 1) * 26) + value;
    //            }
    //            else
    //                return index;
    //        }
    //        return index;
    //    }
    //}
}
