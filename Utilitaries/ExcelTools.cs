using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;

namespace DSToExcel
{
    public static class ExcelTools
    {
        public static MemoryStream ExportDSToExcel(DataSet ds)
        {
            MemoryStream StreamSpreadsheet = new MemoryStream();
            try
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Create(StreamSpreadsheet, SpreadsheetDocumentType.Workbook))
                {
                    // Add a WorkbookPart to the document
                    WorkbookPart workbookpart = spreadsheetDocument.AddWorkbookPart();

                    // Instantiate a Workbook in the WorkbookPart\
                    workbookpart.Workbook = new Workbook();

                    // Add a WorksheetPart to the WorkbookPart
                    WorksheetPart worksheetPart = workbookpart.AddNewPart<WorksheetPart>();

                    // Add Sheets to the Workbook.
                    Sheets sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                    // Declaring Writer to the WorksheetPart on using statement
                    using OpenXmlWriter writer = OpenXmlWriter.Create(worksheetPart);

                    // Declaring row and cell objects to reuse in the construction of the Sheet
                    Row row = null;
                    Cell cell = null;

                    // For each table, generate a sheet and fill it with the correspondent data
                    foreach (DataTable table in ds.Tables)
                    {
                        // Append a new worksheet and associate it with the workbook.
                        Sheet sheet = new Sheet()
                        {
                            Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                            SheetId = 1,
                            Name = table.TableName
                        };
                        sheets.Append(sheet);

                        writer.WriteStartElement(new Worksheet()); // Writes init tag for worksheet
                        writer.WriteStartElement(new SheetData()); // Writes init tag for sheetData

                        #region Header Row

                        row = new Row();
                        writer.WriteStartElement(row); // Init tag Row(Header)

                        List<string> columns = new List<string>(); // Used for searching values in the DataTable

                        foreach (DataColumn column in table.Columns)
                        {
                            columns.Add(column.ColumnName);

                            cell = new Cell();
                            cell.DataType = CellValues.String;
                            cell.CellValue = new CellValue(column.ColumnName);

                            writer.WriteElement(cell);
                        }

                        writer.WriteEndElement(); // End tag Row(Header)

                        #endregion

                        #region Common rows

                        foreach (DataRow dsrow in table.Rows)
                        {
                            row = new Row();
                            writer.WriteStartElement(row); // Init tag Row

                            foreach (string col in columns)
                            {
                                cell = new Cell();

                                // Checks and sets data type for the cell according to the DataTable data type
                                if (dsrow[col] is int)
                                    cell.DataType = CellValues.Number;
                                else if (dsrow[col] is DateTime)
                                    cell.DataType = CellValues.Date;
                                else if (dsrow[col] is bool)
                                    cell.DataType = CellValues.Boolean;
                                else
                                    cell.DataType = CellValues.String;

                                cell.CellValue = new CellValue(dsrow[col].ToString());

                                writer.WriteElement(cell);
                            }

                            writer.WriteEndElement(); // End tag Row
                        }

                        #endregion

                        writer.WriteEndElement(); // End tag SheetData
                        writer.WriteEndElement(); // End tag Worksheet

                        workbookpart.Workbook.Save();

                        writer.Dispose();

                        spreadsheetDocument.Save();
                        spreadsheetDocument.Close();
                    }
                }
            }
            catch (Exception)
            {
                throw;
            }
            return StreamSpreadsheet;
        }
    }
}
