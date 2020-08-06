using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Data;
using System.IO;

namespace MultipleWorksheet
{
    class Program
    {
        public static void Main(string[] args)
        {
            WorkbookPart workbookPart = null;

            try
            {
                using (var memoryStream = new MemoryStream())
                {
                    using (var excel = SpreadsheetDocument.Create(memoryStream, DocumentFormat.OpenXml.SpreadsheetDocumentType.Workbook, true))
                    {
                        workbookPart = excel.AddWorkbookPart();
                        workbookPart.Workbook = new Workbook();
                        uint sheetId = 1;
                        excel.WorkbookPart.Workbook.Sheets = new Sheets();
                        Sheets sheets = excel.WorkbookPart.Workbook.GetFirstChild<Sheets>();

                        DataSet dataSet = new DataSet();

                        for (int i = 0; i < 5; i++)
                        {
                            DataTable dataTable = new DataTable();
                            dataTable.TableName = "Table" + (i + 1).ToString();

                            dataTable.Columns.Add("Column1", typeof(string));
                            dataTable.Columns.Add("Column2", typeof(string));
                            dataTable.Columns.Add("Column3", typeof(string));
                            dataTable.Columns.Add("Column4", typeof(string));
                            dataTable.Columns.Add("Column5", typeof(string));


                            for (int j = 0; j < 5; j++)
                            {
                                DataRow dataRow = dataTable.NewRow();

                                for (int k = 0; k < dataTable.Columns.Count; k++)
                                {
                                    dataRow[k] = "Row" + (j + 1).ToString() + ", " + "Column" + (k + 1).ToString();
                                }

                                dataTable.Rows.Add(dataRow);
                            }

                            dataSet.Tables.Add(dataTable);
                           
                        }


                        for (int i = 0; i < dataSet.Tables.Count; i++)
                        {
                            string relationshipId = "rId" + (i + 1).ToString();
                            WorksheetPart wSheetPart = workbookPart.AddNewPart<WorksheetPart>(relationshipId);
                            string sheetName = dataSet.Tables[i].TableName;
                            Sheet sheet = new Sheet() { Id = relationshipId, SheetId = sheetId, Name = sheetName };
                            sheets.Append(sheet);

                            Worksheet worksheet = new Worksheet();

                            wSheetPart.Worksheet = worksheet;

                            SheetData sheetData = new SheetData();
                            worksheet.Append(sheetData);

                            string[] excelColumns = new string[] {"A","B","C","D","E","F","G" };

                            for (int l = 0; l < dataSet.Tables[i].Rows.Count; l++)
                            {
                                for (int m = 0; m < dataSet.Tables[i].Columns.Count; m++)
                                {
                                    AddToCell(sheetData,Convert.ToUInt32(l+1), excelColumns[m], CellValues.String, Convert.ToString(dataSet.Tables[i].Rows[l][m]));
                                }
                            }

                           
                            sheetId++;
                        }


                        excel.Close();
                    }

                    FileStream fileStream = new FileStream(AppDomain.CurrentDomain.BaseDirectory + "MultipleWorkSheet.xlsx", FileMode.Create, FileAccess.Write);
                    memoryStream.WriteTo(fileStream);
                    fileStream.Close();
                    memoryStream.Close();
                }
            }
            catch (Exception ex)
            {

                throw ex;
            }


        }


        private static void AddToCell(SheetData sheetData, UInt32 uint32rowIndex, string strColumnName, DocumentFormat.OpenXml.EnumValue<CellValues> CellDataType, string strCellValue)
        {
            Row row = new Row() { RowIndex = uint32rowIndex };
            Cell cell = new Cell();

            cell = new Cell();
            cell.CellReference = strColumnName + row.RowIndex.ToString();
            cell.DataType = CellDataType;
            cell.CellValue = new CellValue(strCellValue);
            row.AppendChild(cell);

            sheetData.Append(row);
        }
    }
}
