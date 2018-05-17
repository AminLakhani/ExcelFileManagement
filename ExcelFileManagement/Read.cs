using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.Linq;
using System.Reflection;

namespace ExcelFileManagement
{
    class Read
    {
        /// <summary>
        /// Returns list of object from excel sheet.
        /// </summary>
        /// <typeparam name="T"></typeparam>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<T> ReadExcelDocument<T>(string filePath)
        {
            DataTable dt = ReadExcelToDataTable(filePath);
            return ConvertDataTable<T>(dt);
        }

        /// <summary>
        /// Returns string list containing first row from excel sheet.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static List<string> GetExcelColumnNames(string filePath)
        {

            DataTable dt = new DataTable();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
            {

                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();

                foreach (var openXmlElement in rows.ElementAt(0))
                {
                    var cell = (Cell) openXmlElement;
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                return dt.Columns.Cast<DataColumn>()
                                 .Select(x => x.ColumnName)
                                 .ToList();

            }
        }

        /// <summary>
        /// returns datatable of excel sheet.
        /// </summary>
        /// <param name="filePath"></param>
        /// <returns></returns>
        public static DataTable ReadExcelToDataTable(string filePath)
        {
            DataTable dt = new DataTable();

            using (SpreadsheetDocument spreadSheetDocument = SpreadsheetDocument.Open(filePath, false))
            {

                IEnumerable<Sheet> sheets = spreadSheetDocument.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                string relationshipId = sheets.First().Id.Value;
                WorksheetPart worksheetPart = (WorksheetPart)spreadSheetDocument.WorkbookPart.GetPartById(relationshipId);
                Worksheet workSheet = worksheetPart.Worksheet;
                SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                IEnumerable<Row> rows = sheetData.Descendants<Row>();


                var enumerable = rows as Row[] ?? rows.ToArray();

                foreach (var openXmlElement in enumerable.ElementAt(0))
                {
                    var cell = (Cell) openXmlElement;
                    dt.Columns.Add(GetCellValue(spreadSheetDocument, cell));
                }

                foreach (Row row in enumerable)
                {
                    DataRow tempRow = dt.NewRow();
                    
                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                    {
                        string value = GetCellValue(spreadSheetDocument, row.Descendants<Cell>().ElementAt(i));
                        if (value == "-")
                        {
                            value = string.Empty;
                        }
                        tempRow[i] = value;
                    }

                    dt.Rows.Add(tempRow);
                }

            }
            dt.Rows.RemoveAt(0); //remove header row

            return dt;
        }


        private static List<T> ConvertDataTable<T>(DataTable dt)
        {
            List<T> data = new List<T>();
            foreach (DataRow row in dt.Rows)
            {
                T item = GetItem<T>(row);
                data.Add(item);
            }
            return data;
        }

        private static T GetItem<T>(DataRow dr)
        {
            Type temp = typeof(T);
            T obj = Activator.CreateInstance<T>();

            foreach (DataColumn column in dr.Table.Columns)
            {
                foreach (PropertyInfo pro in temp.GetProperties())
                {
                    if (pro.Name == column.ColumnName)
                    {
                        try
                        {
                            if (temp.GetProperty(pro.Name)?.PropertyType.Name == "DateTime")
                            {
                                try
                                {
                                    pro.SetValue(obj, DateTime.Parse(dr[column.ColumnName].ToString()), null);
                                }
                                catch
                                {
                                    double doubleValue = double.Parse(dr[column.ColumnName].ToString());
                                    pro.SetValue(obj, DateTime.FromOADate(doubleValue), null);
                                }
                            }
                            else if (temp.GetProperty(pro.Name)?.PropertyType.Name == "Decimal")
                            {
                                decimal per = decimal.Parse(dr[column.ColumnName].ToString(), NumberStyles.AllowExponent | NumberStyles.Float);
                                per = decimal.Round(per, 5, MidpointRounding.AwayFromZero);
                                pro.SetValue(obj, per, null);

                            }
                            else if (temp.GetProperty(pro.Name)?.PropertyType.Name == "Int32")
                            {
                                pro.SetValue(obj, int.Parse(dr[column.ColumnName].ToString()), null);
                            }
                            else
                            {
                                pro.SetValue(obj, dr[column.ColumnName], null);
                            }
                        }
                        catch (Exception)
                        {
                            throw new Exception("The value '" + dr[column.ColumnName] + "' for the column " + column.ColumnName + " is the wrong datatype");
                        }

                    }

                }
            }
            return obj;
        }


        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
            string value = cell.CellValue.InnerXml;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
            }
            else
            {
                return value;
            }
        }
    }
}
