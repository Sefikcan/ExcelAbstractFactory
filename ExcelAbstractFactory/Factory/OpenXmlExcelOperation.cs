using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelAbstractFactory.Models;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Reflection;

namespace ExcelAbstractFactory
{
    public class OpenXmlExcelOperation : IExcelOperation
    {
        public byte[] ToExcelBytesMultiSheets<T>(List<ExcelByteModel<T>> excelByteModel)
        {
            byte[] excelBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (SpreadsheetDocument document = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    document.AddWorkbookPart();
                    document.WorkbookPart.Workbook = new Workbook();

                    document.WorkbookPart.Workbook.Append(new BookViews(new WorkbookView()));

                    WorkbookStylesPart stylesPart = document.WorkbookPart.AddNewPart<WorkbookStylesPart>();

                    uint worksheetNumber = 1;

                    foreach (var excelByte in excelByteModel)
                    {
                        WorksheetPart newWorksheetPart = document.WorkbookPart.AddNewPart<WorksheetPart>();
                        newWorksheetPart.Worksheet = new Worksheet();

                        var sheetData = new SheetData();
                        newWorksheetPart.Worksheet.AppendChild(sheetData);

                        var typeName = GetSimpleTypeName(excelByte.ExcelItemList);

                        PropertyInfo[] props = typeof(T).GetProperties();
                        List<PropertyInfo> propList = GetSelectedProperties(props, excelByte.IncludeFields, excelByte.ExcludeFields);

                        List<string> colFixList = null;
                        if (!string.IsNullOrEmpty(excelByte.ColumnFixes))
                        {
                            colFixList = excelByte.ColumnFixes.Split(',').ToList();
                        }

                        var headerRow = sheetData.AppendChild(new Row());
                        int colIdx = 0;
                        foreach (var prop in propList)
                        {
                            var colName = prop.Name.Replace("__", " ");

                            if (colFixList != null)
                            {
                                foreach (var item in colFixList)
                                {
                                    if (item.Contains(colName))
                                    {
                                        colName = item;
                                        break;
                                    }
                                }
                            }

                            headerRow.AppendChild(new Cell
                            {
                                CellValue = new CellValue(colName),
                                DataType = CellValues.String,
                                CellReference = GetColumnAddress(colIdx) + "1"
                            });
                            colIdx++;
                        }

                        int rowIdx = 1;
                        foreach (var item in excelByte.ExcelItemList)
                        {
                            var contentRow = sheetData.AppendChild(new Row());
                            colIdx = 0;
                            foreach (var prop in propList)
                            {
                                var value = prop.GetValue(item, null);

                                if (prop.PropertyType == typeof(bool))
                                {
                                    value = true ? "1" : "0";
                                }

                                contentRow.AppendChild(new Cell
                                {
                                    CellValue = new CellValue(value?.ToString()),
                                    DataType = CellValues.String,
                                    CellReference = GetColumnAddress(colIdx) + (rowIdx + 1).ToString()
                                });

                                colIdx++;
                            }
                            rowIdx++;
                        }

                        newWorksheetPart.Worksheet.Save();

                        if (worksheetNumber == 1)
                        {
                            document.WorkbookPart.Workbook.AppendChild(new Sheets());
                        }

                        document.WorkbookPart.Workbook.GetFirstChild<Sheets>().AppendChild(new Sheet()
                        {
                            Id = document.WorkbookPart.GetIdOfPart(newWorksheetPart),
                            SheetId = worksheetNumber,
                            Name = string.IsNullOrEmpty(excelByte.SheetName) ? typeName : excelByte.SheetName
                        });
                        worksheetNumber++;
                    }

                    document.WorkbookPart.Workbook.Save();
                }

                excelBytes = memoryStream.ToArray();
            }

            return excelBytes;
        }

        public byte[] ToExcelBytesOneSheet<T>(ExcelByteModel<T> excelByteModel)
        {
            byte[] excelBytes;

            using (var memoryStream = new MemoryStream())
            {
                using (var spreadsheetDocument = SpreadsheetDocument.Create(memoryStream, SpreadsheetDocumentType.Workbook))
                {
                    var workbookpart = spreadsheetDocument.AddWorkbookPart();
                    workbookpart.Workbook = new Workbook();
                    workbookpart.Workbook.AppendChild(new FileVersion { ApplicationName = "Microsoft Office Excel" });

                    var worksheetPart = workbookpart.AddNewPart<WorksheetPart>();
                    worksheetPart.Worksheet = new Worksheet();

                    var sheetData = new SheetData();
                    worksheetPart.Worksheet.AppendChild(sheetData);

                    var typeName = GetSimpleTypeName(excelByteModel.ExcelItemList);

                    var sheets = spreadsheetDocument.WorkbookPart.Workbook.AppendChild(new Sheets());

                    sheets.AppendChild(new Sheet
                    {
                        Id = spreadsheetDocument.WorkbookPart.GetIdOfPart(worksheetPart),
                        SheetId = 1,
                        Name = string.IsNullOrEmpty(excelByteModel.SheetName) ? typeName : excelByteModel.SheetName
                    });

                    PropertyInfo[] props = typeof(T).GetProperties();
                    List<PropertyInfo> propList = GetSelectedProperties(props, excelByteModel.IncludeFields, excelByteModel.ExcludeFields);

                    List<string> colFixList = null;
                    if (!string.IsNullOrEmpty(excelByteModel.ColumnFixes))
                    {
                        colFixList = excelByteModel.ColumnFixes.Split(',').ToList();
                    }

                    var headerRow = sheetData.AppendChild(new Row());
                    int colIdx = 0;
                    foreach (var prop in propList)
                    {
                        var colName = prop.Name.Replace("__", " ");

                        if (colFixList != null)
                        {
                            foreach (var item in colFixList)
                            {
                                if (item.Contains(colName))
                                {
                                    colName = item;
                                    break;
                                }
                            }
                        }

                        headerRow.AppendChild(new Cell
                        {
                            CellValue = new CellValue(colName),
                            DataType = CellValues.String,
                            CellReference = GetColumnAddress(colIdx) + "1"
                        });
                        colIdx++;
                    }

                    int rowIdx = 1;
                    foreach (var item in excelByteModel.ExcelItemList)
                    {
                        var contentRow = sheetData.AppendChild(new Row());
                        colIdx = 0;
                        foreach (var prop in propList)
                        {
                            var value = prop.GetValue(item, null);

                            if (prop.PropertyType == typeof(bool))
                            {
                                value = true ? "1" : "0";
                            }

                            contentRow.AppendChild(new Cell
                            {
                                CellValue = new CellValue(value?.ToString()),
                                DataType = CellValues.String,
                                CellReference = GetColumnAddress(colIdx) + (rowIdx + 1).ToString()
                            });

                            colIdx++;
                        }
                        rowIdx++;
                    }
                    workbookpart.Workbook.Save();
                }

                excelBytes = memoryStream.ToArray();
            }

            return excelBytes;
        }

        #region private methods

        //Excelde kolon setleme işlemini yapar A1,A2 vs
        private static string GetColumnAddress(int columnIndex)
        {
            Stack<char> stack = new Stack<char>();
            while (columnIndex >= 0)
            {
                stack.Push((char)('A' + (columnIndex % 26)));
                columnIndex = (columnIndex / 26) - 1;
            }
            return new string(stack.ToArray());
        }

        private static string GetSimpleTypeName<T>(IList<T> list)
        {
            string typeName = list.GetType().ToString();
            int pos = typeName.IndexOf("[") + 1;
            typeName = typeName.Substring(pos, typeName.LastIndexOf("]") - pos);
            typeName = typeName.Substring(typeName.LastIndexOf(".") + 1);
            return typeName;
        }

        //Modelde gelen ama excelde olmasını istemediğimiz kolonları exclude,include yapmamızı sağlar, bunu  admin panel tarzı 
        // yerlerde kullanabiliriz, bunun dışında her excel tipi için class oluşturmak daha iyi olacaktır.
        private static List<PropertyInfo> GetSelectedProperties(PropertyInfo[] props, string include, string exclude)
        {
            List<PropertyInfo> propList = new List<PropertyInfo>();
            if (!string.IsNullOrEmpty(include))
            {
                var includeProps = include.ToLower().Split(',').ToList();
                foreach (var item in props)
                {
                    var propName = includeProps.Where(a => a == item.Name.ToLower()).FirstOrDefault();
                    if (!string.IsNullOrEmpty(propName))
                        propList.Add(item);
                }
            }
            else if (!string.IsNullOrEmpty(exclude))
            {
                var excludeProps = exclude.ToLower().Split(',');
                foreach (var item in props)
                {
                    var propName = excludeProps.Where(a => a == item.Name.ToLower()).FirstOrDefault();
                    if (string.IsNullOrEmpty(propName))
                        propList.Add(item);
                }
            }
            else
            {
                propList.AddRange(props.ToList());
            }
            return propList;
        }

        #endregion
    }
}
