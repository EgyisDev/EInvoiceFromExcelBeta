using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using ExcelDataReader;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
using System.Text.Json;
using System.Text.RegularExpressions;
using System.Text.Unicode;
using System.Threading.Tasks;
using System.Web;

namespace ThinInvoiceUploaderClient
{
    internal class ExcelHelper
    {
        public static DocumentRequestDocumentForSerializeDto GetDocument(string filePath)
        {
            if (!ValidateExcelSheetColumns(filePath))
                return null;

            var dt = GetExcelDataTable(filePath);
            var x = ConvertExcelDataTableToDto(dt);
            var jsonDocument = ConvertExcelDataTableToJson(dt);//.Replace(@"\", "");
            //string validJsonString = JsonConvert.DeserializeObject(jsonDocument);
            var jsonWithGroupedInvoiceLines = ProcessJson(jsonDocument);
            return JsonConvert.DeserializeObject<DocumentRequestDocumentForSerializeDto>(jsonWithGroupedInvoiceLines);
        }


        public static DocumentRequestDocumentForSerializeDto ConvertExcelDataTableToDto(DataTable dataTable)
        {
            // Get columns containing multiple rows, which will be handled differently
            var columnsWithMultipleRows = new List<string> { "invoiceLines", "taxTotals" };

            var columnsWithNumericData = new[]
            {
        "grossWeight","netWeight","invoiceLines","quantity","salesTotal","total","valueDifference","totalTaxableFees","netTotal","itemsDiscount","unitValue","amountEGP","rate","amount","totalDiscountAmount","totalSalesAmount","netAmount","totalAmount","extraDiscountAmount","totalItemsDiscountAmount"
    };

            var dto = new DocumentRequestDocumentForSerializeDto();

            // Loop through each row in the DataTable
            foreach (DataRow row in dataTable.Rows)
            {
                // Loop through each column in the row
                foreach (DataColumn column in dataTable.Columns)
                {
                    // Get the value for the current column in the current row
                    var value = row[column.ColumnName];

                    // Split the column name into its parts, separated by periods
                    var columnNameSegments = column.ColumnName.Split('.');

                    // If the column is one that contains multiple rows or is the first row of a non-multi-row column
                    if (!columnsWithMultipleRows.Contains(columnNameSegments[0]) && dataTable.Rows.IndexOf(row) == 0)
                    {
                        // Create a nested object to hold the current column's value
                        var nestedObject = dto;

                        // Loop through each part of the column name, except the last part
                        for (int i = 0; i < columnNameSegments.Length - 1; i++)
                        {
                            var key = columnNameSegments[i];
                            // If the key doesn't exist in the current nested object, add it as a new object
                            var property = nestedObject.GetType().GetProperty(key);
                            if (property == null)
                            {
                                // Create a new instance of the property's type
                                var newObject = Activator.CreateInstance(nestedObject.GetType().GetProperty(key.Substring(0, key.Length - 2)).PropertyType);
                                // Set the value of the property to the new object
                                nestedObject.GetType().GetProperty(key.Substring(0, key.Length - 2)).SetValue(nestedObject, newObject);
                                // Get the property again now that it has been created
                                property = nestedObject.GetType().GetProperty(key);
                            }
                            if (property.GetValue(nestedObject) == null)
                            {
                                var obj = Activator.CreateInstance(property.PropertyType);
                                property.SetValue(nestedObject, obj);
                            }
                            // Set the nested object to the value of the current key, so that we can keep nesting
                            nestedObject = (dynamic)property.GetValue(nestedObject);
                        }

                        // Set the value of the last part of the column name to the current value
                        var lastKey = columnNameSegments.Last();
                        if (columnsWithNumericData.Contains(lastKey) && string.IsNullOrEmpty(value.ToString()))
                        {
                            value = 0;
                        }
                        var lastProperty = nestedObject.GetType().GetProperty(lastKey);
                        lastProperty.SetValue(nestedObject, value);
                    }
                    else if (columnsWithMultipleRows.Contains(columnNameSegments[0]))
                    {
                        // If the current row is the first row of a multi-row column, create a new list to hold the rows
                        if (dataTable.Rows.IndexOf(row) == 0)
                        {
                            var property = dto.GetType().GetProperty(columnNameSegments[0]);
                            var listType = typeof(List<>).MakeGenericType(property.PropertyType.GenericTypeArguments[0]);
                            var list = Activator.CreateInstance(listType);
                            property.SetValue(dto, list);
                        }

                        // Get the list of rows for the current column
                        var rowsList = (IList)dto.GetType().GetProperty(columnNameSegments[0]).GetValue(dto);

                        // Create a new object to hold the current row's data
                        var rowObject = Activator.CreateInstance(rowsList.GetType().GenericTypeArguments[0]);

                        // Loop through each part of the column name, except the first part
                        for (int i = 1; i < columnNameSegments.Length; i++)
                        {
                            var key = columnNameSegments[i];
                            // If the key doesn't exist in the current nested object, add it as a new object
                            var property = rowObject.GetType().GetProperty(key);
                            if (property.GetValue(rowObject) == null)
                            {
                                var obj = Activator.CreateInstance(property.PropertyType);
                                property.SetValue(rowObject, obj);
                            }
                            // Set the nested object to the value of the current key, so that we can keep nesting
                            rowObject = (dynamic)property.GetValue(rowObject);
                        }

                        // Set the value of the last part of the column name to the current value
                        var lastKey = columnNameSegments.Last();
                        if (columnsWithNumericData.Contains(lastKey) && string.IsNullOrEmpty(value.ToString()))
                        {
                            value = 0;
                        }
                        var lastProperty = rowObject.GetType().GetProperty(lastKey);
                        lastProperty.SetValue(rowObject, value);

                        // If the current row is the last row of the multi-row column, add the list to the main DTO
                        if (dataTable.Rows.IndexOf(row) == dataTable.Rows.Count - 1)
                        {
                            var property = dto.GetType().GetProperty(columnNameSegments[0]);
                            property.SetValue(dto, rowsList);
                        }
                    }
                }
            }

            return dto;
        }


        public static string ConvertExcelDataTableToJson(DataTable dataTable)
        {
            var columnsWithMultipleRows = new List<string> { "invoiceLines", "taxTotals" };
            var columnsWithNumericData = new[] {
        "grossWeight", "netWeight", "invoiceLines", "quantity", "salesTotal",
        "total", "valueDifference", "totalTaxableFees", "netTotal", "itemsDiscount",
        "unitValue", "amountEGP", "rate", "amount", "totalDiscountAmount",
        "totalSalesAmount", "netAmount", "totalAmount", "extraDiscountAmount",
        "totalItemsDiscountAmount"
    };
            var result = new List<Dictionary<string, object>>();
            foreach (DataRow row in dataTable.Rows)
            {
                var rowDictionary = new Dictionary<string, object>();
                foreach (DataColumn column in dataTable.Columns)
                {
                    var value = row[column.ColumnName];
                    var columnNameSegments = column.ColumnName.Split('.');
                    if (columnsWithMultipleRows.Contains(columnNameSegments[0]))
                    {
                        HandleColumnWithMultipleRows(rowDictionary, columnNameSegments, value, columnsWithNumericData);
                    }
                    else if (dataTable.Rows.IndexOf(row) == 0)
                    {
                        HandleSingleRowColumn(rowDictionary, columnNameSegments, value, columnsWithNumericData);
                    }
                }
                result.Add(rowDictionary);
            }
            var documents = new Dictionary<string, object> { ["documents"] = result };
            return JsonConvert.SerializeObject(documents, Formatting.None);
        }

        private static void HandleSingleRowColumn(
            Dictionary<string, object> rowDictionary, string[] columnNameSegments,
            object value, string[] columnsWithNumericData)
        {
            var nestedObject = rowDictionary;
            for (int i = 0; i < columnNameSegments.Length - 1; i++)
            {
                var key = columnNameSegments[i];
                if (!nestedObject.TryGetValue(key, out var temp))
                {
                    temp = new Dictionary<string, object>();
                    nestedObject[key] = temp;
                }
                nestedObject = (Dictionary<string, object>)temp;
            }
            nestedObject[columnNameSegments.Last()] = string.IsNullOrWhiteSpace(value.ToString())
                ? (columnsWithNumericData.Contains(columnNameSegments.Last()) ? 0 : "")
                : value;
        }

        private static void HandleColumnWithMultipleRows(
            Dictionary<string, object> rowDictionary, string[] columnNameSegments,
            object value, string[] columnsWithNumericData)
        {
            if (!rowDictionary.TryGetValue(columnNameSegments[0], out var temp))
            {
                temp = new Dictionary<string, object>();
                rowDictionary[columnNameSegments[0]] = temp;
            }
            var nestedObject = (Dictionary<string, object>)temp;
            for (int i = 1; i < columnNameSegments.Length - 1; i++)
            {
                var key = columnNameSegments[i];
                if (!nestedObject.TryGetValue(key, out var temp2))
                {
                    temp2 = new Dictionary<string, object>();
                    nestedObject[key] = temp2;
                }
                nestedObject = (Dictionary<string, object>)temp2;
            }
            nestedObject[columnNameSegments.Last()] = string.IsNullOrWhiteSpace(value.ToString())
? (columnsWithNumericData.Contains(columnNameSegments.Last()) ? 0 : "")
                : value;
        }



            // This method takes in a DataTable and converts it to a JSON string
            /*   public static string ConvertExcelDataTableToJson(DataTable dataTable)
           {
               // Get columns containing multiple rows, which will be handled differently
               var columnsWithMultipleRows = new List<string> { "invoiceLines", "taxTotals" };

               var columnsWithNumericData = new[]
               {
                   "grossWeight","netWeight","invoiceLines","quantity","salesTotal","total","valueDifference","totalTaxableFees","netTotal","itemsDiscount","unitValue","amountEGP","rate","amount","totalDiscountAmount","totalSalesAmount","netAmount","totalAmount","extraDiscountAmount","totalItemsDiscountAmount"
               };
               // Create a list to hold each row's data as a dictionary
               var result = new List<Dictionary<string, object>>();

               // Loop through each row in the DataTable
               foreach (DataRow row in dataTable.Rows)
               {
                   // Create a dictionary to hold the current row's data
                   var rowDictionary = new Dictionary<string, object>();

                   // Loop through each column in the row
                   foreach (DataColumn column in dataTable.Columns)
                   {
                       // Get the value for the current column in the current row
                       var value = row[column.ColumnName];

                       // Split the column name into its parts, separated by periods
                       var columnNameSegments = column.ColumnName.Split('.');

                       // If the column is one that contains multiple rows or is the first row of a non-multi-row column
                       if (!columnsWithMultipleRows.Contains(columnNameSegments[0]) && dataTable.Rows.IndexOf(row) == 0)
                       {
                           // Create a nested object to hold the current column's value
                           var nestedObject = rowDictionary;

                           // Loop through each part of the column name, except the last part
                           for (int i = 0; i < columnNameSegments.Length - 1; i++)
                           {
                               var key = columnNameSegments[i];
                               // If the key doesn't exist in the current nested object, add it as an empty dictionary
                               if (!nestedObject.ContainsKey(key))
                               {
                                   nestedObject[key] = new Dictionary<string, object>();
                               }
                               // Set the nested object to the value of the current key, so that we can keep nesting
                               nestedObject = (Dictionary<string, object>)nestedObject[key];
                           }

                           // Set the value of the last part of the column name to the current value
                           nestedObject[columnNameSegments.Last()] = !string.IsNullOrEmpty(value.ToString()) ? value : (columnsWithNumericData.Contains(columnNameSegments.Last()) ? 0 : "");
                       }
                       else if (columnsWithMultipleRows.Contains(columnNameSegments[0]))
                       {
                           // Create a nested object to hold the current column's value
                           if (!rowDictionary.ContainsKey(columnNameSegments[0]))
                           {
                               rowDictionary[columnNameSegments[0]] = new Dictionary<string, object>();
                           }
                           var nestedObject = rowDictionary;

                           // Loop through each part of the column name, except the last part
                           for (int i = 0; i < columnNameSegments.Length - 1; i++)
                           {
                               var key = columnNameSegments[i];
                               // If the key doesn't exist in the current nested object, add it as an empty dictionary
                               if (!nestedObject.ContainsKey(key))
                               {
                                   nestedObject[key] = new Dictionary<string, object>();
                               }
                               // Set the nested object to the value of the current key, so that we can keep nesting
                               nestedObject = (Dictionary<string, object>)nestedObject[key];
                           }

                           // Set the value of the last part of the column name to the current value
                           nestedObject[columnNameSegments.Last()] = !string.IsNullOrEmpty(value.ToString()) ? value : (columnsWithNumericData.Contains(columnNameSegments.Last()) ? 0 : "");
                       }

                   }

                   // Add the current row's data to the result list
                   result.Add(rowDictionary);
               }
               var documents = new Dictionary<string, object>();
               documents["documents"] = result;
               // Serialize the result list to a JSON string and return it
               return JsonConvert.SerializeObject(documents, Formatting.None);
           }*/


            public static string ProcessJson(string json)
        {
            var jObject = JObject.Parse(json);
            var documents = (JArray)jObject["documents"];

            foreach (var document in documents)
            {
                var invoiceLines = (JObject)document["invoiceLines"];
                var invoiceLinesArray = invoiceLines.Properties().Select(p => p.Value).OfType<JArray>().FirstOrDefault();

                if (invoiceLinesArray != null)
                {
                    var newInvoiceLines = new JArray();

                    foreach (var invoiceLine in invoiceLinesArray)
                    {
                        newInvoiceLines.Add(invoiceLine);
                    }

                    invoiceLines.Replace(newInvoiceLines);
                }
            }

            return jObject.ToString();
        }



        public static bool ValidateExcelSheetColumns(string filePath)
        {
            var requiredColumns = new[] {
        "issuer.address.branchID",
        "issuer.address.country",
        "issuer.address.governate",
        "issuer.address.regionCity",
        "issuer.address.street",
        "issuer.address.buildingNumber",
        "issuer.address.postalCode",
        "issuer.address.floor",
        "issuer.address.room",
        "issuer.address.landmark",
        "issuer.address.additionalInformation",
        "issuer.type",
        "issuer.id",
        "issuer.name",
        "receiver.address.country",
        "receiver.address.governate",
        "receiver.address.regionCity",
        "receiver.address.street",
        "receiver.address.buildingNumber",
        "receiver.address.postalCode",
        "receiver.address.floor",
        "receiver.address.room",
        "receiver.address.landmark",
        "receiver.address.additionalInformation",
        "receiver.type",
        "receiver.id",
        "receiver.name",
        "documentType",
        "documentTypeVersion",
        "dateTimeIssued",
        "taxpayerActivityCode",
        "internalID",
        "purchaseOrderReference",
        "purchaseOrderDescription",
        "salesOrderReference",
        "salesOrderDescription",
        "proformaInvoiceNumber",
        "payment.bankName",
        "payment.bankAddress",
        "payment.bankAccountNo",
        "payment.bankAccountIBAN",
        "payment.swiftCode",
        "payment.terms",
        "delivery.approach",
        "delivery.packaging",
        "delivery.dateValidity",
        "delivery.exportPort",
        "delivery.grossWeight",
        "delivery.netWeight",
        "delivery.terms",
        "invoiceLines.description",
        "invoiceLines.itemType",
        "invoiceLines.itemCode",
        "invoiceLines.unitType",
        "invoiceLines.quantity",
        "invoiceLines.internalCode",
        "invoiceLines.salesTotal",
        "invoiceLines.total",
        "invoiceLines.valueDifference",
        "invoiceLines.totalTaxableFees",
        "invoiceLines.netTotal",
        "invoiceLines.itemsDiscount",
        "invoiceLines.unitValue.currencySold",
        "invoiceLines.unitValue.amountEGP",
        "invoiceLines.discount.rate",
        "invoiceLines.discount.amount",
        "invoiceLines.taxableItems.taxType",
        "invoiceLines.taxableItems.amount",
        "invoiceLines.taxableItems.subType",
        "invoiceLines.taxableItems.rate",
        "totalDiscountAmount",
        "totalSalesAmount",
        "netAmount",
        "taxTotals.taxType",
        "taxTotals.amount",
        "totalAmount",
        "extraDiscountAmount",
        "totalItemsDiscountAmount"
    };
            using (SpreadsheetDocument document = SpreadsheetDocument.Open(filePath, false))
            {
                WorkbookPart workbookPart = document.WorkbookPart;
                Sheet sheet = workbookPart.Workbook.Descendants<Sheet>().First();
                WorksheetPart worksheetPart = (WorksheetPart)workbookPart.GetPartById(sheet.Id);

                // Get the header row of the sheet
                Row headerRow = worksheetPart.Worksheet.Descendants<Row>().First();

                var columnNames = headerRow.Descendants<Cell>()
                    .Select(c => GetCellValue(c, workbookPart))
                    .ToList();

                return requiredColumns.All(columnNames.Contains);
            }
        }

        //    public static string ConvertExcelDataTableToJson(DataTable dataTable)
        //    {
        //        var result = new List<Dictionary<string, object>>();
        //        var columnsWithMultipleRows = new List<string> { "invoiceLines", "taxTotals" };

        //        var columnsWithNumericData = new[]
        //        {
        //    "grossWeight","netWeight","quantity","salesTotal","total",
        //    "valueDifference","totalTaxableFees","netTotal","itemsDiscount","amountEGP",
        //    "rate","amount","totalDiscountAmount","totalSalesAmount","netAmount","totalAmount","extraDiscountAmount","totalItemsDiscountAmount"
        //};

        //        foreach (DataRow row in dataTable.Rows)
        //        {
        //            var rowDictionary = new Dictionary<string, object>();

        //            foreach (DataColumn column in dataTable.Columns)
        //            {
        //                var value = row[column.ColumnName];
        //                var columnNameSegments = column.ColumnName.Split('.');

        //                if (columnsWithMultipleRows.Contains(columnNameSegments[0]) || (!columnsWithMultipleRows.Contains(columnNameSegments[0]) && dataTable.Columns.IndexOf(column) == 0))
        //                {
        //                    var nestedObject = rowDictionary;

        //                    for (int i = 0; i < columnNameSegments.Length - 1; i++)
        //                    {
        //                        var key = columnNameSegments[i];

        //                        if (!nestedObject.ContainsKey(key) && !columnsWithMultipleRows.Contains(columnNameSegments[0]))
        //                        {
        //                            nestedObject[key] = new Dictionary<string, object>();
        //                        }
        //                        else if (columnsWithMultipleRows.Contains(columnNameSegments[0]) && !nestedObject.ContainsKey(key))
        //                        {
        //                            nestedObject[key] = new List<Dictionary<string, object>>();
        //                        }

        //                        if (!columnsWithMultipleRows.Contains(columnNameSegments[0]))
        //                        {
        //                            nestedObject = (Dictionary<string, object>)nestedObject[key];
        //                        }
        //                        else
        //                        {
        //                            nestedObject = ((List<Dictionary<string, object>>)nestedObject[key]).LastOrDefault() ?? new Dictionary<string, object>();
        //                        }
        //                    }

        //                    var nestedKey = columnNameSegments.Last();

        //                    if (columnsWithNumericData.Contains(nestedKey))
        //                    {
        //                        nestedObject[nestedKey] = value == DBNull.Value ? 0 : Convert.ToDecimal(value);
        //                    }
        //                    else
        //                    {
        //                        nestedObject[nestedKey] = value == DBNull.Value ? string.Empty : value.ToString();
        //                    }

        //                    if (columnsWithMultipleRows.Contains(columnNameSegments[0]))
        //                    {
        //                        var list = (List<Dictionary<string, object>>)rowDictionary[columnNameSegments[0]];
        //                        if (!list.Any() || list.Last() != nestedObject)
        //                        {
        //                            list.Add(nestedObject);
        //                        }
        //                    }
        //                }
        //            }

        //            result.Add(rowDictionary);
        //        }

        //        var output = new Dictionary<string, object>
        //{
        //    { "invoice", result }
        //};

        //        return JsonConvert.SerializeObject(output, Formatting.Indented);
        //    }



        //public static string ConvertExcelDataTableToJson(DataTable dataTable)
        //{
        //    var result = new List<Dictionary<string, object>>();
        //    var columnsWithMultipleRows = new List<string> { "invoiceLines", "taxTotals" };

        //    var columnsWithNumericData = new[]
        //    {
        //            "grossWeight","netWeight","invoiceLines","quantity","salesTotal","total",
        //        "valueDifference","totalTaxableFees","netTotal","itemsDiscount","unitValue","amountEGP",
        //        "rate","amount","totalDiscountAmount","totalSalesAmount","netAmount","totalAmount","extraDiscountAmount","totalItemsDiscountAmount"
        //        };
        //    foreach (DataRow row in dataTable.Rows)
        //    {
        //        var rowDictionary = new Dictionary<string, object>();

        //        foreach (DataColumn column in dataTable.Columns)
        //        {
        //            var value = row[column.ColumnName];
        //            var columnNameSegments = column.ColumnName.Split('.');

        //            if (columnsWithMultipleRows.Contains(columnNameSegments[0]) || (!columnsWithMultipleRows.Contains(columnNameSegments[0]) && dataTable.Rows.IndexOf(row) == 0))
        //            {
        //                var nestedObject = rowDictionary;

        //                for (int i = 0; i < columnNameSegments.Length - 1; i++)
        //                {
        //                    var key = columnNameSegments[i];

        //                    if (!nestedObject.ContainsKey(key) && !columnsWithMultipleRows.Contains(columnNameSegments[0]))
        //                    {
        //                        nestedObject[key] = new Dictionary<string, object>();
        //                    }
        //                    else if (columnsWithMultipleRows.Contains(columnNameSegments[0]) && !result.Exists(dict => dict.ContainsKey(columnNameSegments[0])))
        //                    {
        //                        nestedObject[key] = new List<Dictionary<string, object>>();
        //                    }
        //                    if (!columnsWithMultipleRows.Contains(columnNameSegments[0]))
        //                    {
        //                        nestedObject = (Dictionary<string, object>)nestedObject[key];
        //                    }
        //                    else
        //                    {
        //                        nestedObject = (List<Dictionary<string, object>>())nestedObject[key];
        //                    }

        //                }
        //                nestedObject[columnNameSegments.Last()] = string.IsNullOrEmpty(value.ToString()) ? (columnsWithNumericData.Contains(columnNameSegments.Last()) ? 0 : "") : value;
        //            }
        //        }

        //        result.Add(rowDictionary);
        //    }
        //    var documents = new Dictionary<string, object>();
        //    documents["documents"] = result;
        //    return JsonConvert.SerializeObject(result, Formatting.Indented);
        //}


       

        private static string GetCellValue(Cell cell, WorkbookPart workbookPart)
        {
            if (cell == null || cell.CellValue == null)
                return null;
            string value = cell.CellValue.InnerText;
            if (cell.DataType != null && cell.DataType == CellValues.SharedString)
                value = workbookPart.SharedStringTablePart.SharedStringTable
                    .Elements<SharedStringItem>().ElementAt(int.Parse(value)).InnerText;
            return value;
        }
        public static DataTable GetExcelDataTable(string fileName)
        {
            using (var stream = File.Open(fileName, FileMode.Open, FileAccess.Read))
            {
                // Define a custom encoding provider that can handle the Arabic content
                Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);

                // Create the Excel reader with the custom encoding
                var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration
                {
                    FallbackEncoding = Encoding.GetEncoding("windows-1256")
                });

                // Read the Excel data into a DataSet
                var dataSet = reader.AsDataSet(new ExcelDataSetConfiguration
                {
                    ConfigureDataTable = _ => new ExcelDataTableConfiguration
                    {
                        UseHeaderRow = true
                    }
                });
                // Return the first DataTable in the DataSet
                return dataSet.Tables[0];
            }
        }
    }
}