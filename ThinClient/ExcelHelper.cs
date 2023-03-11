using DocumentFormat.OpenXml.Drawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using ExcelDataReader;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.IO.Packaging;
using System.Linq;
using System.Text;
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
            var jsonDocument = ConvertExcelDataTableToJson(dt);
            return JsonConvert.DeserializeObject<DocumentRequestDocumentForSerializeDto>(jsonDocument);
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

        public static string ConvertExcelDataTableToJson(DataTable dataTable)
        {
            var result = new List<Dictionary<string, object>>();

            foreach (DataRow row in dataTable.Rows)
            {
                var item = new Dictionary<string, object>();

                foreach (DataColumn column in dataTable.Columns)
                {
                    var value = row[column.ColumnName];

                    var keys = column.ColumnName.Split('.');
                    var nestedObject = item;

                    for (int i = 0; i < keys.Length - 1; i++)
                    {
                        var key = keys[i];
                        if (!nestedObject.ContainsKey(key))
                        {
                            nestedObject[key] = new Dictionary<string, object>();
                        }
                        nestedObject = (Dictionary<string, object>)nestedObject[key];
                    }

                    nestedObject[keys.Last()] = value;
                }

                result.Add(item);
            }

            return JsonConvert.SerializeObject(result);
        }

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