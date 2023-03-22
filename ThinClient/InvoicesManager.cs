using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Net;
using DocumentFormat.OpenXml.Office.Word;
using DocumentFormat.OpenXml.Office2016.Drawing.ChartDrawing;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Vml;
using Newtonsoft.Json;

namespace ThinInvoiceUploaderClient
{
    internal class InvoicesManager
    {
        public string ExtractDocumentRequests()
        {
            var fileName = "InvoiceSample4.xlsx";
            var documents = new List<DocumentRequestDocumentForSerializeDto>();

            using (var spreadsheetDocument = SpreadsheetDocument.Open(fileName, false))
            {
                //TODO: Get Issuer data
                var issuer = GetIssuer(spreadsheetDocument);
                if (issuer == null)
                {
                    Console.ForegroundColor = ConsoleColor.Red;
                    Console.WriteLine("Error! No Issuer data found! ");   
                }
                //Docs
                var docsWorksheetPart = GetWorksheetPartByName(spreadsheetDocument, "Docs");
                var docsWorksheet = docsWorksheetPart.Worksheet;
                var docsSheetData = docsWorksheet.GetFirstChild<SheetData>();

                foreach (var row in docsSheetData.Elements<Row>().Skip(1))
                {
                    if (string.IsNullOrWhiteSpace(GetCellValue(spreadsheetDocument, row.ChildElements[11] as Cell)))
                    {
                        break;
                    }
                    //Receiver
                    var receiverAddressCountry = GetCellValue(spreadsheetDocument, row.ChildElements[0] as Cell);
                    var receiverAddressGovernate = GetCellValue(spreadsheetDocument, row.ChildElements[1] as Cell);
                    var receiverAddressRegionCity = GetCellValue(spreadsheetDocument, row.ChildElements[2] as Cell);
                    var receiverAddressStreet = GetCellValue(spreadsheetDocument, row.ChildElements[3] as Cell);
                    var receiverAddressBuildingNumber = GetCellValue(spreadsheetDocument, row.ChildElements[4] as Cell);
                    var receiverAddressPostalCode = GetCellValue(spreadsheetDocument, row.ChildElements[5] as Cell);
                    var receiverAddressFloor = GetCellValue(spreadsheetDocument, row.ChildElements[6] as Cell);
                    var receiverAddressRoom = GetCellValue(spreadsheetDocument, row.ChildElements[7] as Cell);
                    var receiverAddressLandmark = GetCellValue(spreadsheetDocument, row.ChildElements[8] as Cell);
                    var receiverAddressAdditionalInformation = GetCellValue(spreadsheetDocument, row.ChildElements[9] as Cell);
                    var receiverType = GetCellValue(spreadsheetDocument, row.ChildElements[10] as Cell);
                    var receiverId = GetCellValue(spreadsheetDocument, row.ChildElements[11] as Cell);
                    var receiverName = GetCellValue(spreadsheetDocument, row.ChildElements[12] as Cell);
                    //Document
                    var documentType = GetCellValue(spreadsheetDocument, row.ChildElements[13] as Cell);
                    var documentTypeVersion = GetCellValue(spreadsheetDocument, row.ChildElements[14] as Cell);
                    var dateTimeIssued = DateTime.Parse(GetCellValue(spreadsheetDocument, row.ChildElements[15] as Cell));
                    var taxpayerActivityCode = GetCellValue(spreadsheetDocument, row.ChildElements[16] as Cell);
                    var internalID = GetCellValue(spreadsheetDocument, row.ChildElements[17] as Cell);
                    var purchaseOrderReference = GetCellValue(spreadsheetDocument, row.ChildElements[18] as Cell);
                    var purchaseOrderDescription = GetCellValue(spreadsheetDocument, row.ChildElements[19] as Cell);
                    var salesOrderReference = GetCellValue(spreadsheetDocument, row.ChildElements[20] as Cell);
                    var salesOrderDescription = GetCellValue(spreadsheetDocument, row.ChildElements[21] as Cell);
                    var proformaInvoiceNumber = GetCellValue(spreadsheetDocument, row.ChildElements[22] as Cell);
                    //Payment
                    var paymentBankName = GetCellValue(spreadsheetDocument, row.ChildElements[23] as Cell);
                    var paymentBankAddress = GetCellValue(spreadsheetDocument, row.ChildElements[24] as Cell);
                    var paymentBankAccountNumber = GetCellValue(spreadsheetDocument, row.ChildElements[25] as Cell);
                    var paymentBankAccountIBAN = GetCellValue(spreadsheetDocument, row.ChildElements[26] as Cell);
                    var paymentBankSwiftCode = GetCellValue(spreadsheetDocument, row.ChildElements[27] as Cell);
                    var paymentTerms = GetCellValue(spreadsheetDocument, row.ChildElements[28] as Cell);
                    //Delivery
                    var deliveryApproach = GetCellValue(spreadsheetDocument, row.ChildElements[29] as Cell);
                    var deliveryPackaging = GetCellValue(spreadsheetDocument, row.ChildElements[30] as Cell);
                    var deliveryDateValidity = Convert.ToDateTime(GetCellValue(spreadsheetDocument, row.ChildElements[31] as Cell));
                    var deliveryExportPort = GetCellValue(spreadsheetDocument, row.ChildElements[32] as Cell);
                    var deliveryGrossWeight = Convert.ToDouble(GetCellValue(spreadsheetDocument, row.ChildElements[33] as Cell));
                    var deliveryNetWeight = Convert.ToDouble(GetCellValue(spreadsheetDocument, row.ChildElements[34] as Cell));
                    var deliveryTerms = GetCellValue(spreadsheetDocument, row.ChildElements[35] as Cell);
                    //InvoiceLines
                    var invoiceLineDescription = GetCellValue(spreadsheetDocument, row.ChildElements[36] as Cell);
                    var invoiceLineItemType = GetCellValue(spreadsheetDocument, row.ChildElements[37] as Cell);
                    var invoiceLineItemCode = GetCellValue(spreadsheetDocument, row.ChildElements[38] as Cell);
                    var invoiceLineUnitType = GetCellValue(spreadsheetDocument, row.ChildElements[39] as Cell);
                    var invoiceLineQuantity = (int)Math.Floor(Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[40] as Cell)));
                    var invoiceLineInternalCode = GetCellValue(spreadsheetDocument, row.ChildElements[41] as Cell);
                    var invoiceLineSalesTotal = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[42] as Cell));
                    var invoiceLineTotal = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[43] as Cell));
                    var invoiceLineValueDifference = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[44] as Cell));
                    var invoiceLineTotalTaxableFees = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[45] as Cell));
                    var invoiceLineNetTotal = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[46] as Cell));
                    var invoiceLineItemsDiscount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[47] as Cell));
                    //InvoiceLines.UnitValue
                    var unitValueCurrencySold = GetCellValue(spreadsheetDocument, row.ChildElements[48] as Cell);
                    var unitValueAmountEGP = Convert.ToDouble(GetCellValue(spreadsheetDocument, row.ChildElements[49] as Cell));

                    //InvoiceLines.Discount
                    var discountRate = (int)Math.Floor(Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[50] as Cell)));
                    var discountAmount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[51] as Cell));

                    //InvoiceLines.TaxableItems
                    var taxableItemsTaxType = GetCellValue(spreadsheetDocument, row.ChildElements[52] as Cell);
                    var taxableItemsAmount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[53] as Cell));
                    var taxableSubType = GetCellValue(spreadsheetDocument, row.ChildElements[54] as Cell);
                    var taxableRate = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[55] as Cell));
                    var taxableItems = new List<TaxableItem>
                    {
                        new TaxableItem
                        {
                            TaxType = taxableItemsTaxType,
                            Amount = taxableItemsAmount,
                            SubType = taxableSubType,
                            Rate = taxableRate
                        }
                    };
                    //Continue: document
                    decimal documentTotalDiscountAmount = 0;
                    double documentTotalSalesAmount = 0;
                    double documentNetAmount = 0;
                    if (!string.IsNullOrWhiteSpace(GetCellValue(spreadsheetDocument, row.ChildElements[56] as Cell)))
                    {
                        documentTotalDiscountAmount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[56] as Cell));
                        documentTotalSalesAmount = Convert.ToDouble(GetCellValue(spreadsheetDocument, row.ChildElements[57] as Cell));
                        documentNetAmount = Convert.ToDouble(GetCellValue(spreadsheetDocument, row.ChildElements[58] as Cell));
                    }
                    //TaxTotals
                    string taxTotalsTaxType = "";
                    decimal taxTotalsAmount = 0;
                    if (!string.IsNullOrWhiteSpace(GetCellValue(spreadsheetDocument, row.ChildElements[59] as Cell)))
                    {
                        taxTotalsTaxType = GetCellValue(spreadsheetDocument, row.ChildElements[59] as Cell);
                        taxTotalsAmount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[60] as Cell));
                    }
                    //Continue: document
                    double docuemntTotalAmount = 0;
                    decimal docuemntExtraDiscountAmount = 0;
                    decimal docuemntTotalItemsDiscountAmount = 0;
                    if (!string.IsNullOrWhiteSpace(GetCellValue(spreadsheetDocument, row.ChildElements[61] as Cell)))
                    {
                        docuemntTotalAmount = Convert.ToDouble(GetCellValue(spreadsheetDocument, row.ChildElements[61] as Cell));
                        docuemntExtraDiscountAmount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[62] as Cell));
                        docuemntTotalItemsDiscountAmount = Convert.ToDecimal(GetCellValue(spreadsheetDocument, row.ChildElements[63] as Cell));
                    }

                    var invoiceLine = new InvoiceLine
                    {
                        Description = invoiceLineDescription,
                        ItemType = invoiceLineItemType,
                        ItemCode = invoiceLineItemCode,
                        UnitType = invoiceLineUnitType,
                        Quantity = invoiceLineQuantity,
                        InternalCode = invoiceLineInternalCode,
                        SalesTotal = invoiceLineSalesTotal,
                        Total = invoiceLineTotal,
                        ValueDifference = invoiceLineValueDifference,
                        TotalTaxableFees = invoiceLineTotalTaxableFees,
                        NetTotal = invoiceLineNetTotal,
                        ItemsDiscount = invoiceLineItemsDiscount,
                        UnitValue = new UnitValue
                        {
                            CurrencySold = unitValueCurrencySold,
                            AmountEGP = unitValueAmountEGP
                        },
                        Discount = new DocumentRequestDiscountDto
                        {
                            Rate = discountRate,
                            Amount = discountAmount
                        },
                        TaxableItems = taxableItems
                    };

                    var taxTotal = new TaxTotal
                    {
                        Amount = taxTotalsAmount,
                        TaxType = taxTotalsTaxType
                    };

                    var document = documents.FirstOrDefault(inv => inv.InternalID == internalID);

                    if (document == null)
                    {
                        document = new DocumentRequestDocumentForSerializeDto
                        {
                            DateTimeIssued = dateTimeIssued,
                            InternalID = internalID,
                            Delivery = new DocumentRequestDeliveryDto()
                            {
                                Approach = deliveryApproach,
                                DateValidity = deliveryDateValidity,
                                ExportPort = deliveryExportPort,
                                GrossWeight = deliveryGrossWeight,
                                NetWeight = deliveryNetWeight,
                                Packaging = deliveryPackaging,
                                Terms = deliveryTerms
                            },
                            DocumentType = documentType,
                            DocumentTypeVersion = documentTypeVersion,
                            ExtraDiscountAmount = docuemntExtraDiscountAmount,
                            NetAmount = documentNetAmount,
                            ProformaInvoiceNumber = proformaInvoiceNumber,
                            PurchaseOrderDescription = purchaseOrderDescription,
                            PurchaseOrderReference = purchaseOrderReference,
                            SalesOrderDescription = salesOrderDescription,
                            SalesOrderReference = salesOrderReference,
                            TaxpayerActivityCode = taxpayerActivityCode,
                            TotalAmount = docuemntTotalAmount,
                            TotalDiscountAmount = documentTotalDiscountAmount,
                            TotalItemsDiscountAmount = docuemntTotalItemsDiscountAmount,
                            TotalSalesAmount = documentTotalSalesAmount,
                            InvoiceLines = new List<InvoiceLine>(),
                            TaxTotals = new List<TaxTotal>(),
                            Payment = new Payment()
                            {
                                BankAccountIBAN = paymentBankAccountIBAN,
                                BankAccountNo = paymentBankAccountNumber,
                                BankAddress = paymentBankAddress,
                                BankName = paymentBankName,
                                SwiftCode = paymentBankSwiftCode,
                                Terms = paymentTerms,
                            },
                            Receiver = new Receiver()
                            {
                                Address = new DocumentRequestAddressDto()
                                {
                                    Country = receiverAddressCountry,
                                    Governate = receiverAddressGovernate,
                                    RegionCity = receiverAddressRegionCity,
                                    Street = receiverAddressStreet,
                                    BuildingNumber = receiverAddressBuildingNumber,
                                    PostalCode = receiverAddressPostalCode,
                                    Floor = receiverAddressFloor,
                                    Room = receiverAddressRoom,
                                    Landmark = receiverAddressLandmark,
                                    AdditionalInformation = receiverAddressAdditionalInformation
                                },
                                Type = receiverType,
                                Id = receiverId,
                                Name = receiverName,
                            }
                        };

                        documents.Add(document);
                    }

                    document.InvoiceLines.Add(invoiceLine);
                    document.TaxTotals.Add(taxTotal);
                    document.Issuer = issuer;
                }
            }
            var json = JsonConvert.SerializeObject(documents, Formatting.Indented);
            Console.WriteLine(json);
            return json;
        }
        private Issuer GetIssuer(SpreadsheetDocument spreadsheetDocument)
        {
            Issuer issuer = new Issuer();

            //Docs
            var issuerWorksheetPart = GetWorksheetPartByName(spreadsheetDocument, "Issuer");
            var issuerWorksheet = issuerWorksheetPart.Worksheet;
            var issuerSheetData = issuerWorksheet.GetFirstChild<SheetData>();

            foreach (var row in issuerSheetData.Elements<Row>().Skip(1))
            {
                if (string.IsNullOrWhiteSpace(GetCellValue(spreadsheetDocument, row.ChildElements[12] as Cell)))
                {
                    break;
                }
                //issuer address
                var issuerAddressBranchId = GetCellValue(spreadsheetDocument, row.ChildElements[0] as Cell);
                var issuerAddressCountry = GetCellValue(spreadsheetDocument, row.ChildElements[1] as Cell);
                var issuerAddressGovernate = GetCellValue(spreadsheetDocument, row.ChildElements[2] as Cell);
                var issuerAddressRegionCity = GetCellValue(spreadsheetDocument, row.ChildElements[3] as Cell);
                var issuerAddressStreet = GetCellValue(spreadsheetDocument, row.ChildElements[4] as Cell);
                var issuerAddressBuildingNumber = GetCellValue(spreadsheetDocument, row.ChildElements[5] as Cell);
                var issuerAddressPostalCode = GetCellValue(spreadsheetDocument, row.ChildElements[6] as Cell);
                var issuerAddressFloor = GetCellValue(spreadsheetDocument, row.ChildElements[7] as Cell);
                var issuerAddressRoom = GetCellValue(spreadsheetDocument, row.ChildElements[8] as Cell);
                var issuerAddressLandmark = GetCellValue(spreadsheetDocument, row.ChildElements[9] as Cell);
                var issuerAddressAdditionalInformation = GetCellValue(spreadsheetDocument, row.ChildElements[10] as Cell);
                var issuerType = GetCellValue(spreadsheetDocument, row.ChildElements[11] as Cell);
                var issuerId = GetCellValue(spreadsheetDocument, row.ChildElements[12] as Cell);
                var issuerName = GetCellValue(spreadsheetDocument, row.ChildElements[13] as Cell);

                if (string.IsNullOrWhiteSpace(issuer.Id))
                {
                    issuer = new Issuer()
                    {
                        Id = issuerId,
                        Name = issuerName,
                        Type = issuerType,
                        Address = new DocumentRequestAddressDto()
                        {
                            BranchID = issuerAddressBranchId,
                            Country = issuerAddressCountry,
                            Governate = issuerAddressGovernate,
                            RegionCity = issuerAddressRegionCity,
                            Street = issuerAddressStreet,
                            BuildingNumber = issuerAddressBuildingNumber,
                            PostalCode = issuerAddressPostalCode,
                            Floor = issuerAddressFloor,
                            Room = issuerAddressRoom,
                            Landmark = issuerAddressLandmark,
                            AdditionalInformation = issuerAddressAdditionalInformation,
                        },
                    };
                    break;
                }
            }
            return issuer;
        }
        private static WorksheetPart GetWorksheetPartByName(SpreadsheetDocument document, string sheetName)
        {
            var sheets = document.WorkbookPart.Workbook.Descendants<Sheet>().Where(s => s.Name == sheetName);

            if (sheets.Count() == 0)
            {
                throw new ArgumentException($"The specified sheet ({sheetName}) does not exist.");
            }

            var sheet = sheets.First();
            var relationshipId = sheet.Id.Value;
            var worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);

            return worksheetPart;
        }

        private static string GetCellValue(SpreadsheetDocument document, Cell cell)
        {
            var value = cell.InnerText;

            if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
            {
                var sharedStringTable = document.WorkbookPart.SharedStringTablePart.SharedStringTable;
                value = sharedStringTable.ElementAt(int.Parse(value)).InnerText;
            }

            return value;
        }
    }
    public class Invoice
    {
        public string InvoiceNumber { get; set; }
        public List<InvoiceLine> InvoiceLines { get; set; }
    }

    public class Item
    {
        public string ItemDesc { get; set; }
        public decimal ItemQty { get; set; }
        public decimal ItemPrice { get; set; }
    }
}
