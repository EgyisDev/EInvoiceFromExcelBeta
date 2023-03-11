using ExcelDataReader;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ThinInvoiceUploaderClient
{
    internal class ExcelHelper
    {
        public static DocumentRequestDto ConvertExcelToDocumentRequestDto()
        {
            string filePath = "InvoiceSample.xlsx";
            // Load Excel file
            System.Text.Encoding.RegisterProvider(System.Text.CodePagesEncodingProvider.Instance);
            var documentRequestDto = new DocumentRequestDto();

            using (var stream = File.Open(filePath, FileMode.Open, FileAccess.Read))
            {
                using (var reader = ExcelReaderFactory.CreateReader(stream))
                {
                    while (reader.Read())
                    {
                        var document = new DocumentRequestDocumentForSerializeDto();
                        document.Issuer = new Issuer();
                        document.Receiver = new Receiver();
                        document.Payment = new Payment();
                        document.Delivery = new DocumentRequestDeliveryDto();
                        document.InvoiceLines = new List<InvoiceLine>();
                        document.TaxTotals = new List<TaxTotal>();

                        document.Issuer.Name = reader.GetString(0);
                        document.Issuer.Id = reader.GetString(1);
                        document.Issuer.Address = new DocumentRequestAddressDto()
                        {
                            AdditionalInformation = reader.GetString(2),
                            BranchID = reader.GetString(3),
                            BuildingNumber = reader.GetString(4),
                            Country = reader.GetString(5),
                            Floor = reader.GetString(6),
                            Governate = reader.GetString(7),
                            Landmark = reader.GetString(8),
                            PostalCode = reader.GetString(9),
                            Street = reader.GetString(10),
                            RegionCity = reader.GetString(11),
                            Room = reader.GetString(12),
                        };

                        document.Receiver.Name = reader.GetString(3);
                        document.Receiver.Id = reader.GetString(4);
                        document.Receiver.Address = new DocumentRequestAddressDto()
                        {
                            AdditionalInformation = reader.GetString(2),
                            BranchID = reader.GetString(3),
                            BuildingNumber = reader.GetString(4),
                            Country = reader.GetString(5),
                            Floor = reader.GetString(6),
                            Governate = reader.GetString(7),
                            Landmark = reader.GetString(8),
                            PostalCode = reader.GetString(9),
                            Street = reader.GetString(10),
                            RegionCity = reader.GetString(11),
                            Room = reader.GetString(12),
                        };

                        document.DocumentType = reader.GetString(6);
                        document.DocumentTypeVersion = reader.GetString(7);
                        document.DateTimeIssued = reader.GetDateTime(8);
                        document.TaxpayerActivityCode = reader.GetString(9);
                        document.InternalID = reader.GetString(10);
                        document.PurchaseOrderReference = reader.GetString(11);
                        document.PurchaseOrderDescription = reader.GetString(12);
                        document.SalesOrderReference = reader.GetString(13);
                        document.SalesOrderDescription = reader.GetString(14);
                        document.ProformaInvoiceNumber = reader.GetString(15);

                        document.Payment = new Payment()
                        {
                            BankAccountIBAN = reader.GetString(16),
                            BankAccountNo = reader.GetString(17),
                            BankAddress = reader.GetString(18),
                            BankName = reader.GetString(19),
                            SwiftCode = reader.GetString(20),
                            Terms = reader.GetString(21),
                        };

                        document.Delivery = new DocumentRequestDeliveryDto()
                        {
                            Approach = reader.GetString(22),
                            DateValidity = reader.GetDateTime(23),
                            ExportPort = reader.GetString(24),
                            GrossWeight = reader.GetDouble(25),
                            NetWeight = reader.GetDouble(26),
                            Packaging = reader.GetString(27),
                            Terms = reader.GetString(22),
                        };

                        while (reader.Name != "Invoice Lines")
                        {
                            reader.Read();
                        }

                        reader.Read(); //Skip the header row of Invoice Lines

                        while (reader.Name == "Invoice Lines")
                        {
                            var line = new InvoiceLine();

                            line.Description = reader.GetString(1);
                            line.ItemType = reader.GetString(2);
                            line.ItemCode = reader.GetString(3);
                            line.UnitType = reader.GetString(4);
                            line.Quantity = reader.GetInt32(5);
                            line.InternalCode = reader.GetString(6);
                            line.SalesTotal = reader.GetDecimal(7);
                            line.Total = reader.GetDecimal(8);
                            line.ValueDifference = reader.GetDecimal(9);
                            line.TotalTaxableFees = reader.GetDecimal(10);
                            line.NetTotal = reader.GetDecimal(11);
                            line.ItemsDiscount = reader.GetDecimal(12);
                            line.UnitValue = new UnitValue();
                            line.UnitValue.AmountEGP = reader.GetDouble(13);
                            line.UnitValue.CurrencySold = reader.GetString(14);
                            line.Discount = new DocumentRequestDiscountDto
                            {
                                Rate = reader.GetInt32(15),
                                Amount = reader.GetDecimal(16)
                            };

                            document.InvoiceLines.Add(line);

                            if (!reader.Read())
                            {
                                break;
                            }
                        }

                        while (reader.Name != "Total")
                        {
                            reader.Read();
                        }

                        document.TotalDiscountAmount = reader.GetDecimal(1);
                        document.TotalSalesAmount = reader.GetDouble(1);
                        documentRequestDto.Documents.Add(document);
                    }
                }
            }
            return documentRequestDto;
        }
    }
}