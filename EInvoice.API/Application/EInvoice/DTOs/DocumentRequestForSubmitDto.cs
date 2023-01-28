using System.Text.Json.Serialization;

namespace EInvoice.Application.EInvoice.DTOs;

public class DocumentRequestDto
{
    public DocumentRequestDto()
    {
        Documents = new List<DocumentRequestDocumentForSerializeDto>();
    }

    [JsonPropertyName("documents")]
    public List<DocumentRequestDocumentForSerializeDto> Documents { get; set; }
}

public class DocumentRequestForSubmitDto
{
    public DocumentRequestForSubmitDto()
    {
        Documents = new List<DocumentRequestDocumentForSubmitDto>();
    }

    [JsonPropertyName("documents")]
    public List<DocumentRequestDocumentForSubmitDto> Documents { get; set; }
}

public class DocumentRequestAddressDto
{
    [JsonPropertyName("branchID")]
    public string? BranchID { get; set; }

    [JsonPropertyName("country")]
    public string Country { get; set; }

    [JsonPropertyName("governate")]
    public string Governate { get; set; }

    [JsonPropertyName("regionCity")]
    public string RegionCity { get; set; }

    [JsonPropertyName("street")]
    public string Street { get; set; }

    [JsonPropertyName("buildingNumber")]
    public string BuildingNumber { get; set; }

    [JsonPropertyName("postalCode")]
    public string PostalCode { get; set; }

    [JsonPropertyName("floor")]
    public string Floor { get; set; }

    [JsonPropertyName("room")]
    public string Room { get; set; }

    [JsonPropertyName("landmark")]
    public string Landmark { get; set; }

    [JsonPropertyName("additionalInformation")]
    public string AdditionalInformation { get; set; }
}

public class DocumentRequestDeliveryDto
{
    [JsonPropertyName("approach")]
    public string Approach { get; set; }

    [JsonPropertyName("packaging")]
    public string Packaging { get; set; }

    [JsonPropertyName("dateValidity")]
    public DateTime DateValidity { get; set; }

    [JsonPropertyName("exportPort")]
    public string ExportPort { get; set; }

    [JsonPropertyName("grossWeight")]
    public double GrossWeight { get; set; }

    [JsonPropertyName("netWeight")]
    public double NetWeight { get; set; }

    [JsonPropertyName("terms")]
    public string Terms { get; set; }
}

public class DocumentRequestDiscountDto
{
    [JsonPropertyName("rate")]
    public int Rate { get; set; }

    [JsonPropertyName("amount")]
    public decimal Amount { get; set; }
}

public class DocumentRequestDocumentForSubmitDto
{
    public DocumentRequestDocumentForSubmitDto()
    {
        Signatures = new List<Signature>();
    }
    
    [JsonPropertyName("issuer")]
    public Issuer Issuer { get; set; }

    [JsonPropertyName("receiver")]
    public Receiver Receiver { get; set; }

    [JsonPropertyName("documentType")]
    public string DocumentType { get; set; }

    [JsonPropertyName("documentTypeVersion")]
    public string DocumentTypeVersion { get; set; }

    [JsonPropertyName("dateTimeIssued")]
    public DateTime DateTimeIssued { get; set; }

    [JsonPropertyName("taxpayerActivityCode")]
    public string TaxpayerActivityCode { get; set; }

    [JsonPropertyName("internalID")]
    public string InternalID { get; set; }

    [JsonPropertyName("purchaseOrderReference")]
    public string PurchaseOrderReference { get; set; }

    [JsonPropertyName("purchaseOrderDescription")]
    public string PurchaseOrderDescription { get; set; }

    [JsonPropertyName("salesOrderReference")]
    public string SalesOrderReference { get; set; }

    [JsonPropertyName("salesOrderDescription")]
    public string SalesOrderDescription { get; set; }

    [JsonPropertyName("proformaInvoiceNumber")]
    public string ProformaInvoiceNumber { get; set; }

    [JsonPropertyName("payment")]
    public Payment Payment { get; set; }

    [JsonPropertyName("delivery")]
    public DocumentRequestDeliveryDto Delivery { get; set; }

    [JsonPropertyName("invoiceLines")]
    public List<InvoiceLine> InvoiceLines { get; set; }

    [JsonPropertyName("totalDiscountAmount")]
    public decimal TotalDiscountAmount { get; set; }

    [JsonPropertyName("totalSalesAmount")]
    public double TotalSalesAmount { get; set; }

    [JsonPropertyName("netAmount")]
    public double NetAmount { get; set; }

    [JsonPropertyName("taxTotals")]
    public List<TaxTotal> TaxTotals { get; set; }

    [JsonPropertyName("totalAmount")]
    public double TotalAmount { get; set; }

    [JsonPropertyName("extraDiscountAmount")]
    public decimal ExtraDiscountAmount { get; set; }

    [JsonPropertyName("totalItemsDiscountAmount")]
    public decimal TotalItemsDiscountAmount { get; set; }

    [JsonPropertyName("signatures")]
    public List<Signature>? Signatures { get; set; } = new List<Signature>();
}


public class DocumentRequestDocumentForSerializeDto
{
    [JsonPropertyName("issuer")]
    public Issuer Issuer { get; set; }

    [JsonPropertyName("receiver")]
    public Receiver Receiver { get; set; }

    [JsonPropertyName("documentType")]
    public string DocumentType { get; set; }

    [JsonPropertyName("documentTypeVersion")]
    public string DocumentTypeVersion { get; set; }

    [JsonPropertyName("dateTimeIssued")]
    public DateTime DateTimeIssued { get; set; }

    [JsonPropertyName("taxpayerActivityCode")]
    public string TaxpayerActivityCode { get; set; }

    [JsonPropertyName("internalID")]
    public string InternalID { get; set; }

    [JsonPropertyName("purchaseOrderReference")]
    public string PurchaseOrderReference { get; set; }

    [JsonPropertyName("purchaseOrderDescription")]
    public string PurchaseOrderDescription { get; set; }

    [JsonPropertyName("salesOrderReference")]
    public string SalesOrderReference { get; set; }

    [JsonPropertyName("salesOrderDescription")]
    public string SalesOrderDescription { get; set; }

    [JsonPropertyName("proformaInvoiceNumber")]
    public string ProformaInvoiceNumber { get; set; }

    [JsonPropertyName("payment")]
    public Payment Payment { get; set; }

    [JsonPropertyName("delivery")]
    public DocumentRequestDeliveryDto Delivery { get; set; }

    [JsonPropertyName("invoiceLines")]
    public List<InvoiceLine> InvoiceLines { get; set; }

    [JsonPropertyName("totalDiscountAmount")]
    public decimal TotalDiscountAmount { get; set; }

    [JsonPropertyName("totalSalesAmount")]
    public double TotalSalesAmount { get; set; }

    [JsonPropertyName("netAmount")]
    public double NetAmount { get; set; }

    [JsonPropertyName("taxTotals")]
    public List<TaxTotal> TaxTotals { get; set; }

    [JsonPropertyName("totalAmount")]
    public double TotalAmount { get; set; }

    [JsonPropertyName("extraDiscountAmount")]
    public decimal ExtraDiscountAmount { get; set; }

    [JsonPropertyName("totalItemsDiscountAmount")]
    public decimal TotalItemsDiscountAmount { get; set; }
}


public class InvoiceLine
{
    [JsonPropertyName("description")]
    public string Description { get; set; }

    [JsonPropertyName("itemType")]
    public string ItemType { get; set; }

    [JsonPropertyName("itemCode")]
    public string ItemCode { get; set; }

    [JsonPropertyName("unitType")]
    public string UnitType { get; set; }

    [JsonPropertyName("quantity")]
    public int Quantity { get; set; }

    [JsonPropertyName("internalCode")]
    public string InternalCode { get; set; }

    [JsonPropertyName("salesTotal")]
    public decimal SalesTotal { get; set; }

    [JsonPropertyName("total")]
    public decimal Total { get; set; }

    [JsonPropertyName("valueDifference")]
    public decimal ValueDifference { get; set; }

    [JsonPropertyName("totalTaxableFees")]
    public decimal TotalTaxableFees { get; set; }

    [JsonPropertyName("netTotal")]
    public decimal NetTotal { get; set; }

    [JsonPropertyName("itemsDiscount")]
    public decimal ItemsDiscount { get; set; }

    [JsonPropertyName("unitValue")]
    public UnitValue UnitValue { get; set; }

    [JsonPropertyName("discount")]
    public DocumentRequestDiscountDto Discount { get; set; }

    [JsonPropertyName("taxableItems")]
    public List<TaxableItem> TaxableItems { get; set; }
}

public class Issuer
{
    [JsonPropertyName("address")]
    public DocumentRequestAddressDto Address { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; }

    [JsonPropertyName("id")]
    public string Id { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; }
}

public class Payment
{
    [JsonPropertyName("bankName")]
    public string BankName { get; set; }

    [JsonPropertyName("bankAddress")]
    public string BankAddress { get; set; }

    [JsonPropertyName("bankAccountNo")]
    public string BankAccountNo { get; set; }

    [JsonPropertyName("bankAccountIBAN")]
    public string BankAccountIBAN { get; set; }

    [JsonPropertyName("swiftCode")]
    public string SwiftCode { get; set; }

    [JsonPropertyName("terms")]
    public string Terms { get; set; }
}

public class Receiver
{
    [JsonPropertyName("address")]
    public DocumentRequestAddressDto Address { get; set; }

    [JsonPropertyName("type")]
    public string Type { get; set; }

    [JsonPropertyName("id")]
    public string Id { get; set; }

    [JsonPropertyName("name")]
    public string Name { get; set; }
}


public class Signature
{
    [JsonPropertyName("signatureType")]
    public string SignatureType { get; set; }

    [JsonPropertyName("value")]
    public string Value { get; set; }
}

public class TaxableItem
{
    [JsonPropertyName("taxType")]
    public string TaxType { get; set; }

    [JsonPropertyName("amount")]
    public decimal Amount { get; set; }

    [JsonPropertyName("subType")]
    public string SubType { get; set; }

    [JsonPropertyName("rate")]
    public decimal Rate { get; set; }
}

public class TaxTotal
{
    [JsonPropertyName("taxType")]
    public string TaxType { get; set; }

    [JsonPropertyName("amount")]
    public decimal Amount { get; set; }
}

public class UnitValue
{
    [JsonPropertyName("currencySold")]
    public string CurrencySold { get; set; }

    [JsonPropertyName("amountEGP")]
    public double AmountEGP { get; set; }
}
