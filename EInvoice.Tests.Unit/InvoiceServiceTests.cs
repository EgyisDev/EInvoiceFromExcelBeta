using System.Text.Json;
using EInvoice.API.Services;
using EInvoice.Application.EInvoice.DTOs;
using EInvoice.Services;
using FluentAssertions;

namespace EInvoice.Tests.Unit
{
    public class InvoiceServiceTests
    {
        [Fact]
        public void Test1()
        {
            var service = new DocumentSerializationService();

            var fileToDeserializePath = Path.Combine(Directory.GetCurrentDirectory(), "InvoiceExampleToSerialize.json");

            var content = File.ReadAllText(fileToDeserializePath);

            var result = service.Serialize(content);

            result.Should().NotBeNullOrEmpty();

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "SerializedInvoiceExpectedResult.txt");

            var fileContent = File.ReadAllText(filePath);

            result.Should().Be(fileContent);
        }


        //[Fact]
        //public void ShouldGenerateSignature()
        //{
        //    var service = new DocumentSerializationService();

        //    var fileToDeserializePath = Path.Combine(Directory.GetCurrentDirectory(), "InvoiceWithoutSignature.json");

        //    var content = File.ReadAllText(fileToDeserializePath);

        //    var invoiceRequestDto = JsonSerializer.Deserialize<DocumentRequestDocumentForSubmitDto>(content);

        //    var resultList = new DocumentRequestDto();

        //    resultList.Documents.Add(invoiceRequestDto);

        //    var cononical = service.Serialize(content);

        //    var signerService = new SignerService();

        //  var base64Signature =  signerService.SignWithCMS2(cononical);

        //  invoiceRequestDto.Signatures = new List<Signature>
        //  {
        //      new()
        //      {
        //          SignatureType = "I",
        //          Value = base64Signature
        //      }
        //  };

        //  //invoiceRequestDto.Documents[0].Signatures = new List<Signature>
        //  //{
        //  //    new()
        //  //    {
        //  //        SignatureType = "I",
        //  //        Value = base64Signature
        //  //    }
        //  //};

        //    var result = JsonSerializer.Serialize(invoiceRequestDto);

        //    result.Should().NotBeNullOrEmpty();

        //    var filePath = Path.Combine(Directory.GetCurrentDirectory(), "InvoiceWithSignature.json");

        //    var fileContent = File.ReadAllText(filePath);

        //    result.Should().Be(fileContent);
        //}
    }
}