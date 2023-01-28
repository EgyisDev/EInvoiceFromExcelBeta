using EInvoice.API.Services;
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

            var result = service.Serialize(content, null);

            result.Should().NotBeNullOrEmpty();

            var filePath = Path.Combine(Directory.GetCurrentDirectory(), "SerializedInvoiceExpectedResult.txt");

            var fileContent = File.ReadAllText(filePath);

            result.Should().Be(fileContent);
        }
    }
}