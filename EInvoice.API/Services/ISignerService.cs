namespace EInvoice.Services;

public interface ISignerService
{
    string SignWithCMS(string serializedJson);
}