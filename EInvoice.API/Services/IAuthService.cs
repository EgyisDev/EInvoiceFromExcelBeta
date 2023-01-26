namespace EInvoice.API.Services;

public interface IAuthService
{
    Task<string> GetAccessToken();
}