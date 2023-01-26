namespace EInvoice.Common.Configurations;

public class AuthConfiguration
{
    public string ClientId { get; set; }
    public string ClientSecret { get; set; }
    public string Scope { get; set; }
    public string GrantType { get; set; }
    public string TokenUrl { get; set; }

}