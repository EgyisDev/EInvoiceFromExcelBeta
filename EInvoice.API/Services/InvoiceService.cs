using EInvoice.Application.EInvoice.DTOs;
using EInvoice.Services;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Text;
using System.Text.Json;
using EInvoice.Common.Configurations;
using Microsoft.Extensions.Options;

namespace EInvoice.API.Services;

public interface IInvoiceService
{
    Task<string> CreateInvoice(DocumentRequestDto invoiceRequestDto);
}

public class InvoiceService : IInvoiceService
{
    private readonly IAuthService _authService;

    //private readonly IToolkitHandler _toolkitHandler;
    //private ISignerService _signerService;

    private readonly HttpClient _httpClient;
    private readonly ISignerService _signerService;
    private readonly EInvoicingConfiguration _eInvoicingConfiguration;

    public InvoiceService(IAuthService authService, HttpClient httpClient, ISignerService signerService, IOptions<EInvoicingConfiguration> eInvoicingConfiguration)
    {
        _authService = authService;
        _httpClient = httpClient;
        _signerService = signerService;
        _eInvoicingConfiguration = eInvoicingConfiguration.Value;
        //_signerService = signerService;
    }
    public async Task<string> CreateInvoice(DocumentRequestDto invoiceRequestDto)
    {
        var accessToken = await _authService.GetAccessToken();
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        var headers = new Dictionary<string, string>
        {
            // { "User-agent", "egyis-erp" },
            // //{ "User-agent", "Testing Gateway" },
            //{ "Accept", "**/**" },
            //{ "Accept-Encoding", "gzip, deflate, br" },
            //{ "Connection", "keep-alive" },
            //{ "cache-control", "no-cache" }
        };


        //var result = await _httpClient.PostAsJsonAsync("https://api.preprod.invoicing.eta.gov.eg/api/v1.0/documentsubmissions", invoiceRequestDto);



        var serializer = new DocumentSerializationService();

        var content = JsonSerializer.Serialize(invoiceRequestDto);

        var serializedResult = serializer.Serialize(content);

        //var signerService  = new SignerService();

        if (invoiceRequestDto.Documents[0].DocumentTypeVersion != "0.9")
        {
            var base64Signature = _signerService.SignWithCMS(serializedResult);

            invoiceRequestDto.Documents[0].Signatures = new List<Signature>
            {
                new()
                {
                    SignatureType = "CMS",
                    Value = base64Signature
                }
            };
        }

        var request = new HttpRequestMessage(HttpMethod.Post, $"{_eInvoicingConfiguration.BaseUrl}/{_eInvoicingConfiguration.SubmitUrl}");

        var serializedContent = JsonSerializer.Serialize(invoiceRequestDto, new JsonSerializerOptions());

        request.Content = new StringContent(serializedContent, Encoding.UTF8, MediaTypeNames.Application.Json);

        foreach (var header in headers)
        {
            request.Headers.Add(header.Key, header.Value);
        }

        var response = await _httpClient.SendAsync(request);

        if (!response.IsSuccessStatusCode)
            throw new Exception(await response.Content.ReadAsStringAsync());

        var responseContent = await response.Content.ReadAsStringAsync();

        return responseContent;
    }
}