using AutoMapper;
using EInvoice.Application.EInvoice.DTOs;
using EInvoice.Common.Configurations;
using EInvoice.Services;
using Microsoft.Extensions.Options;
using Newtonsoft.Json;
using Newtonsoft.Json.Serialization;
using System.Net.Http.Headers;
using System.Net.Mime;
using System.Text;

namespace EInvoice.API.Services;

public interface IInvoiceService
{
    Task<string> CreateInvoice(DocumentRequestDto invoiceRequest);
}

public class InvoiceService : IInvoiceService
{
    private readonly IAuthService _authService;

    private readonly HttpClient _httpClient;
    private readonly ISignerService _signerService;
    private readonly EInvoicingConfiguration _eInvoicingConfiguration;
    private readonly IMapper _mapper;

    public InvoiceService(IAuthService authService, HttpClient httpClient, ISignerService signerService, IOptions<EInvoicingConfiguration> eInvoicingConfiguration, IMapper mapper)
    {
        _authService = authService;
        _httpClient = httpClient;
        _signerService = signerService;
        _mapper = mapper;
        _eInvoicingConfiguration = eInvoicingConfiguration.Value;
    }
    public async Task<string> CreateInvoice(DocumentRequestDto invoiceRequest)
    {
        var accessToken = await _authService.GetAccessToken();
        _httpClient.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Bearer", accessToken);

        var headers = new Dictionary<string, string>
        {
        };

        var serializer = new DocumentSerializationService();

        var invoice = invoiceRequest.Documents[0];

        var jsonSettings = new JsonSerializerSettings()
        {
            FloatFormatHandling = FloatFormatHandling.String,
            FloatParseHandling = FloatParseHandling.Decimal,
            DateFormatHandling = DateFormatHandling.IsoDateFormat,
            DateParseHandling = DateParseHandling.None,
            NullValueHandling = NullValueHandling.Ignore,
            ContractResolver = new CamelCasePropertyNamesContractResolver(),

        };

        var content = JsonConvert.SerializeObject(invoice, jsonSettings);

        var canonicalString = serializer.Serialize(content, jsonSettings);

        var invoicesForSubmit = new DocumentRequestForSubmitDto();

        var invoiceForSubmit = _mapper.Map<DocumentRequestDocumentForSubmitDto>(invoice);

        if (invoice.DocumentTypeVersion != "0.9")
        {
            var base64Signature = _signerService.SignWithCMS(canonicalString);

            invoiceForSubmit.Signatures = new List<Signature>(){
                new()
            {
                SignatureType = "I",
                Value = base64Signature
            }};
        }

        var request = new HttpRequestMessage(HttpMethod.Post, $"{_eInvoicingConfiguration.BaseUrl}/{_eInvoicingConfiguration.SubmitUrl}");

        invoicesForSubmit.Documents.Add(invoiceForSubmit);

        var serializedContent = JsonConvert.SerializeObject(invoicesForSubmit, jsonSettings);

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