using EInvoice.API.Services;
using EInvoice.Application.EInvoice.DTOs;
using Microsoft.AspNetCore.Mvc;

namespace EInvoice.API.Controllers;

[Route("api/[controller]")]
[ApiController]
public class InvoiceController : ControllerBase
{
    private readonly IInvoiceService _invoiceService;

    public InvoiceController(IInvoiceService invoiceService)
    {
        _invoiceService = invoiceService;
    }

    [HttpPost(Name = "invoice")]
    public async Task<ActionResult> Post(DocumentRequestDto documentRequestDto)
    {
        var result = await _invoiceService.CreateInvoice(documentRequestDto);

        return Ok(result);
    }
}