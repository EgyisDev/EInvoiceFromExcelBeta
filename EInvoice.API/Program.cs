using EInvoice.API.Services;
using EInvoice.Common.Configurations;
using EInvoice.Services;
using Net.Pkcs11Interop.HighLevelAPI.Factories;

var builder = WebApplication.CreateBuilder(args);

// Add services to the container.

builder.Services.AddControllers();
// Learn more about configuring Swagger/OpenAPI at https://aka.ms/aspnetcore/swashbuckle
builder.Services.AddEndpointsApiExplorer();
builder.Services.AddSwaggerGen();
builder.Services.AddMemoryCache();

//builder.Services.AddSingleton<IToolkitHandler, ToolkitHandler>();

builder.Services.Configure<AuthConfiguration>(builder.Configuration.GetSection("Auth"));
builder.Services.Configure<ApplicationConfiguration>(builder.Configuration.GetSection("Application"));
builder.Services.Configure<EInvoicingConfiguration>(builder.Configuration.GetSection("EInvoicing"));

//var toolKitConfig = builder.Configuration.GetSection("ToolkitConfig").Get<ToolkitConfig>();


//builder.Services.AddToolkit(builder.Configuration, toolKitConfig);

builder.Services.AddScoped<IInvoiceService, InvoiceService>();

builder.Services.AddScoped<IAuthService, AuthService>();
builder.Services.AddScoped<IPkcs11LibraryFactory, Pkcs11LibraryFactory>();
builder.Services.AddScoped<ISignerService, SignerService>();


builder.Services.AddHttpClient();

var app = builder.Build();

// Configure the HTTP request pipeline.
if (app.Environment.IsDevelopment())
{
    app.UseSwagger();
    app.UseSwaggerUI();
}

//app.MigrateToolkitLocalStorage();

app.UseHttpsRedirection();

app.UseAuthorization();

app.MapControllers();

app.Run();
