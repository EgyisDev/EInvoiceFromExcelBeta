﻿using System.Text.Json;
using System.Text;
using ThinInvoiceUploaderClient;

ExcelHelper.ConvertExcelToDocumentRequestDto();
var invoiceUploader = "https://localhost:7072/InvoiceController";






var client = new HttpClient();

//foreach (var employee in employees)
//{
//    var request = new HttpRequestMessage(HttpMethod.Post, payrollCalculatorUrl);
//    var employeeAdapter = new PayrollSystemEmployeeAdapter(employee);
//    request.Content = new StringContent(JsonSerializer.Serialize(employeeAdapter), Encoding.UTF8, "application/json");

//    var response = await client.SendAsync(request);
//    var responseJson = await response.Content.ReadAsStringAsync();
//    var salary = decimal.Parse(responseJson);

//    Console.WriteLine($"Salary for employee `{employee.FullName}` as of today = {salary}");
//}
Console.ReadKey();

