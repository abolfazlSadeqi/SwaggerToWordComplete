using Microsoft.AspNetCore.Mvc;
using Serilog;
using SwaggerToWordComplete.Models;
using SwaggerToWordComplete.Services;

namespace SwaggerToWordComplete.Controllers;

public class DocController : Controller
{
    private readonly DocGenerationService _service;

    public DocController(DocGenerationService service)
    {
        _service = service;
    }

    [HttpGet]
    public IActionResult Index()
    {
        Log.Information("GET /Doc/Index called.");

        return View(); // Views/Doc/Index.cshtml
    }

    [HttpPost]
    public async Task<IActionResult> Generate([FromForm] DocGenerationRequest request)
    {
        Log.Information("POST /Doc/Generate called.");

        if (request.Swagger == null || request.Swagger.Length == 0)
        {
            Log.Warning("Generate called without Swagger file.");

            TempData["Error"] = "please select Swagger (JSON) .";
            return RedirectToAction("Index");
        }

        try
        {
            var docBytes = await _service.GenerateDocxAsync(request);
            Log.Information("Successfully generated documentation. Output size: {Size} bytes ,Filename={FileName}", docBytes.Length, request.Swagger.FileName);

            return File(docBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "API_Documentation.docx");
        }
        catch (Exception ex)
        {
            Log.Error(ex, "Error occurred in Generate()");

            // در محیط تولید لاگ کنید
            TempData["Error"] = "Error in Generate: " + ex.Message;
            return RedirectToAction("Index");
        }
    }
}
