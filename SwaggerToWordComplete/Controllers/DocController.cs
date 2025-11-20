using SwaggerToWordComplete.Models;
using SwaggerToWordComplete.Services;
using Microsoft.AspNetCore.Mvc;

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
        return View(); // Views/Doc/Index.cshtml
    }

    [HttpPost]
    public async Task<IActionResult> Generate([FromForm] DocGenerationRequest request)
    {
        if (request.Swagger == null || request.Swagger.Length == 0)
        {
            TempData["Error"] = "please select Swagger (JSON) .";
            return RedirectToAction("Index");
        }

        try
        {
            var docBytes = await _service.GenerateDocxAsync(request);
            return File(docBytes, "application/vnd.openxmlformats-officedocument.wordprocessingml.document", "API_Documentation.docx");
        }
        catch (Exception ex)
        {
            // در محیط تولید لاگ کنید
            TempData["Error"] = "Error in Generate: " + ex.Message;
            return RedirectToAction("Index");
        }
    }
}
