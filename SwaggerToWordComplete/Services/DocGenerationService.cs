using SwaggerToWordComplete.Helpers;
using SwaggerToWordComplete.Models;
using DocumentFormat.OpenXml.Packaging;
using Microsoft.OpenApi.Models;
using Microsoft.OpenApi.Readers;


namespace SwaggerToWordComplete.Services;

public class DocGenerationService
{
    public DocGenerationService()
    {
    }

    public async Task<byte[]> GenerateDocxAsync(DocGenerationRequest request)
    {
        // read OpenAPI document
        OpenApiDocument doc;
        var reader = new OpenApiStreamReader();
        using (var ms = new MemoryStream())
        {
            await request.Swagger.CopyToAsync(ms);
            ms.Position = 0;
            doc = reader.Read(ms, out var diag);
        }

        // map request -> DocSettings
        var settings = MapSettings(request);

        // build docx into memory
        using var mem = new MemoryStream();
        using (var word = WordprocessingDocument.Create(mem, DocumentFormat.OpenXml.WordprocessingDocumentType.Document, true))
        {
            WordBuilder.Build(word, doc, settings, request);
        }

        return mem.ToArray();
    }

    private DocSettings MapSettings(DocGenerationRequest r)
    {
        bool isRtl = (r.Rtl ?? "rtl").ToLower() == "rtl";
        string NormalizeColor(string color)
        {
            if (string.IsNullOrWhiteSpace(color)) return "000000";
            if (color.StartsWith("#")) color = color[1..];
            if (color.Length == 3) color = string.Concat(color[0], color[0], color[1], color[1], color[2], color[2]);
            return color.Length == 6 ? "#" + color : "#000000";
        }

        return new DocSettings
        {
            Font = string.IsNullOrWhiteSpace(r.Font) ? "Tahoma" : r.Font,
            IsRtl = isRtl,
            H1Size = r.H1 <= 0 ? 28 : r.H1,
            H2Size = r.H2 <= 0 ? 24 : r.H2,
            H3Size = r.H3 <= 0 ? 22 : r.H3,
            BodySize = r.BodySize <= 0 ? 12 : r.BodySize,
            HeadColor = NormalizeColor(r.HeadColor),
            BorderSize = r.BorderSize <= 0 ? 8 : r.BorderSize,
            BorderColor = NormalizeColor(r.BorderColor),
            MarginTop = r.MTop <= 0 ? 720 : r.MTop,
            MarginBottom = r.MBottom <= 0 ? 720 : r.MBottom,
            MarginLeft = r.MLeft <= 0 ? 720 : r.MLeft,
            MarginRight = r.MRight <= 0 ? 720 : r.MRight,
            IncludeTOC = r.Toc,
            IncludePageNumbers = r.PageNumber,


            Intro = string.IsNullOrWhiteSpace(r.Intro) ? "مقدمه مستند" : r.Intro,
            SharedModel = string.IsNullOrWhiteSpace(r.SharedResponseModel) ? "" : r.SharedResponseModel,
            SharedErrors = string.IsNullOrWhiteSpace(r.SharedErrors) ? "" : r.SharedErrors,


            TitleAll = string.IsNullOrWhiteSpace(r.Title_all) ? "اطلاعات کلی" : r.Title_all,
            Title = string.IsNullOrWhiteSpace(r.Title) ? "API Documentation" : r.Title,


            IndexTitle = string.IsNullOrWhiteSpace(r.IndexTitle) ? "فهرست متدها (Index)" : r.IndexTitle,
            ChangelogTitle = string.IsNullOrWhiteSpace(r.ChangelogTitle) ? "لیست تغییرات (Change Log)" : r.ChangelogTitle,


            TitleTableOfContents = string.IsNullOrWhiteSpace(r.TitleTableOfContents) ? "فهرست مطالب" : r.TitleTableOfContents,
            TitleIntro = string.IsNullOrWhiteSpace(r.TitleIntro) ? "مقدمه مستند" : r.TitleIntro,
            TitleSharedResponseModel = string.IsNullOrWhiteSpace(r.TitleSharedResponseModel) ? "" : r.TitleSharedResponseModel,
            TitleSharedErrors = string.IsNullOrWhiteSpace(r.TitleSharedErrors) ? "" : r.TitleSharedErrors,


            AuthRequired = r.AuthRequired ? false : r.AuthRequired,
            AuthHeader = string.IsNullOrWhiteSpace(r.TitleSharedErrors) ? "" : r.AuthHeader,
            AuthValue = string.IsNullOrWhiteSpace(r.TitleSharedErrors) ? "" : r.AuthValue,
       
        



        };
    }
}

