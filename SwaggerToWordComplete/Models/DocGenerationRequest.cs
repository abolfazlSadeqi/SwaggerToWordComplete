namespace SwaggerToWordComplete.Models;

public class DocGenerationRequest
{
    // file input name "swagger"
    public IFormFile Swagger { get; set; }

    // form fields
    public string Changelog { get; set; }
    public bool IncludeExamples { get; set; } = true;
    public bool AuthRequired { get; set; } = true;
    public string AuthHeader { get; set; } = "Authorization";
    public string AuthValue { get; set; } = "Bearer {TOKEN}";

    public string Author { get; set; }
    public string For { get; set; }
    public string Intro { get; set; }
    public string SharedResponseModel { get; set; }
    public string SharedErrors { get; set; }
    public string Methods { get; set; }

    // UI settings
    public int H1 { get; set; } = 18;
    public int H2 { get; set; } = 16;
    public int H3 { get; set; } = 14;
    public int BodySize { get; set; } = 12;
    public string HeadColor { get; set; } = "#000000";

    public string Font { get; set; } = "Tahoma";
    public string Rtl { get; set; } = "rtl";
    public string Title { get; set; } = "API Documentation";
    public string Title_all { get; set; } = "اطلاعات کلی";
    public string IndexTitle { get; set; } = "فهرست متدها (Index)";
    public string ChangelogTitle { get; set; } = "لیست تغییرات (Change Log)";

    // margins etc.
    public int MTop { get; set; } = 720;
    public int MBottom { get; set; } = 720;
    public int MLeft { get; set; } = 720;
    public int MRight { get; set; } = 720;
    public int BorderSize { get; set; } = 8;
    public string BorderColor { get; set; } = "#2B579A";
    public bool Toc { get; set; } = true;
    public bool PageNumber { get; set; } = true;



    public string TitleTableOfContents { get; init; } = "فهرست";
    public string TitleIntro { get; init; } = "مقدمه مستند";
    public string TitleSharedResponseModel { get; init; } = "مدل خروجی مشترک";
    public string TitleSharedErrors { get; init; } = "مدل خروجی خطا";
}
