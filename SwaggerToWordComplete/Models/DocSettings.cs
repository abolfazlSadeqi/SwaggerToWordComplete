namespace SwaggerToWordComplete.Models;


public record DocSettings
{
    public string Font { get; init; } = "Tahoma";
    public bool IsRtl { get; init; } = true;
    public int H1Size { get; init; } = 28;
    public int H2Size { get; init; } = 24;
    public int H3Size { get; init; } = 22;
    public int BodySize { get; init; } = 22;
    public string HeadColor { get; init; } = "#000000";
    public int BorderSize { get; init; } = 8;
    public string BorderColor { get; init; } = "2B579A";
    public int MarginTop { get; init; } = 720;
    public int MarginBottom { get; init; } = 720;
    public int MarginLeft { get; init; } = 720;
    public int MarginRight { get; init; } = 720;
    public bool IncludeTOC { get; init; } = true;
    public bool IncludePageNumbers { get; init; } = true;

    public string Title { get; init; } = "API Documentation";
    public string Intro { get; init; } = "مقدمه مستند";
    public string SharedModel { get; init; } = "";
    public string SharedErrors { get; init; } = "";
    public string IndexTitle { get; init; } = "فهرست متدها (Index)";
    public string ChangelogTitle { get; init; } = "لیست تغییرات (Change Log)";
    public string TitleAll { get; init; } = "اطلاعات کلی";


    public string TitleTableOfContents { get; init; } = "فهرست";
    public string TitleIntro { get; init; } = "مقدمه مستند";
    public string TitleSharedResponseModel { get; init; } = "مدل خروجی مشترک";

    public string TitleSharedErrors { get; init; } = "مدل خروجی خطا";


    public bool AuthRequired { get; set; }
    public string AuthHeader { get; set; } = "Authorization";
    public string AuthValue { get; set; } = "Bearer {TOKEN}";


}

