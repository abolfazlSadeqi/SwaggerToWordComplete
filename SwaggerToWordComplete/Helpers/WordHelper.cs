using SwaggerToWordComplete.Models;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Microsoft.OpenApi.Models;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using System.Text;

namespace SwaggerToWordComplete.Helpers;

public static class WordBuilder
{
    public static void Build(WordprocessingDocument word, OpenApiDocument doc, DocSettings settings, object requestObj)
    {
        var main = word.AddMainDocumentPart();
        main.Document = new Document(new Body());
        var body = main.Document.Body;

        // styles and numbering
        AddStylesAndNumbering(main, settings);

        // Title
        AddTitle(body, settings.Title, settings);

        // TOC
        if (settings.IncludeTOC)
            InsertTableOfContents(body, settings);

        // Info page
        AddHeadingWithNumbering(body, settings.TitleAll, 1, settings);
        var infoTable = new List<List<string>>
        {
            new(){"عنوان",  "-"},
            new(){"Base URL", doc.Servers?.FirstOrDefault()?.Url ?? "-"},
            new(){"نسخه", doc.Info?.Version ?? "-"},
            new(){"API Type", "REST / JSON"},
            new(){"Rate limiting", "-"},
            new(){"کاربرد", "-"},
            new(){"تاریخ انتشار", DateTime.UtcNow.ToString("yyyy-MM-dd")},
            new(){"محدودیت ها", "-"}
        };
        AddTable(body, infoTable, settings, true);
        AddPageBreak(body);

        // Intro
        AddHeadingWithNumbering(body, settings.TitleIntro, 1, settings);
        AddParagraph(body, settings.Intro, settings);
        AddPageBreak(body);

        // Index
        AddHeadingWithNumbering(body, settings.IndexTitle, 1, settings);
        var indexRows = new List<List<string>> { new() { "آدرس", "متد", "خلاصه", "توضیحات" } };
        foreach (var p in doc.Paths)
            foreach (var op in p.Value.Operations)
            {
                indexRows.Add(new List<string> {
                    p.Key,
                    op.Key.ToString().ToUpper(),
                    op.Value.Summary ?? "-",
                    op.Value.Description ?? "-"
                });
            }
        AddTable(body, indexRows, settings, true);
        AddPageBreak(body);

        // Changelog
        AddHeadingWithNumbering(body, settings.ChangelogTitle, 1, settings);
        // if request had changelog, try to render — requestObj may be DocGenerationRequest
        if (requestObj is SwaggerToWordComplete.Models.DocGenerationRequest req && !string.IsNullOrWhiteSpace(req.Changelog))
            AddChangeLogSection(body, req.Changelog, settings);
        else
            AddParagraph(body, "-", settings);
        AddPageBreak(body);

        // Shared model & errors
        AddHeadingWithNumbering(body, settings.TitleSharedResponseModel, 1, settings);
        AddParagraph(body, settings.SharedModel ?? "-", settings);
        AddPageBreak(body);

        AddHeadingWithNumbering(body, settings.TitleSharedErrors, 1, settings);
        AddParagraph(body, settings.SharedErrors ?? "-", settings);
        AddPageBreak(body);

        // Endpoints
        AddHeadingWithNumbering(body, "Endpoints", 1, settings);
        foreach (var p in doc.Paths)
        {
            foreach (var op in p.Value.Operations)
            {
                var method = op.Key.ToString().ToUpper();
                var operation = op.Value;

                AddHeadingWithNumbering(body, p.Key, 2, settings);
                AddParagraph(body, "توضیحات :", settings);
                AddParagraph(body, operation.Description ?? "-", settings);
                bool AuthRequiredType = operation.Security?.Any() == true ? true : settings.AuthRequired ? true : false;
                var generalRows = new List<List<string>>
                {
                    new(){"فیلد", "توضیح"},
                    new(){"مسیر (URL)", p.Key},
                    new(){"HTTP Method", method},
                   new(){"نیاز به توکن", AuthRequiredType ? "بله" : "خیر"},
                    new(){"کاربرد", operation.Summary ?? "-"}
                };
                AddTable(body, generalRows, settings, true);

                AddParametersTables(body, operation, doc.Components, settings);
                AddResponsesTables(body, operation, doc.Components, settings);

                AddHeadingWithNumbering(body, "Curl:", 3, settings);
                AddCodeBlock(body, BuildCurl(p.Key, method, operation,true, doc.Components, settings.AuthRequired,settings.AuthHeader, settings.AuthValue), settings);

                AddPageBreak(body);
            }
        }

        // page border and margins
        AddPageBorder(body, settings);

        if (settings.IncludePageNumbers)
            AddPageNumbers(main, settings);

        main.Document.Save();
    }

    #region Styles & Helpers (simplified copies of your functions)

    static void AddStylesAndNumbering(MainDocumentPart main, DocSettings settings)
    {
        var numberingPart = main.AddNewPart<NumberingDefinitionsPart>();
        numberingPart.Numbering = new Numbering(
            new AbstractNum(
                new Level(new NumberingFormat() { Val = NumberFormatValues.Decimal }, new LevelText() { Val = "%1." }) { LevelIndex = 0 },
                new Level(new NumberingFormat() { Val = NumberFormatValues.Decimal }, new LevelText() { Val = "%1.%2." }) { LevelIndex = 1 }
            )
            { AbstractNumberId = 1 },
            new NumberingInstance(new AbstractNumId() { Val = 1 }) { NumberID = 1 }
        );

        var stylePart = main.AddNewPart<StyleDefinitionsPart>();
        stylePart.Styles = new Styles();
        stylePart.Styles.Append(
            CreateHeadingStyle("Heading1", "Heading 1", 0, settings.H1Size, settings),
            CreateHeadingStyle("Heading2", "Heading 2", 1, settings.H2Size, settings),
            CreateHeadingStyle("Heading3", "Heading 3", 2, settings.H3Size, settings),
            CreateNormalStyle("Normal", settings)
        );
    }

    static Style CreateHeadingStyle(string styleId, string name, int outline, int size, DocSettings settings)
    {
        var fontSizeVal = (size * 2).ToString();
        var runProps = new StyleRunProperties(new Bold(), new RunFonts() { Ascii = settings.Font }, new FontSize() { Val = fontSizeVal });
        runProps.Append(new Color() { Val = settings.HeadColor?.TrimStart('#') ?? "000000" });

        return new Style(
            new StyleName() { Val = name },
            new BasedOn() { Val = "Normal" },
            new NextParagraphStyle() { Val = "Normal" },
            new UIPriority() { Val = 9 },
            new UnhideWhenUsed(),
            new PrimaryStyle(),
            runProps,
            new StyleParagraphProperties(new OutlineLevel() { Val = outline })
        )
        { Type = StyleValues.Paragraph, StyleId = styleId };
    }

    static Style CreateNormalStyle(string styleId, DocSettings settings)
    {
        var runProps = new StyleRunProperties(new RunFonts() { Ascii = settings.Font }, new FontSize() { Val = (settings.BodySize * 2).ToString() });
        return new Style(new StyleName() { Val = "Normal" }, new UIPriority() { Val = 1 }, new PrimaryStyle(), runProps) { Type = StyleValues.Paragraph, StyleId = styleId };
    }

    static void AddTitle(Body body, string text, DocSettings settings)
    {
        var p = new Paragraph();
        var pp = new ParagraphProperties();
        if (settings.IsRtl) pp.Append(new BiDi());
        pp.Append(new ParagraphStyleId() { Val = "Title" });
        p.Append(pp);

        var r = new Run();
        r.Append(new RunFonts { Ascii = settings.Font });
        r.Append(new FontSize { Val = (settings.BodySize * 2).ToString() });
        r.Append(new Text(text ?? ""));
        p.Append(r);
        body.Append(p);
    }

    static void AddParagraph(Body body, string text, DocSettings settings)
    {
        var p = new Paragraph();
        var pp = new ParagraphProperties();
        if (settings.IsRtl) { pp.Append(new BiDi()); pp.Append(new Justification { Val = JustificationValues.Left }); }
        else pp.Append(new Justification { Val = JustificationValues.Right });
        p.Append(pp);

        var r = new Run();
        r.Append(new RunFonts { Ascii = settings.Font });
        r.Append(new FontSize { Val = (settings.BodySize * 2).ToString() });
        r.Append(new Text(text ?? ""));
        p.Append(r);
        body.Append(p);
    }

    static void AddHeadingWithNumbering(Body body, string text, int level, DocSettings settings)
    {
        var p = new Paragraph();
        var pp = new ParagraphProperties();
        if (settings.IsRtl) { pp.Append(new BiDi()); pp.Append(new Justification { Val = JustificationValues.Left }); }
        else pp.Append(new Justification { Val = JustificationValues.Right });

        pp.Append(new ParagraphStyleId() { Val = "Heading" + level });
        p.Append(pp);

        var r = new Run();
        int size = level == 1 ? settings.H1Size : level == 2 ? settings.H2Size : settings.H3Size;
        var rp = new RunProperties(new Bold(), new RunFonts { Ascii = settings.Font }, new FontSize { Val = (size * 2).ToString() });
        rp.Append(new Color { Val = settings.HeadColor?.TrimStart('#') ?? "000000" });
        r.Append(rp);
        r.Append(new Text(text ?? ""));
        p.Append(r);
        body.Append(p);
    }

    static void AddCodeBlock(Body body, string code, DocSettings settings)
    {
        var p = new Paragraph();
        var pp = new ParagraphProperties();
        if (settings.IsRtl) { pp.Append(new BiDi()); pp.Append(new Justification { Val = JustificationValues.Left }); }
        else pp.Append(new Justification { Val = JustificationValues.Right });
        p.Append(pp);

        var r = new Run();
        var rp = new RunProperties(new RunFonts { Ascii = "Consolas" }, new FontSize { Val = (Math.Max(10, settings.BodySize - 2) * 2).ToString() });
        r.Append(rp);
        r.Append(new Text(code ?? ""));
        p.Append(r);
        body.Append(p);
    }

    static void AddPageBreak(Body body)
    {
        body.Append(new Paragraph(new Run(new Break() { Type = BreakValues.Page })));
    }

    static string BuildCurl(string path, string method, OpenApiOperation operation, bool includeExamples, OpenApiComponents components, bool authRequired, string authHeader, string authValue)
    {

        // Replace path parameters (e.g. {id})
        string fullUrl = "[baseUrl]" + path;
        if (operation.Parameters != null)
        {
            foreach (var p in operation.Parameters.Where(x => x.In == ParameterLocation.Path))
            {
                string sample = GetDefaultValue(p.Schema) ?? "1";
                fullUrl = fullUrl.Replace("{" + p.Name + "}", sample);
            }
        }

        // Query parameters
        var queryParams = operation.Parameters?.Where(x => x.In == ParameterLocation.Query).ToList();
        if (queryParams?.Any() == true)
        {
            string q = string.Join("&", queryParams
                .Select(p => $"{p.Name}={GetDefaultValue(p.Schema)}"));
            fullUrl += "?" + q;
        }

        var sb = new StringBuilder();
        sb.Append($"curl -X {method.ToUpper()} \"{fullUrl}\"");

        // Header parameters
        var headers = operation.Parameters?.Where(x => x.In == ParameterLocation.Header).ToList();
        if (headers?.Any() == true || authRequired)
        {
            foreach (var h in headers)
            {
                string val = GetDefaultValue(h.Schema) ?? "value";
                sb.Append($" -H \"{h.Name}: {val}\"");
            }

        }

        // Authorization header if required
        if (authRequired)
        {
            sb.Append($" -H \"{authHeader}: {authValue}\"");
        }


        // Request body section
        if (operation.RequestBody != null)
        {
            var obj = operation.RequestBody.Content.FirstOrDefault();
            var mediaType = obj.Key;
            var content = obj.Value;

            sb.Append($" -H \"Content-Type: {mediaType}\"");

            string rawBody = null;

            // Example first priority
            if (content.Example != null)
            {
                rawBody = ConvertIOpenApiAnyToString(content.Example);
            }
            else if (content.Examples != null && content.Examples.Count > 0)
            {
                var first = content.Examples.First().Value.Value;
                rawBody = ConvertIOpenApiAnyToString(first);
            }
            else
            {
                // Build mock JSON from schema
                var schema = ResolveSchemaReference(content.Schema, components);
                if (schema != null)
                    rawBody = GenerateSampleJson(schema);
            }

            if (!string.IsNullOrWhiteSpace(rawBody))
            {
                rawBody = rawBody.Replace("\"", "\\\""); // escape quotes
                sb.Append($" -d \"{rawBody}\"");
            }
        }

        return sb.ToString().Trim();
    }
    static void AddTable(Body body, List<List<string>> rows, DocSettings settings, bool reverseColumns = true)
    {
        var table = new Table();
        var tblProps = new TableProperties(new TableJustification() { Val = settings.IsRtl ? TableRowAlignmentValues.Left : TableRowAlignmentValues.Right }, new TableWidth() { Type = TableWidthUnitValues.Pct, Width = "5000" },
            new TableBorders(
                new TopBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize) },
                new BottomBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize) },
                new RightBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize) },
                new LeftBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize) },
                new InsideHorizontalBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize) },
                new InsideVerticalBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize) }
            ));
        table.AppendChild(tblProps);

        int colCount = rows.FirstOrDefault()?.Count ?? 1;
        int colWidth = 5000 / colCount;

        for (int i = 0; i < rows.Count; i++)
        {
            var tr = new TableRow();
            IEnumerable<string> cells = reverseColumns ? rows[i].AsEnumerable().Reverse() : rows[i];

            foreach (var cellText in cells)
            {
                var tc = new TableCell();
                var p = new Paragraph();
                var pp = new ParagraphProperties();
                if (settings.IsRtl) { pp.Append(new BiDi()); pp.Append(new Justification { Val = JustificationValues.Left }); }
                else pp.Append(new Justification { Val = JustificationValues.Right });
                p.Append(pp);

                var r = new Run();
                var rp = new RunProperties(new RunFonts { Ascii = settings.Font }, new FontSize { Val = (settings.BodySize * 2).ToString() });

                if (i == 0)
                {
                    rp.Append(new Bold());
                    rp.Append(new Color { Val = "FFFFFF" });
                    tc.Append(new TableCellProperties(new Shading { Val = ShadingPatternValues.Clear, Color = "auto", Fill = settings.BorderColor.TrimStart('#') }, new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = colWidth.ToString() }));
                }
                else
                {
                    tc.Append(new TableCellProperties(new TableCellWidth { Type = TableWidthUnitValues.Pct, Width = colWidth.ToString() }));
                }

                r.Append(rp);
                r.Append(new Text(cellText ?? "-"));
                p.Append(r);
                tc.Append(p);
                tr.Append(tc);
            }

            table.Append(tr);
        }

        body.Append(table);
    }

    static void AddPageBorder(Body body, DocSettings settings)
    {
        var sectionProps = body.Elements<SectionProperties>().FirstOrDefault();
        if (sectionProps == null)
        {
            sectionProps = new SectionProperties();
            body.Append(sectionProps);
        }

        var pageBorders = new PageBorders
        {
            Display = PageBorderDisplayValues.AllPages,
            ZOrder = PageBorderZOrderValues.Back,
            TopBorder = new TopBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize), Space = 4, Color = settings.BorderColor.TrimStart('#') },
            BottomBorder = new BottomBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize), Space = 4, Color = settings.BorderColor.TrimStart('#') },
            LeftBorder = new LeftBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize), Space = 4, Color = settings.BorderColor.TrimStart('#') },
            RightBorder = new RightBorder { Val = BorderValues.Single, Size = (uint)Math.Max(1, settings.BorderSize), Space = 4, Color = settings.BorderColor.TrimStart('#') }
        };

        sectionProps.RemoveAllChildren<PageBorders>();
        sectionProps.Append(pageBorders);

        var pageMargin = new PageMargin
        {
            Top = (Int32Value)settings.MarginTop,
            Bottom = (Int32Value)settings.MarginBottom,
            Left = (UInt32Value)(uint)settings.MarginLeft,
            Right = (UInt32Value)(uint)settings.MarginRight,
            Header = (UInt32Value)450U,
            Footer = (UInt32Value)450U,
            Gutter = (UInt32Value)0U
        };

        sectionProps.RemoveAllChildren<PageMargin>();
        sectionProps.Append(pageMargin);
    }

    static void InsertTableOfContents(Body body, DocSettings settings)
    {
        AddHeadingWithNumbering(body, settings.TitleTableOfContents, 1, settings);
        var tocParagraph = new Paragraph(
            new ParagraphProperties(new Justification { Val = settings.IsRtl ? JustificationValues.Right : JustificationValues.Left }, new BiDi()),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Begin }),
            new Run(new FieldCode("TOC \\o \"1-3\" \\h \\z \\u")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.Separate }),
            new Run(new Text(" Word Update .....")),
            new Run(new FieldChar { FieldCharType = FieldCharValues.End })
        );
        body.Append(tocParagraph);
        AddPageBreak(body);
    }

    static void AddPageNumbers(MainDocumentPart main, DocSettings settings)
    {
        var footerPart = main.AddNewPart<FooterPart>();
        var footer = new Footer();
        var paragraph = new Paragraph(new ParagraphProperties(new Justification() { Val = JustificationValues.Center }));
        paragraph.Append(new Run(new FieldChar() { FieldCharType = FieldCharValues.Begin }, new FieldCode(" PAGE "), new FieldChar() { FieldCharType = FieldCharValues.End }));
        footer.Append(paragraph);
        footerPart.Footer = footer;

        var sectProps = main.Document.Body.Elements<SectionProperties>().LastOrDefault();
        if (sectProps == null) { sectProps = new SectionProperties(); main.Document.Body.Append(sectProps); }
        sectProps.RemoveAllChildren<FooterReference>();
        sectProps.Append(new FooterReference() { Id = main.GetIdOfPart(footerPart), Type = HeaderFooterValues.Default });
    }

    static void AddChangeLogSection(Body body, string changelogText, DocSettings settings)
    {
        if (!string.IsNullOrWhiteSpace(changelogText))
        {
            try
            {
                var j = JToken.Parse(changelogText);
                if (j is JArray arr)
                {
                    var rows = new List<List<string>> { new() { "ردیف", "نسخه", "تاریخ", "توسعه‌دهنده", "توضیحات تغییر" } };
                    int r = 1;
                    foreach (var it in arr)
                    {
                        rows.Add(new List<string> {
                            r++.ToString(),
                            it["version"]?.ToString() ?? "-",
                            it["date"]?.ToString() ?? "-",
                            it["developer"]?.ToString() ?? "-",
                            it["notes"]?.ToString() ?? "-"
                        });
                    }
                    AddTable(body, rows, settings, true);
                    return;
                }
            }
            catch { /* ignore parse error and fallback to plain text */ }
            AddParagraph(body, changelogText, settings);
        }
        else
        {
            AddParagraph(body, "-", settings);
        }
    }

    static void AddParametersTables(Body body, OpenApiOperation operation, OpenApiComponents components, DocSettings settings)
    {
        
        var headers = operation.Parameters?.Where(x => x.In == ParameterLocation.Header).ToList() ?? new List<OpenApiParameter>();

        if (settings.AuthRequired)
        {
            headers.Add(new OpenApiParameter
            {
                Name = settings.AuthHeader,
                In = ParameterLocation.Header,
                Required = true,
                Schema = new OpenApiSchema { Type = "string", Default = new Microsoft.OpenApi.Any.OpenApiString(settings.AuthValue) },
                Description = "-"
            });
        }

        if (headers.Any())
        {
            AddHeadingWithNumbering(body, "Header Parameters:", 3, settings);
            var rows = new List<List<string>> { new() { "مثال", "شرح", "الزامی", "نوع", "Header" } };
            foreach (var h in headers)
                rows.Add(new List<string> {
                    GetDefaultValue(h.Schema),
                    h.Description ?? "-",
                    h.Required ? "بله" : "خیر",
                    h.Schema?.Type ?? "-",
                    h.Name ?? "-"
                });
            AddTable(body, rows, settings, false);
        }
        else
        {
            AddHeadingWithNumbering(body, "Header Parameters:", 3, settings);
            AddParagraph(body, "-", settings);
        }

        var queries = operation.Parameters?.Where(x => x.In == ParameterLocation.Query).ToList() ?? new List<OpenApiParameter>();
        if (queries.Any())
        {
            AddHeadingWithNumbering(body, "Query Parameters:", 3, settings);
            var rows = new List<List<string>> { new() { "مثال", "شرح", "الزامی", "نوع", "پارامتر" } };
            foreach (var q in queries)
                rows.Add(new List<string> {
                    GetDefaultValue(q.Schema),
                    q.Description ?? "-",
                    q.Required ? "بله" : "خیر",
                    q.Schema?.Type ?? "-",
                    q.Name ?? "-"
                });
            AddTable(body, rows, settings, false);
        }
        else
        {
            AddHeadingWithNumbering(body, "Query Parameters:", 3, settings);
            AddParagraph(body, "-", settings);
        }

        if (operation.RequestBody != null)
        {
            foreach (var c in operation.RequestBody.Content)
            {
                AddHeadingWithNumbering(body, "Body Parameters:", 3, settings);
                AddParagraph(body, $" Body Type: {c.Key}", settings);

                var schema = ResolveSchemaReference(c.Value.Schema, components);
                if (schema?.Properties != null)
                {
                    var bodyRows = new List<List<string>> { new() { "فیلد", "نوع", "شرح", "مثال", "الزامی", "Validation" } };
                    foreach (var prop in schema.Properties)
                        bodyRows.Add(new List<string>{
                            prop.Key,
                            prop.Value.Type,
                            prop.Value.Description ?? "-",
                            prop.Value.Example?.ToString() ?? "-",
                            schema.Required?.Contains(prop.Key) == true ? "*" : "",
                            "-"
                        });
                    AddTable(body, bodyRows, settings, true);
                }
                else AddParagraph(body, "-", settings);
            }
        }
        else
        {
            AddHeadingWithNumbering(body, "Body Parameters:", 3, settings);
            AddParagraph(body, "-", settings);
        }
    }

    static void AddResponsesTables(Body body, OpenApiOperation operation, OpenApiComponents components, DocSettings settings)
    {
        if (operation.Responses.Any())
        {
            AddHeadingWithNumbering(body, "HTTP Code:", 3, settings);
            var respRows = new List<List<string>> { new() { "مثال Response", "شرح", "Error Code", "HTTP Code" } };
            foreach (var r in operation.Responses)
                respRows.Add(new List<string>{
                    GetExampleFromResponse(r.Value) ?? "-",
                    r.Value.Description ?? "-",
                    r.Key,
                    r.Key
                });
            AddTable(body, respRows, settings, true);

            AddHeadingWithNumbering(body, "Response", 3, settings);
            var modelRows = new List<List<string>> { new() { "مثال", "شرح", "نوع داده", "فیلد" } };
            foreach (var r in operation.Responses)
                foreach (var c in r.Value.Content)
                {
                    var schema = ResolveSchemaReference(c.Value.Schema, components);
                    if (schema?.Properties != null)
                        foreach (var prop in schema.Properties)
                            modelRows.Add(new List<string>{
                                schema.Example?.ToString() ?? "-",
                                prop.Value.Description ?? "-",
                                prop.Value.Type ?? "-",
                                prop.Key
                            });
                }
            AddTable(body, modelRows, settings, false);
        }
        else
        {
            AddHeadingWithNumbering(body, "Response:", 3, settings);
            AddParagraph(body, "-", settings);
        }
    }

    static string GetDefaultValue(OpenApiSchema schema, string fallback = "sample")
    {
        if (schema?.Default != null) return ConvertIOpenApiAnyToString(schema.Default);
        return fallback;
    }

    static string GetExampleFromResponse(OpenApiResponse response)
    {
        if (response == null) return "-";
        if (response.Content != null)
        {
            foreach (var c in response.Content)
            {
                if (c.Value.Example != null) return ConvertIOpenApiAnyToString(c.Value.Example);
                if (c.Value.Examples != null && c.Value.Examples.Count > 0)
                {
                    var firstExample = c.Value.Examples.First().Value.Value;
                    if (firstExample != null) return ConvertIOpenApiAnyToString(firstExample);
                }
            }
        }
        return "-";
    }

    static string ConvertIOpenApiAnyToString(Microsoft.OpenApi.Any.IOpenApiAny any)
    {
        if (any == null) return "-";
        return any switch
        {
            Microsoft.OpenApi.Any.OpenApiString s => s.Value,
            Microsoft.OpenApi.Any.OpenApiInteger i => i.Value.ToString(),
            Microsoft.OpenApi.Any.OpenApiLong l => l.Value.ToString(),
            Microsoft.OpenApi.Any.OpenApiBoolean b => b.Value.ToString(),
            Microsoft.OpenApi.Any.OpenApiDouble d => d.Value.ToString(),
            _ => any.ToString()
        };
    }

    static OpenApiSchema ResolveSchemaReference(OpenApiSchema schema, OpenApiComponents components)
    {
        if (schema == null) return null;
        if (schema.Reference != null && components != null && components.Schemas.TryGetValue(schema.Reference.Id, out var resolved))
            schema = resolved;

        if (schema.AllOf != null && schema.AllOf.Count > 0)
        {
            var merged = new OpenApiSchema { Properties = new Dictionary<string, OpenApiSchema>(), Type = "object", Required = new SortedSet<string>() };
            foreach (var s in schema.AllOf)
            {
                var r = ResolveSchemaReference(s, components);
                if (r?.Properties != null)
                {
                    foreach (var kv in r.Properties)
                        if (!merged.Properties.ContainsKey(kv.Key)) merged.Properties[kv.Key] = kv.Value;
                }
                if (r?.Required != null)
                {
                    foreach (var req in r.Required) merged.Required.Add(req);
                }
            }
            if (!string.IsNullOrEmpty(schema.Description)) merged.Description = schema.Description;
            return merged;
        }
        return schema;
    }

    static string GenerateSampleJson(OpenApiSchema schema)
    {
        var dict = new Dictionary<string, object>();
        if (schema?.Properties == null) return "{}";
        foreach (var prop in schema.Properties)
        {
            string type = prop.Value.Type;
            dict[prop.Key] = type switch
            {
                "string" => "sample",
                "integer" => 1,
                "number" => 1.0,
                "boolean" => true,
                _ => "value"
            };
        }
        return JsonConvert.SerializeObject(dict, Formatting.None);
    }

    #endregion
}
