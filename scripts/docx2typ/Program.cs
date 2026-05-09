using System.Text;
using System.Text.RegularExpressions;
using System.Xml;
using System.Text.Json.Serialization;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

string docxPath = args.Length > 0 ? args[0] : null;
if (docxPath == null || !File.Exists(docxPath))
{
    Console.Error.WriteLine("Usage: Docx2Typ <input.docx>");
    return 1;
}

string fileName = Path.GetFileNameWithoutExtension(docxPath);
string safeName = Regex.Replace(fileName, @"[^\w\- ]", "_");
string outputDir = Path.Combine(Path.GetDirectoryName(docxPath)!, $"{safeName}-typ");
string figuresDir = Path.Combine(outputDir, "figures");
Directory.CreateDirectory(figuresDir);

Console.WriteLine($"输入: {docxPath}");
Console.WriteLine($"输出: {outputDir}");

using var doc = WordprocessingDocument.Open(docxPath, false);
var body = doc.MainDocumentPart!.Document.Body!;

// ── 检测文档行距：统计所有段落中最常见的 line 值（240ths of a line）──
var lineValues = new List<int>();
foreach (var para in body.Elements<Paragraph>())
{
    var spacing = para.ParagraphProperties?.SpacingBetweenLines;
    if (spacing?.Line?.Value != null && spacing.LineRule?.Value == LineSpacingRuleValues.Auto)
        lineValues.Add(spacing.Line.Value);
}
// 默认 1.5 倍（360/240），兼容无行距信息或行距信息不全的文档
double lineMult = 1.5;
if (lineValues.Count > 0)
{
    var mode = lineValues.GroupBy(v => v).OrderByDescending(g => g.Count()).First().Key;
    lineMult = Math.Round(mode / 240.0, 1);
    if (lineMult < 0.5) lineMult = 1.0;
    if (lineMult > 3.0) lineMult = 1.5;
}
Console.WriteLine($"检测到行距: {lineMult}x ({lineValues.Count} 个段落)");

// ══════════════════════════════════════════════════
// 第1步：提取图片 + 建立 rId → 文件名 映射
// ══════════════════════════════════════════════════
Console.WriteLine("\n提取图片...");
var imageParts = doc.MainDocumentPart!.ImageParts;
var imageFiles = new Dictionary<string, string>(); // URI → filename
int imgIdx = 0;
foreach (var imgPart in imageParts)
{
    imgIdx++;
    string ext = imgPart.ContentType switch
    {
        "image/jpeg" => ".jpg",
        "image/png" => ".png",
        "image/gif" => ".gif",
        "image/tiff" => ".tiff",
        "image/bmp" => ".bmp",
        _ => ".bin"
    };
    string filename = $"figure{imgIdx}{ext}";
    string savePath = Path.Combine(figuresDir, filename);
    using var srcStream = imgPart.GetStream();
    using var dstStream = File.Create(savePath);
    srcStream.CopyTo(dstStream);

    // TIFF → PNG 自动转换（Typst 不支持 TIFF）
    if (ext == ".tiff" || ext == ".tif")
    {
        string pngName = $"figure{imgIdx}.png";
        string pngPath = Path.Combine(figuresDir, pngName);
        dstStream.Close();
        try
        {
            byte[] imageBytes = File.ReadAllBytes(savePath);
            using var ms = new MemoryStream(imageBytes);
            using var img = System.Drawing.Image.FromStream(ms);
            img.Save(pngPath, System.Drawing.Imaging.ImageFormat.Png);
            File.Delete(savePath);
            filename = pngName;
            Console.WriteLine($"  -> 已转换为 {pngName}");
        }
        catch (Exception ex)
        {
            Console.WriteLine($"  ⚠ TIFF 转换失败: {ex.Message}，保留原始 {filename}");
        }
    }

    imageFiles[imgPart.Uri.ToString()] = filename;
    Console.WriteLine($"  {imgPart.Uri} -> {filename}");
}

// 建立 rId (关系ID) → 图片文件名 的映射
// 通过读取 word/_rels/document.xml.rels 获取
var ridToFile = new Dictionary<string, string>();
var relsXml = doc.MainDocumentPart!.Parts;
foreach (var part in relsXml)
{
    if (part.OpenXmlPart is ImagePart imgPart)
    {
        string rid = part.RelationshipId;
        if (imageFiles.TryGetValue(imgPart.Uri.ToString(), out string? fname))
            ridToFile[rid] = fname;
    }
}

// ══════════════════════════════════════════════════
// 第2步：遍历段落，提取内容 + 检测图片引用
// ══════════════════════════════════════════════════
// 对每个段落拿到原始 XML 文本，用正则找 <a:blip r:embed="rIdX"/>
// 这比 OpenXML 的 Descendants<T>() 可靠，能 100% 抓到所有图片引用
var blocks = new List<(string Type, string Content)>();
bool inToc = false;
int paraIndex = 0;

using (var xmlReader = XmlReader.Create(doc.MainDocumentPart!.GetStream()))
{
    // 直接用 OpenXML 的 body.Elements() 遍历段落
} // 保留原遍历方式，但补充 XML 文本扫描

// 辅助：获取段落的原始 XML 字符串
string GetParagraphXml(Paragraph p)
{
    using var sw = new StringWriter();
    using var xw = XmlWriter.Create(sw, new XmlWriterSettings { OmitXmlDeclaration = true });
    p.WriteTo(xw);
    xw.Flush();
    return sw.ToString();
}

foreach (var element in body.Elements())
{
    if (element is Paragraph para)
    {
        paraIndex++;
        string styleId = para.ParagraphProperties?.ParagraphStyleId?.Val?.Value ?? "";
        string text = string.Concat(para.Elements<Run>().SelectMany(r =>
            r.Elements<Text>().Select(t => t.Text))).Trim();

        // ── 图片检测：空段落也要检查是否有图片 ──
        var paraXml = GetParagraphXml(para);
        var blipMatches = Regex.Matches(paraXml, @"a:blip[^>]*r:embed=""([^""]+)""");
        var foundImages = new List<string>();
        foreach (System.Text.RegularExpressions.Match m in blipMatches)
        {
            string rid = m.Groups[1].Value;
            if (ridToFile.TryGetValue(rid, out string? imgFile))
                foundImages.Add(imgFile);
        }

        if (foundImages.Count > 0)
        {
            // 有图题注则标记，无则直接输出图片
            bool isFigCaption = !inToc && Regex.Match(text ?? "", @"^图\s*\d").Success;
            if (isFigCaption) blocks.Add(("fig_caption", text));
            foreach (var img in foundImages)
                blocks.Add(("image", img));
            continue;
        }

        // ── 纯文字图题注（图在邻段，本段只有文字）──
        bool isFigCaptionText = !inToc && text.Length > 0 && text.Length < 60 && Regex.Match(text, @"^图\s*\d").Success;
        if (isFigCaptionText && foundImages.Count == 0)
        {
            blocks.Add(("fig_caption", text));
            continue;
        }

        if (string.IsNullOrWhiteSpace(text))
        {
            if (!inToc && blocks.Count > 0 && blocks.Last().Type != "empty")
                blocks.Add(("empty", ""));
            continue;
        }

        // ── 特殊标记 → 跳过 ──
        if (styleId == "af0" && text is "目录" or "表目录" or "图目录")
        {
            inToc = true;
            blocks.Add(("skip", text));
            continue;
        }

        // ── 标题检测（样式优先，文本保底） ──
        string? headingLevel = null;
        string headingText = text;

        if (styleId == "1" || (styleId != "af0" && Regex.Match(text, @"^[一二三四五六七八九十]、").Success))
        {
            headingLevel = "section";
            headingText = Regex.Replace(text, @"^[一二三四五六七八九十]、\s*", "");
        }
        else if (styleId == "2" || Regex.Match(text, @"^（[一二三四五六七八九十]）").Success)
        {
            headingLevel = "subsection";
            headingText = Regex.Replace(text, @"^（[一二三四五六七八九十]）\s*", "");
        }
        else if (styleId == "4" || Regex.Match(text, @"^\d+[．\.、]").Success)
        {
            headingLevel = "subsubsection";
            headingText = Regex.Replace(text, @"^\d+[．\.、]\s*", "");
        }

        if (styleId == "af0" && text == "前言") { headingLevel = "section"; headingText = "前言"; inToc = false; }

        if (headingLevel != null) { inToc = false; blocks.Add((headingLevel, headingText)); continue; }

        // ── 封面大标题 → 跳过 ──
        var firstRun = para.Elements<Run>().FirstOrDefault();
        if (firstRun?.RunProperties?.FontSize?.Val != null)
        {
            string halfPt = firstRun.RunProperties.FontSize.Val.Value;
            if (int.TryParse(halfPt, out int hp) && hp >= 48) { blocks.Add(("skip", text)); continue; }
        }

        // ── 目录内容 → 跳过 ──
        if (inToc) { blocks.Add(("skip", text)); continue; }

        // ── 表格标题 ──（去除 表1.1、表1、附表2 等前缀，保留描述文字）
        if (text.StartsWith("表") || text.StartsWith("附表"))
        {
            var tableCap = Regex.Replace(text, @"^(表|附表)\s*\d+(\.\d+)?\s*", "");
            if (!string.IsNullOrWhiteSpace(tableCap))
                blocks.Add(("table_caption", tableCap));
            continue;
        }

        // ── 正文 ──
        blocks.Add(("para", text));
    }
    else if (element is Table table)
    {
        var rows = table.Elements<TableRow>().ToList();
        var tblPr = table.Elements<TableProperties>().FirstOrDefault();

        // — 读取 tblGrid 列宽（DXA），用于计算 Q 列宽度比例 —
        var grid = table.Elements<TableGrid>().FirstOrDefault();
        var gridWidths = grid?.Elements<GridColumn>()
            .Select(g => int.TryParse(g.Width?.Value ?? "", out int w) ? Math.Max(w, 1) : 2000)
            .ToList() ?? new List<int>();

        // — 辅助：读取一个 Run 的带格式文本 —
        string GetRunText(Run r, out double fontSizePt)
        {
            var txt = string.Concat(r.Elements<Text>().Select(t => t.Text));
            fontSizePt = 0;
            if (string.IsNullOrEmpty(txt)) return "";
            var rPr = r.RunProperties;
            bool bold = (rPr?.Bold?.Val?.Value ?? false) || (rPr?.BoldComplexScript?.Val?.Value ?? false);
            bool italic = (rPr?.Italic?.Val?.Value ?? false) || (rPr?.ItalicComplexScript?.Val?.Value ?? false);
            var va = rPr?.VerticalTextAlignment?.Val?.Value;
            string sz = rPr?.FontSize?.Val?.Value ?? "";
            string szCs = rPr?.FontSizeComplexScript?.Val?.Value ?? "";
            string fs = !string.IsNullOrEmpty(sz) ? sz : !string.IsNullOrEmpty(szCs) ? szCs : "";
            if (!string.IsNullOrEmpty(fs) && int.TryParse(fs, out int hv))
                fontSizePt = hv / 2.0;
            txt = EscapeTypst(txt);
            if (va == VerticalPositionValues.Superscript) txt = $"#super[{txt}]";
            else if (va == VerticalPositionValues.Subscript) txt = $"#sub[{txt}]";
            if (bold && italic) txt = $"#text(weight: \"bold\", style: \"italic\")[{txt}]";
            else if (bold) txt = $"#text(weight: \"bold\")[{txt}]";
            else if (italic) txt = $"#text(style: \"italic\")[{txt}]";
            return txt;
        }

        // — 辅助：读取一个 Cell 的文本 + 字号 —
        (string text, double fontSize) GetCellInfo(TableCell cell)
        {
            double maxFz = 0;
            string result = "";
            foreach (var r in cell.Descendants<Run>())
            {
                string part = GetRunText(r, out double fz);
                result += part;
                if (fz > maxFz) maxFz = fz;
            }
            return (result.Trim(), maxFz);
        }

        // 第一步：读取每个单元格的文字、跨列/跨行、对齐
        var rawRows = new List<List<CellRec>>();
        foreach (var row in rows)
        {
            var cells = row.Elements<TableCell>().ToList();
            var rowRecs = new List<CellRec>();
            foreach (var cell in cells)
            {
                var (text, fz) = GetCellInfo(cell);
                int cs = 1, rs = 1;
                string align = "", bg = "";
                var tcPr = cell.TableCellProperties;
                if (tcPr?.GridSpan?.Val?.Value != null)
                    cs = tcPr.GridSpan.Val.Value;
                var vm = tcPr?.VerticalMerge?.Val?.Value;
                if (vm == MergedCellValues.Restart) rs = 2;
                else if (vm == MergedCellValues.Continue) { text = null!; rs = 0; }

                // 水平对齐（从首段 pPr/jc 读取）
                var firstP = cell.Descendants<ParagraphProperties>().FirstOrDefault();
                var jc = firstP?.Justification?.Val?.Value;
                if (jc == JustificationValues.Center) align = "c";
                else if (jc == JustificationValues.Right) align = "r";
                else if (jc == JustificationValues.Both) align = "j";

                // 单元格底色
                var shd = tcPr?.Shading;
                if (shd?.Fill?.Value != null && shd.Fill.Value != "auto")
                    bg = shd.Fill.Value;

                rowRecs.Add(new CellRec { T = text, C = cs, R = rs, A = align, Bg = bg, Fz = fz });
            }
            rawRows.Add(rowRecs);
        }

        // 第二步：计算跨行数
        for (int ri = 0; ri < rawRows.Count; ri++)
            for (int ci = 0; ci < rawRows[ri].Count; ci++)
                if (rawRows[ri][ci].R == 2)
                {
                    int rs = 1;
                    for (int rj = ri + 1; rj < rawRows.Count && ci < rawRows[rj].Count && rawRows[rj][ci].R == 0; rj++, rs++) ;
                    rawRows[ri][ci].R = rs;
                }

        // 第三步：计算物理列数
        int maxCol = rawRows.Max(r => {
            int sum = 0;
            foreach (var c in r) { if (c.R > 0) sum += c.C; }
            return sum;
        });

        // — 第四步：列对齐 + 宽度感知 —
        // 统计每列主要对齐方向
        var colAlign = new string[maxCol];
        for (int c = 0; c < maxCol; c++) colAlign[c] = "c"; // 全部居中，不超出页面

        // — 内容感知权重 —
        var colTextMaxLen = new int[maxCol];
        for (int ri = 0; ri < rawRows.Count; ri++)
        {
            int pos = 0;
            for (int ci = 0; ci < rawRows[ri].Count && pos < maxCol; ci++)
            {
                var cr = rawRows[ri][ci];
                if (cr.R == 0) { pos++; continue; }
                double adjLen = cr.T?.Sum(ch => ch > 127 ? 2.0 : (char.IsDigit(ch) ? 1.8 : 1.2)) ?? 0;
                int avgPerCol = cr.C > 0 ? (int)Math.Ceiling(adjLen / cr.C) : (int)adjLen;
                for (int j = 0; j < cr.C && pos + j < maxCol; j++)
                    colTextMaxLen[pos + j] = Math.Max(colTextMaxLen[pos + j], avgPerCol);
                pos += cr.C;
            }
        }

        // — 列宽（X 比例权重，保证不超出页面）—
        bool narrowTable = maxCol >= 8;
        int totalW = gridWidths.Count >= maxCol ? gridWidths.Take(maxCol).Sum() : maxCol * 2000;
        if (totalW <= 0) totalW = maxCol * 2000;
        int baseTotal = narrowTable ? maxCol * 3 : maxCol * 2;

        int[] weights = new int[maxCol];
        for (int c = 0; c < maxCol; c++)
        {
            // 基础权重来自 DOCX 网格列宽（归一化）
            int baseW = c < gridWidths.Count ? gridWidths[c] : (totalW / maxCol);
            weights[c] = Math.Max(1, (int)Math.Round((double)baseW / totalW * baseTotal));
            // 内容宽 → 权重加成
            if (colTextMaxLen[c] > 16) weights[c] += 4;
            else if (colTextMaxLen[c] > 10) weights[c] += 3;
            else if (colTextMaxLen[c] > 5) weights[c] += 1;
            // 前 2 列（标签列）→ 额外加成
            if (c < 2 && maxCol > 6)
                weights[c] = Math.Max(weights[c], baseTotal / maxCol * 2);
        }
        // 保证总权重合理且不爆出页面
        int finalTotal = weights.Sum();
        int maxAllowed = narrowTable ? baseTotal + maxCol * 2 : baseTotal + maxCol;
        if (finalTotal > maxAllowed)
            for (int c = 0; c < maxCol; c++)
                weights[c] = Math.Max(1, weights[c] * maxAllowed / finalTotal);

        string colSpec = string.Join(" ", Enumerable.Range(0, maxCol).Select(c =>
        {
            string a = colAlign[c];
            return $"X[{a}, {weights[c]}]";
        }));

        // — 第五步：rowhead 检测 —
        // 中文政府报告中，首行含 span（如 土地利用类型 c=2）则表头通常为 2 行
        int headerRows = 1;
        if (rawRows.Count >= 2)
        {
            bool row0HasSpan = rawRows[0].Any(c => (c.C > 1 || c.R > 1) && c.R > 0);
            bool row1LooksHeader = rawRows[1].Any(c => c.R > 0 && !string.IsNullOrWhiteSpace(c.T));
            if (row0HasSpan && row1LooksHeader) headerRows = 2;
        }

        // — 表格级默认字号（统计所有有内容的格的字体，取主流）—
        var fzSizes = rawRows.SelectMany(r => r)
            .Where(c => c.R > 0 && c.Fz > 0)
            .GroupBy(c => c.Fz)
            .OrderByDescending(g => g.Count())
            .ToList();
        double tableFz = fzSizes.FirstOrDefault()?.Key ?? 0;
        // 如果主流字号跟正文（12pt）差不多，就不要额外设
        if (tableFz > 0 && Math.Abs(tableFz - 12) < 0.5) tableFz = 0;

        // — 第六步：构建每行物理列 —
        var tableData = new List<List<object>>();
        for (int ri = 0; ri < rawRows.Count; ri++)
        {
            var outRow = new List<object>();
            int pos = 0;
            for (int ci = 0; ci < rawRows[ri].Count && pos < maxCol; ci++)
            {
                var cr = rawRows[ri][ci];
                if (cr.R == 0)
                {
                    outRow.Add("");
                    pos += 1;
                }
                else if (cr.C > 1 || cr.R > 1)
                {
                    var cellObj = new Dictionary<string, object> { ["t"] = cr.T, ["c"] = cr.C, ["r"] = cr.R };
                    if (!string.IsNullOrEmpty(cr.Bg)) cellObj["bg"] = cr.Bg;
                    outRow.Add(cellObj);
                    pos += cr.C;
                }
                else
                {
                    if (!string.IsNullOrEmpty(cr.Bg))
                        outRow.Add(new Dictionary<string, object> { ["t"] = cr.T, ["bg"] = cr.Bg });
                    else
                        outRow.Add(cr.T);
                    pos += 1;
                }
            }
            while (pos < maxCol) { outRow.Add(""); pos++; }
            tableData.Add(outRow);
        }

        blocks.Add(("table_data2", System.Text.Json.JsonSerializer.Serialize(new
        {
            cols = maxCol,
            colspec = colSpec,
            rowhead = headerRows,
            tableFz = tableFz,
            rows = tableData
        })));
    }
}

// ══════════════════════════════════════════════════
// 第3步：生成 Typst 内容
// ══════════════════════════════════════════════════
Console.WriteLine("\n生成 Typst 文件...");

var bodyTyp = new StringBuilder();
int tableCounter = 0;
int figureCounter = 0;
string _pendingTableCaption = "";
string _pendingFigCaption = "";

foreach (var (type, content) in blocks)
{
    switch (type)
    {
        case "skip": break;

        case "section":
            bodyTyp.AppendLine($"= {EscapeTypst(content)}");
            bodyTyp.AppendLine();
            break;

        case "subsection":
            bodyTyp.AppendLine($"== {EscapeTypst(content)}");
            bodyTyp.AppendLine();
            break;

        case "subsubsection":
            bodyTyp.AppendLine($"=== {EscapeTypst(content)}");
            bodyTyp.AppendLine();
            break;

        case "para":
            bodyTyp.AppendLine(EscapeTypst(content));
            bodyTyp.AppendLine();
            break;

        case "empty":
            bodyTyp.AppendLine();
            break;

        case "fig_caption":
            var figCaption = content;
            var figPrefixMatch = Regex.Match(figCaption, @"^图\s*\d+(\.\d+)?[a-zA-Z]?\s+");
            if (figPrefixMatch.Success)
                figCaption = figCaption[figPrefixMatch.Length..].Trim();
            if (!string.IsNullOrEmpty(figCaption) && string.IsNullOrEmpty(_pendingFigCaption))
                _pendingFigCaption = EscapeTypst(figCaption);
            break;

        case "image":
            figureCounter++;
            bodyTyp.AppendLine("#figure(");
            bodyTyp.AppendLine($"  image(\"figures/{content}\", width: 80%),");
            if (!string.IsNullOrEmpty(_pendingFigCaption))
            {
                bodyTyp.AppendLine($"  caption: [{_pendingFigCaption}],");
                _pendingFigCaption = "";
            }
            else
            {
                bodyTyp.AppendLine("  caption: [],");
            }
            bodyTyp.AppendLine($"  supplement: [图],");
            bodyTyp.AppendLine($"  kind: \"image\",");
            bodyTyp.AppendLine(")");
            bodyTyp.AppendLine();
            break;

        case "table_caption":
            _pendingTableCaption = EscapeTypst(content);
            break;

        case "table_data2":
        {
            using var jsonDoc = System.Text.Json.JsonDocument.Parse(content);
            int colCount = jsonDoc.RootElement.GetProperty("cols").GetInt32();
            var rowsJson = jsonDoc.RootElement.GetProperty("rows").EnumerateArray();

            string colSpec = jsonDoc.RootElement.TryGetProperty("colspec", out var csProp)
                ? csProp.GetString() ?? string.Concat(Enumerable.Repeat("X", colCount))
                : string.Concat(Enumerable.Repeat("X", colCount));
            int headerRows = jsonDoc.RootElement.TryGetProperty("rowhead", out var rhProp)
                ? Math.Max(rhProp.GetInt32(), 1) : 1;
            double tFz = jsonDoc.RootElement.TryGetProperty("tableFz", out var fzProp) ? fzProp.GetDouble() : 0;

            var allRows = new List<List<(string text, int cs, int rs, string bg)>>();
            foreach (var r in rowsJson)
            {
                var rowCells = new List<(string text, int cs, int rs, string bg)>();
                foreach (var c in r.EnumerateArray())
                {
                    string t = ""; int cs = 1, rs = 1; string bg = "";
                    if (c.ValueKind == System.Text.Json.JsonValueKind.String)
                        { t = c.GetString() ?? ""; }
                    else
                    {
                        t = c.TryGetProperty("t", out var tv) ? tv.GetString() ?? "" : "";
                        if (c.TryGetProperty("c", out var cv)) cs = cv.GetInt32();
                        if (c.TryGetProperty("r", out var rv)) rs = rv.GetInt32();
                        if (c.TryGetProperty("bg", out var bv)) bg = bv.GetString() ?? "";
                    }
                    rowCells.Add((t, cs, rs, bg));
                }
                allRows.Add(rowCells);
            }

            // — 列宽：从 colspec 提取权重 —
            var weights = new List<int>();
            foreach (System.Text.RegularExpressions.Match wm in Regex.Matches(colSpec, @"X\[[lcrj]?[,\s]*(\d+)\]"))
                weights.Add(int.Parse(wm.Groups[1].Value));
            while (weights.Count < colCount) weights.Add(1);
            bool allEqual = weights.Distinct().Count() == 1;

            // — 收集所有单元格 —
            var headerCells = new List<string>();
            var bodyCells = new List<string>();
            int cellCount = 0;

            for (int ri = 0; ri < allRows.Count; ri++)
            {
                int pos = 0;
                for (int ci = 0; ci < allRows[ri].Count && pos < colCount; ci++)
                {
                    var (text, cs, rs, bg) = allRows[ri][ci];
                    int span = Math.Max(cs, 1);

                    if (rs == 0) { pos += span; continue; }

                    if (string.IsNullOrEmpty(text))
                    {
                        cellCount++;
                        pos += span;
                        continue;
                    }

                    if (cs > 1 || rs > 1)
                    {
                        var cellArgs = new List<string>();
                        if (cs > 1) cellArgs.Add($"colspan: {cs}");
                        if (rs > 1) cellArgs.Add($"rowspan: {rs}");
                        if (ri < headerRows)
                            headerCells.Add($"table.cell({string.Join(", ", cellArgs)})[{FormatCellTypst(text)}]");
                        else
                            bodyCells.Add($"table.cell({string.Join(", ", cellArgs)})[{FormatCellTypst(text)}]");
                    }
                    else
                    {
                        if (ri < headerRows)
                            headerCells.Add($"[{FormatCellTypst(text)}]");
                        else
                            bodyCells.Add($"[{FormatCellTypst(text)}]");
                    }
                    cellCount++;
                    pos += span;
                }
            }

            tableCounter++;
            bodyTyp.AppendLine("#figure(");

            // 字号作用域（用 text() 包裹代替代码块，保证跨页断开）
            bool wrapWithText = (tFz > 0 && Math.Abs(tFz - 12) > 0.5) || colCount >= 8;
            double textSz = tFz > 0 && Math.Abs(tFz - 12) > 0.5 ? tFz :
                            colCount >= 12 ? 9 :
                            colCount >= 8 ? 10 : 12;

            if (wrapWithText)
                bodyTyp.AppendLine($"  text(size: {textSz}pt,");

            bodyTyp.AppendLine("    table(");

            // 列定义
            if (allEqual)
                bodyTyp.AppendLine($"      columns: {colCount},");
            else
            {
                var colDefs = string.Join(", ", weights.Take(colCount).Select(w => w + "fr"));
                bodyTyp.AppendLine($"      columns: ({colDefs}),");
            }

            bodyTyp.AppendLine("      stroke: 0.5pt,");

            // 表头行
            if (headerCells.Count > 0)
            {
                bodyTyp.AppendLine("      table.header(");
                for (int i = 0; i < headerCells.Count; i++)
                    bodyTyp.AppendLine($"        {headerCells[i]},");
                bodyTyp.AppendLine("      ),");
            }

            // 表体行
            for (int i = 0; i < bodyCells.Count; i++)
                bodyTyp.AppendLine($"      {bodyCells[i]},");

            if (wrapWithText)
            {
                bodyTyp.AppendLine("    )");
                bodyTyp.AppendLine("  ),");
            }
            else
            {
                bodyTyp.AppendLine("    ),");
            }

            // 题注
            if (!string.IsNullOrEmpty(_pendingTableCaption))
            {
                bodyTyp.AppendLine($"  caption: [{_pendingTableCaption}],");
                _pendingTableCaption = "";
            }
            bodyTyp.AppendLine("  supplement: [表],");
            bodyTyp.AppendLine("  kind: \"table\",");
            bodyTyp.AppendLine(")");
            bodyTyp.AppendLine();
            break;
        }
    }
}

// ══════════════════════════════════════════════════
// 第4步：写出文件
// ══════════════════════════════════════════════════

// ── 辅助函数 ──

string EscapeTypst(string text)
{
    // 去掉零宽字符
    text = Regex.Replace(text, @"[" + '​' + '‌' + '‍' + '‎' + '‏' + '﻿' + @"]", "");
    // Typst content 模式下转义 \ # [ ] < >
    return text
        .Replace("\\", "\\\\")
        .Replace("#", "\\#")
        .Replace("[", "\\[")
        .Replace("]", "\\]")
        .Replace("<", "\\<")
        .Replace(">", "\\>");
}

string FormatCellTypst(string text)
{
    // text is already EscapeTypst'd; Typst handles digit wrapping natively
    return text;
}

// ── 写 main.typ（格式规则 + 内容一体化） ──
string typPreambleBefore = @"// ── 页面设置 ──
#set page(
  margin: (left: 2.5cm, right: 2.5cm, top: 3cm, bottom: 2.5cm),
  numbering: ""1"",
)

// ── 正文样式（小四号宋体）──
#set text(
  font: (""SimSun"", ""Times New Roman""),
  size: 12pt,
  lang: ""zh"",
)

// ── 标题编号（每级独立格式）──
#set heading(numbering: (..numbers) => {
  let level = numbers.pos().len()
  if level == 1 { numbering(""一、"", numbers.pos().at(0)) }
  else if level == 2 { numbering(""（一）"", numbers.pos().at(1)) }
  else if level == 3 { numbering(""1."", numbers.pos().at(2)) }
  else if level == 4 { numbering(""(1)"", numbers.pos().at(3)) }
  else { numbering(""a."", numbers.pos().at(4)) }
})

// ── 标题样式 ──
#show heading.where(level: 1): it => {
  set text(font: (""SimHei"",), size: 16pt, weight: ""bold"")
  it
}
#show heading.where(level: 2): it => {
  set text(font: (""SimSun"",), size: 14pt, weight: ""bold"")
  it
}
#show heading.where(level: 3): it => {
  set text(font: (""SimHei"",), size: 12pt, weight: ""bold"")
  it
}
#show heading.where(level: 4): it => {
  set text(font: (""SimSun"",), size: 12pt, weight: ""bold"")
  it"
string typLineSpacing = $"	// ── 行距（等效Word {lineMult}倍行距）──\n\t#set par(leading: {lineMult}em, spacing: {lineMult}em)\n\n";
string typPreambleAfter = @"
// ── 图表规则 ──
#show figure.where(kind: ""table""): set block(breakable: true)
#show figure.where(kind: ""table""): set figure.caption(position: top)

";

string typPreamble = typPreambleBefore + typLineSpacing + typPreambleAfter;

string mainTyp = $@"// Auto-converted from DOCX: {EscapeTypst(fileName)}
{typPreamble}// === 正文开始 ===
{bodyTyp.ToString().TrimEnd()}
";
File.WriteAllText(Path.Combine(outputDir, "main.typ"), mainTyp);
Console.WriteLine("  main.typ");

File.WriteAllText(Path.Combine(outputDir, "ref.bib"), "// 参考文献\n");
Console.WriteLine("  ref.bib");

// ── 写 .vscode/settings.json (Tinymist) ──
var vsDir = Path.Combine(outputDir, ".vscode");
Directory.CreateDirectory(vsDir);
File.WriteAllText(Path.Combine(vsDir, "settings.json"), @"{
  ""tinymist.exportPdf"": ""onSave"",
  ""tinymist.formatterMode"": ""typstyle"",
  ""[typst]"": {
    ""editor.formatOnSave"": true
  }
}
");
Console.WriteLine("  .vscode/settings.json");

Console.WriteLine("\n✅ 转换完成！");
Console.WriteLine($"   Typst 文件: cd \"{outputDir}\" && typst compile main.typ");
return 0;

// 表格单元格记录（含跨行列信息）
class CellRec { public string T { get; set; } = ""; public int C { get; set; } = 1; public int R { get; set; } = 1; public string A { get; set; } = ""; public string Bg { get; set; } = ""; public double Fz { get; set; } = 0; }
