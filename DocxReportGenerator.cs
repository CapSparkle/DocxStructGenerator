﻿using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using Tap.Zis3.Domain.Services.Reports.TemplateDtos;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;

class DocxReportGenerator
{
    static Regex blockStartRegex = new Regex(@"\{#(RepeatableBlockStart|BlockStart):([\w\d_]+):([\w\d_]+)\}");
    static Regex blockEndRegex = new Regex(@"\{#(RepeatableBlockEnd|BlockEnd):([\w\d_]+):([\w\d_]+)\}");
    static Regex blockTagStartRegex = new Regex(@"\{#");
    static Regex blockTagEndRegex = new Regex(@"\}");

    public static void Main()
    {
        string templatePath = DocxStructGenerator.filePath;
        string outputPath = Path.Combine(AppContext.BaseDirectory, "Reports", "Templates", "FILLEDInformationAnalyticCardForComplexTemplate.docx"); ;

        var data = new InformationAnalyticCardForComplexTemplateDto
        {
            LandPlotProperties = new List<LandPlotProperties>
            {
                new LandPlotProperties { Area = "500 ha", LandPlotRightsType = "Ownership", plotNotFormed = "No", plotFormed = "Yes" },
                new LandPlotProperties { Area = "300 ha", LandPlotRightsType = "Lease", plotNotFormed = "Yes", plotFormed = "No" }
            },
            CommonComplexProperties = new CommonComplexProperties
            {
                Address = "123 Main St",
                TotalBuildingsArea = "15000 sq.m",
                AreaMeassure = "Square meters"
            }
        };

        FillDocx(templatePath, outputPath, data);
        Console.WriteLine("Document filled successfully!");
    }

    public static void FillDocx<T>(string templatePath, string outputPath, T data)
    {
        File.Copy(templatePath, outputPath, true);

        NormalizeDocument(outputPath);

        using (WordprocessingDocument document = WordprocessingDocument.Open(outputPath, true))
        {
            var mainPart = document.MainDocumentPart;
            if (mainPart?.Document?.Body != null)
            {
                var body = mainPart.Document.Body;

                // Step 1: Identify and Process Repeatable Blocks
                ProcessRepeatableBlocks(body, data);

                // Step 2: Replace all placeholders in remaining text
                ReplacePlaceholders(body, data);

                document.Save();
            }
        }
    }

    private static void ProcessRepeatableBlocks<T>(Body body, T data)
    {
        var paragraphs = body.Elements<Paragraph>().ToList();
        int i = 0;

        while (i < paragraphs.Count)
        {
            Paragraph paragraph = paragraphs[i];

            // Check if this is a tag-paragraph (Start of Block or RepeatableBlock)
            if (FindRunsContainingTag(paragraph, out _, out _, blockStartRegex))
            {
                // Extract repeatable block metadata
                string tagText = paragraph.InnerText.Trim('{', '}');
                string[] parts = tagText.Split(':');

                if (parts.Length == 3 && (parts[0] == "#BlockStart" || parts[0] == "#RepeatableBlockStart"))
                {
                    string className = parts[1];
                    string propertyName = parts[2];

                    // Find the corresponding end tag paragraph
                    var endTagParagraph = paragraphs.FirstOrDefault(p => p.InnerText.Contains("{#BlockEnd:" + className) ||
                                                                          p.InnerText.Contains("{#RepeatableBlockEnd:" + className));
                    if (endTagParagraph == null) continue;

                    // Extract all paragraphs between start & end tags
                    var blockContent = ExtractBlockContent(paragraphs, paragraph, endTagParagraph);
                    if (blockContent.Count == 0) continue;

                    // Find the property in the data object
                    var property = typeof(T).GetProperty(propertyName);
                    if (property != null)
                    {
                        // If it's a simple block (single object)
                        if (!typeof(System.Collections.IEnumerable).IsAssignableFrom(property.PropertyType))
                        {
                            var value = property.GetValue(data);
                            if (value != null)
                            {
                                ReplacePlaceholders(blockContent, value);
                                foreach (var element in blockContent)
                                {
                                    paragraph.InsertBeforeSelf(element);
                                }
                            }
                        }
                        // If it's a repeatable block (list of objects)
                        else
                        {
                            var list = (System.Collections.IEnumerable)property.GetValue(data);
                            if (list != null)
                            {
                                foreach (var item in list)
                                {
                                    var clonedBlock = blockContent.Select(el => el.CloneNode(true)).ToList();
                                    ReplacePlaceholders(clonedBlock, item);
                                    foreach (var element in clonedBlock)
                                    {
                                        paragraph.InsertBeforeSelf(element);
                                    }
                                }
                            }
                        }
                    }

                    // Remove tag-paragraphs (start & end)
                    paragraph.Remove();
                    endTagParagraph.Remove();

                    // Skip to the next paragraph (to avoid reprocessing removed elements)
                    i++;
                    continue;
                }
            }

            // If it's not a tag-paragraph, replace placeholders for the main object
            ReplacePlaceholders(new List<OpenXmlElement> { paragraph }, data);

            i++; // Move to the next paragraph
        }
    }



    private static List<OpenXmlElement> ExtractBlockContent(List<Paragraph> paragraphs, int startIndex, string className, string propertyName)
    {
        List<OpenXmlElement> blockContent = new List<OpenXmlElement>();
        int i = startIndex + 1;

        while (i < paragraphs.Count)
        {
            var textElement = paragraphs[i].Elements<Run>().Select(r => r.GetFirstChild<Text>()).FirstOrDefault(t => t != null && t.Text.Contains("{#RepeatableBlockEnd"));

            if (textElement != null && textElement.Text.Contains($"{className}:{propertyName}"))
            {
                break;
            }

            blockContent.Add(paragraphs[i]);
            i++;
        }

        return blockContent;
    }

    private static void ReplacePlaceholders(IEnumerable<OpenXmlElement> elements, object data)
    {
        Dictionary<string, string> replacements = CollectPropertyValues(data);

        foreach (var element in elements)
        {
            foreach (var text in element.Descendants<Text>())
            {
                if (text != null)
                {
                    foreach (var key in replacements.Keys)
                    {
                        if (text.Text.Contains($"{{{key}}}"))
                        {
                            text.Text = text.Text.Replace($"{{{key}}}", replacements[key]);
                        }
                    }
                }
            }
        }
    }

    private static Dictionary<string, string> CollectPropertyValues(object obj)
    {
        Dictionary<string, string> values = new Dictionary<string, string>();
        if (obj == null) return values;

        Type type = obj.GetType();
        foreach (PropertyInfo property in type.GetProperties())
        {
            values[property.Name] = property.GetValue(obj)?.ToString() ?? "";
        }
        return values;
    }

    static void NormalizeDocument(string filePath)
    {
        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, true))
        {
            var paragraphs = doc.MainDocumentPart.Document.Body.Elements<Paragraph>().ToList();

            foreach (var paragraph in paragraphs)
            {
                NormalizeParagraph(paragraph);
            }
        }
    }

    static void NormalizeParagraph(Paragraph paragraph)
    {
        while (FindRunsContainingTag(paragraph, out int firstRunIndex, out int lastRunIndex,
            blockStartRegex, 
            blockEndRegex,
            Regex))
        {
            CutOffTagParagraph(paragraph, firstRunIndex, lastRunIndex);
        }
    }

    public static bool FindRunsContainingTag(Paragraph paragraph, out int firstRunIndex, out int lastRunIndex, params Regex[] regexPatterns)
    {
        firstRunIndex = -1;
        lastRunIndex = -1;

        List<Run> runs = paragraph.Descendants<Run>().ToList();

        // Step 1: Combine all run texts properly using InnerText
        string combinedText = string.Concat(runs.Select(r => r.InnerText));

        // Step 2: Find the first match of any regex pattern
        int startIndex = -1;
        int matchLength = 0;

        foreach (var regex in regexPatterns)
        {
            Match match = regex.Match(combinedText);
            if (match.Success)
            {
                if (startIndex == -1 || match.Index < startIndex) // Pick the first occurring match
                {
                    startIndex = match.Index;
                    matchLength = match.Length;
                }
            }
        }

        if (startIndex == -1)
        {
            return false; // No match found
        }

        int endIndex = startIndex + matchLength;
        int currentCharIndex = 0;

        // Step 3: Find which runs contain the match
        for(int i = 0; i < runs.Count; i++)
        {
            string runText = runs[i].InnerText;
            int runLength = runText.Length;

            // If the run contains any part of the matched substring, include it
            if (currentCharIndex < endIndex && (currentCharIndex + runLength) > startIndex)
            {
                if (firstRunIndex == -1)
                    firstRunIndex = i;
            }
            else if(firstRunIndex != -1)
            {
                lastRunIndex = i;
                break;
            }

            currentCharIndex += runLength;
        }

        return true;
    }

    /// <summary>
    /// Вырезать новый параграф из старого.
    /// </summary>
    /// <remarks>
    /// Для корректной работы функции предполагается, что в <paramref name="originalParagraph"/>
    /// присутствуют только Run -ы и ParagraphProperties
    /// </remarks>
    /// <param name="originalParagraph"></param>
    /// <param name="firstRunIndex"></param>
    /// <returns>Последний (нижний) по ходу документа параграф</returns>
    public static void CutOffTagParagraph(Paragraph originalParagraph, int firstRunIndex, int lastRunIndex)
    {
        List<Run> originalRuns;
        
        if (originalParagraph == null || firstRunIndex < 0 || firstRunIndex > lastRunIndex)
        {
            return;
        }
        else
        {
            originalRuns = originalParagraph.Elements<Run>().ToList();
            
            if(lastRunIndex >= originalRuns.Count())
            {
                return;
            }
        }

        if (firstRunIndex > 0)
        {
            Paragraph newParagraph = SplitParagraph(originalParagraph, firstRunIndex)!;
            originalParagraph.InsertBeforeSelf(newParagraph);

            // обновим значения после сокращения
            lastRunIndex -= firstRunIndex;
            originalRuns = originalParagraph.Elements<Run>().ToList();
        }

        if (lastRunIndex < (originalRuns.Count() - 1))
        {
            Paragraph newParagraph = SplitParagraph(originalParagraph, lastRunIndex + 1)!;
            originalParagraph.InsertBeforeSelf(newParagraph);
        }
    }

    public static Paragraph? SplitParagraph(Paragraph originalParagraph, int splitRunIndex)
    {
        if (originalParagraph == null || splitRunIndex < 0) 
            return null;

        Paragraph newParagraph = (Paragraph)originalParagraph.CloneNode(true);

        var runs = originalParagraph.Elements<Run>().ToList();
        var newRuns = newParagraph.Elements<Run>().ToList();

        for (int i = 0; i < runs.Count; i++)
        {
            if(i < splitRunIndex)
            {
                runs[i].Remove();
            }
            else
            {
                newRuns[i].Remove();
            }
        }

        return newParagraph;
    }

    //static void MoveTextToProperParagraph(Paragraph paragraph, Run? runToMove = null)
    //{
    //    var parent = paragraph.Parent;
    //    string textToMove = runToMove != null
    //        ? runToMove.GetFirstChild<Text>()?.Text ?? ""
    //        : paragraph.InnerText.Trim();

    //    if (string.IsNullOrEmpty(textToMove)) return;

    //    var previousParagraph = paragraph.PreviousSibling<Paragraph>();

    //    if (previousParagraph != null && !ContainsBlockTag(previousParagraph.InnerText.Trim(),
    //        new Regex(@"\{#(RepeatableBlockStart|BlockStart):([\w\d_]+):([\w\d_]+)\}"),
    //        new Regex(@"\{#(RepeatableBlockEnd|BlockEnd):([\w\d_]+):([\w\d_]+)\}")))
    //    {
    //        // Append content to an existing non-tag paragraph
    //        previousParagraph.Append(new Run(new Text(textToMove)));
    //    }
    //    else
    //    {
    //        // Create a new paragraph for the moved text
    //        var newParagraph = new Paragraph(new Run(new Text(textToMove)));
    //        parent.InsertAfter(newParagraph, paragraph);
    //    }

    //    // Remove the moved text from the original paragraph
    //    if (runToMove != null)
    //    {
    //        runToMove.Remove();
    //    }
    //    else
    //    {
    //        paragraph.RemoveAllChildren<Run>();
    //    }
    //}

}
