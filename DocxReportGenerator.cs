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

class DocxReportGenerator
{
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
        for (int i = 0; i < paragraphs.Count; i++)
        {
            var textElement = paragraphs[i]
                .Elements<Run>()
                .Select(r => r.GetFirstChild<Text>())
                .FirstOrDefault(t => t != null && t.Text.Contains("{#RepeatableBlockStart"));

            if (textElement != null)
            {
                string tagText = textElement.Text.Trim('{', '}');
                string[] parts = tagText.Split(':');
                if (parts.Length == 3 && parts[0] == "#RepeatableBlockStart")
                {
                    string className = parts[1];
                    string propertyName = parts[2];

                    var blockContent = ExtractBlockContent(paragraphs, i, className, propertyName);
                    if (blockContent != null)
                    {
                        // Find the corresponding property in the data object
                        var property = typeof(T).GetProperty(propertyName);
                        if (property != null && typeof(System.Collections.IEnumerable).IsAssignableFrom(property.PropertyType))
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
                                        body.InsertAfter(element, paragraphs[i]); // Insert after the original block
                                    }
                                }
                            }
                        }
                    }

                    // Step 3: Remove tag-paragraphs
                    paragraphs[i].Remove(); // Remove {#RepeatableBlockStart}
                    paragraphs.RemoveAt(i); // Shift index to point at next paragraph

                    var endTagIndex = paragraphs.FindIndex(p => p.InnerText.Contains("{#RepeatableBlockEnd"));
                    if (endTagIndex >= 0)
                    {
                        paragraphs[endTagIndex].Remove(); // Remove {#RepeatableBlockEnd}
                        paragraphs.RemoveAt(endTagIndex);
                    }
                }
            }
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
}
