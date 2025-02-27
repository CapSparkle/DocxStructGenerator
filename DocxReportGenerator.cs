using DocumentFormat.OpenXml.Packaging;
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
using System.Xml.Linq;
using LanguageExt;

class DocxReportGenerator
{
    static Regex blockStartRegex = new Regex(@"\{#(RepeatableBlockStart|BlockStart):([\w\d_]+):([\w\d_]+)\}");
    static Regex blockEndRegex = new Regex(@"\{#(RepeatableBlockEnd|BlockEnd):([\w\d_]+):([\w\d_]+)\}");
    static Regex blockTagStartRegex = new Regex(@"\{#");
    static Regex blockTagEndRegex = new Regex(@"\}");

    static Regex tableRowRegex = new Regex(@"\{#(tableRow):([\w\d_]+):([\w\d_]+)\}");

    public static void Main()
    {
        string templatePath = DocxStructGenerator.filePath;
        string outputPath = Path.Combine(DocxStructGenerator.baseFolder, "Reports", "Templates", "FILLEDInformationAnalyticCardForComplexTemplate.docx"); ;

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
                //ReplacePlaceholders(body, data);

                document.Save();
            }
        }
    }

    private static void ProcessRepeatableBlocks<T>(Body body, T data)
    {
        var docElements = body.Elements<OpenXmlElement>().ToList();
        int i = 0;

        for (; i < docElements.Count; i++)
        {
            OpenXmlElement docElement = docElements[i];

            // Check if this is a tag-paragraph (Start of Block or RepeatableBlock)

            Paragraph paragraph = docElement as Paragraph;

            //if (paragraph is null) {
            //    //fill table
            //    continue;
            //}

            if (paragraph != null 
                && FindRunsContainingTag(paragraph, out _, out _, blockStartRegex, blockEndRegex))
            {
                // Extract repeatable block metadata
                string tagText = docElement.InnerText.Trim('{', '}');
                string[] parts = tagText.Split(':');

                if (parts.Length == 3 && (parts[0] == "#BlockStart" || parts[0] == "#RepeatableBlockStart"))
                {
                    string className = parts[1];
                    string propertyName = parts[2];

                    // Find the corresponding end tag paragraph
                    var endTagParagraph = docElements.FirstOrDefault(
                        p => p.InnerText.Contains("{#BlockEnd:" + className) 
                        || p.InnerText.Contains("{#RepeatableBlockEnd:" + className)) 
                            as Paragraph;

                    if (endTagParagraph == null)
                    {
                        throw new Exception("Template markup syntax error");
                    }

                    // Extract all paragraphs between start & end tags
                    var blockContent = ExtractBlockContent(docElements, paragraph, endTagParagraph);
                    if (blockContent.Count == 0)  
                        continue;

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
                            }
                        }
                        // If it's a repeatable block (list of objects)
                        else
                        {
                            System.Collections.IEnumerable list = (System.Collections.IEnumerable)property.GetValue(data);

                            if (list != null)
                            {
                                var enumerator = list.GetEnumerator();
                                enumerator.MoveNext();
                                object current;

                                if (enumerator.Current != null)
                                {
                                    current = enumerator.Current;

                                    var clonedBlock = blockContent.Select(el => el.CloneNode(true))
                                        .ToList();
                                    
                                    ReplacePlaceholders(blockContent, current);

                                    while (enumerator.MoveNext()) 
                                    {
                                        current = enumerator.Current;

                                        var lastBlockElement = blockContent.Last();
                                        foreach (var element in clonedBlock)
                                        {
                                            lastBlockElement.Parent.InsertBefore(element, lastBlockElement);
                                            //lastBlockElement.InsertAfterSelf(element);
                                            lastBlockElement = element;
                                        }

                                        blockContent = clonedBlock;

                                        clonedBlock = blockContent.Select(el => el.CloneNode(true)).ToList();

                                        ReplacePlaceholders(blockContent, current);
                                    }
                                }
                            }
                        }
                    }

                    // Remove tag-paragraphs (start & end)
                    docElement.Remove();
                    endTagParagraph.Remove();

                    // Skip to the next paragraph (to avoid reprocessing removed elements)
                    continue;
                }
            }

            // If it's not a tag-paragraph, replace placeholders for the main object
            ReplacePlaceholders(new List<OpenXmlElement> { docElement }, data);
        }
    }

    private static List<OpenXmlElement> ExtractBlockContent(List<OpenXmlElement> docElements, Paragraph startTag, Paragraph endTag)
    {
        List<OpenXmlElement> blockContent = new List<OpenXmlElement>();
        bool insideBlock = false;

        foreach (var element in docElements)
        {
            if (element == startTag)
            {
                insideBlock = true;
                continue; // Skip the start tag itself
            }

            if (element == endTag)
            {
                break; // Stop when reaching the end tag
            }

            if (insideBlock)
            {
                blockContent.Add(element); // Clone to avoid modifying original references
            }
        }

        return blockContent;
    }


    private static void ReplacePlaceholders(IEnumerable<OpenXmlElement> elements, object data)
    {
        Dictionary<string, string> replacements = CollectPropertyValues(data);

        foreach (var element in elements)
        {
            //element либо Table либо Paragraph
            
            //Вставка обычных тэгов
            foreach (var key in replacements.Keys)
            {
                var keyTag = $"{{{key}}}";

                int i = 0, maxIterations = 1000;
                for (i = 0; (i < maxIterations) && (element.InnerText.Contains(keyTag)); i++)
                {
                    //{ключ} может оказаться разбитым по нескольким Text или даже Run
                    //Поэтому явно вынесем его в один Text одного Run
                    var run = MoveTagIntoSingleTextOfRun(element.Descendants<Run>().ToList(), keyTag);

                    var text = run.Elements<Text>().First();

                    if (text.Text.Contains($"{{{key}}}"))
                    {
                        text.Text = text.Text.Replace($"{{{key}}}", replacements[key]);
                    }
                    else
                    {
                        throw new Exception("Сбой в работе алгоритма объединения тэгов");
                    }
                }

                if(i == maxIterations)
                {
                    throw new Exception("Выполнено прерывание вечного цикла");
                }
            }

            //Вставка табличных значений
            if(element is Table table)
            {
                var tableRows = table.Elements<TableRow>();
                foreach (TableRow tableRow in tableRows)
                {
                    var matches = tableRowRegex.Matches(tableRow.InnerText);

                    var matchGroups = matches.GroupBy(x => {
                        var pieces = x.Value.Split(':');
                        return pieces[1];
                    });

                    if(matchGroups.Count() > 1)
                        throw new Exception("Нарушение разметки - в одной строке таблицы более одного типа тэга");
                    else if(matchGroups.Count() == 0)
                        continue;

                    // Extract property name from the tag
                    string propertyName = ((Match)matchGroups.First());

                    var property = data.GetType().GetProperty(propertyName);
                    if (property == null || !typeof(System.Collections.IEnumerable).IsAssignableFrom(property.PropertyType))
                        throw new Exception($"Property '{propertyName}' not found or is not a list.");

                    var list = (System.Collections.IEnumerable)property.GetValue(data);
                    if (list == null) continue;

                    // Store the last row where new rows should be inserted after
                    OpenXmlElement lastInsertedRow = tableRow;

                    foreach (var item in list)
                    {
                        var newRow = (TableRow)tableRow.CloneNode(true);
                        ReplacePlaceholders(new List<OpenXmlElement> { newRow }, item);
                        lastInsertedRow.InsertAfterSelf(newRow);
                        lastInsertedRow = newRow; // Update last inserted row reference
                    }

                    // Remove the original template row with placeholders
                    tableRow.Remove();
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
            if (property.PropertyType == typeof(string))
            {
                values[property.Name] = property.GetValue(obj)?.ToString() ?? "";

                if (values[property.Name] == $"{{{property.Name}}}")
                {
                    throw new Exception("недопустимое значение - якорь равняется вставляемому значению");
                }
            }

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
        while (FindRunsContainingTag(paragraph, out int firstRunIndex, out int lastRunIndex, blockStartRegex, blockEndRegex)
                && ((firstRunIndex > 0)
                    || (lastRunIndex < (paragraph.Descendants<Run>().Count() - 1))))
        {
            ExtractTagParagraph(paragraph, firstRunIndex, lastRunIndex);
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
        int i = 0;
        for (; i < runs.Count; i++)
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
                break;
            }

            currentCharIndex += runLength;
        }

        lastRunIndex = i - 1;

        return true;
    }

    public static Run MoveTagIntoSingleTextOfRun(List<Run> runs, string tag)
    {
        List<OpenXmlElement> elementsWithTag = new();
        StringBuilder allTextsSymbols = new StringBuilder("");

        var elements = runs.SelectMany(run => run.Elements<OpenXmlElement>())
            .Where(elem => (elem is Text) || (elem is SymbolChar))
            .ToList();

        if (elements.Count() == 0)
            throw new Exception("No tag inside runs");

        int j = 0;
        for (; !allTextsSymbols.ToString().Contains(tag); j++)
        {
            elementsWithTag.Add(elements[j]);
            allTextsSymbols.Append(elements[j].InnerText);
        }

        for (j = 0; allTextsSymbols.ToString().Contains(tag); j++)
        {
            elementsWithTag.RemoveAt(0);
            allTextsSymbols.Remove(0, elements[j].InnerText.Length);
        }

        j--;


        elementsWithTag.Add(elements[j]);
        allTextsSymbols.Insert(0, elements[j].InnerText);

        if (elementsWithTag.Count() == 0)
            throw new Exception("No tag inside runs");

        Run runToPutAfter = (Run)elements[j].Parent!;
        Run newRun = (Run)runToPutAfter.CloneNode(false);

        runToPutAfter.Parent!.InsertAfter(newChild: newRun, runToPutAfter);

        foreach (var element in elementsWithTag)
        {
            element.Remove();
            //newRun.AppendChild(element);
        }

        newRun.AppendChild<Text>(new Text(allTextsSymbols.ToString()));

        return newRun;
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
public static void ExtractTagParagraph(Paragraph originalParagraph, int firstRunIndex, int lastRunIndex)
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
            originalParagraph.Parent.InsertBefore(newParagraph, originalParagraph);

            // обновим значения после сокращения
            lastRunIndex -= firstRunIndex;
            originalRuns = originalParagraph.Elements<Run>().ToList();
        }

        if (lastRunIndex < (originalRuns.Count() - 1))
        {
            Paragraph newParagraph = SplitParagraph(originalParagraph, lastRunIndex + 1)!;
            originalParagraph.Parent.InsertBefore(newParagraph, originalParagraph);
        }
    }

    public static Paragraph? SplitParagraph(Paragraph originalParagraph, int splitRunIndex)
    {
        if (originalParagraph == null || splitRunIndex < 0) 
            return null;

        Paragraph newParagraph = (Paragraph)originalParagraph.CloneNode(true);
        //newParagraph. = originalParagraph.Parent;

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
