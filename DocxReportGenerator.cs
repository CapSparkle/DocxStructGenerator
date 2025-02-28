using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Office2010.Word.DrawingGroup;
using System.Xml.Linq;
using LanguageExt;

class DocxReportGenerator
{
    public static void Main()
    {
        string templatePath = DocxTemplateDtoGenerator.filePath;
        string outputPath = Path.Combine(DocxTemplateDtoGenerator.baseFolder, "Reports", "Templates", "FILLEDInformationAnalyticCardForComplexTemplate.docx"); ;

        var data = IACDataGenerator.GenerateExampleData();

        FillDocx(templatePath, outputPath, data);
        Console.WriteLine("Document filled successfully!");
    }

    public static void FillDocx<T>(string templatePath, string outputPath, T data)
    {
        File.Copy(templatePath, outputPath, true);

        DocxTemplateTools.NormalizeDocument(outputPath);

        using (WordprocessingDocument document = WordprocessingDocument.Open(outputPath, true))
        {
            var mainPart = document.MainDocumentPart;
            if (mainPart?.Document?.Body != null)
            {
                var body = mainPart.Document.Body;

                // Step 1: Identify and Process Repeatable Blocks
                var bodyElements = body.Elements<OpenXmlElement>().ToList();

                bool filterElems = true;

                if (filterElems)
                {
                    bodyElements = bodyElements
                        .Where(elem => ((elem is Paragraph) || (elem is Table)))
                        .ToList();
                }
                else
                {
                    var anomalElems = bodyElements.Where(elem => (!(elem is Paragraph) && !(elem is Table)));
                    if (anomalElems.Count() > 0)
                        throw new Exception("Непредусмотренные XML элементы в теле документа: "
                            + anomalElems.First().GetType().Name
                            + ((anomalElems.Count() > 1) ? ", ..." : ""));
                }


                ProcessRepeatableBlocks(bodyElements, data);

                // Step 2: Replace all placeholders in remaining text
                //ReplacePlaceholders(body, data);

                document.Save();
            }
        }
    }

    private static void ProcessRepeatableBlocks(List<OpenXmlElement> docElements, object data)
    {
        for (int i = 0; i < docElements.Count; i++)
        {
            OpenXmlElement docElement = docElements[i];

            // Check if this is a tag-paragraph (Start of Block or RepeatableBlock)

            Paragraph paragraph = docElement as Paragraph;

            if (paragraph != null 
                && DocxTemplateTools.FindRunsContainingTag(paragraph, out _, out _, DocxTemplateTools.blockStartRegex, DocxTemplateTools.blockEndRegex))
            {
                // Это tag-paragraph

                // Extract repeatable block metadata
                string tagText = DocxTemplateTools.blockStartRegex.Match(docElement.InnerText).Value.Trim('{', '}');
                string[] parts = tagText.Split(':');

                //# RepeatableBlockStart
                //

                if (!(parts.Length == 3 && (parts[0] == "#BlockStart" || parts[0] == "#RepeatableBlockStart")))
                    throw new Exception("Нарушение синтаксиса тэга параграфа");

                string className = parts[1];
                string propertyName = parts[2];

                // Find the corresponding end tag paragraph
                var endTagParagraph = docElements.FirstOrDefault(
                    p => p.InnerText.Contains("{#BlockEnd:" + className) 
                    || p.InnerText.Contains("{#RepeatableBlockEnd:" + className)) 
                        as Paragraph;

                if (endTagParagraph == null)
                    throw new Exception("Template markup syntax error");

                // Extract all paragraphs between start & end tags
                var blockContent = ExtractBlockContent(docElements, paragraph, endTagParagraph);
                if (blockContent.Count == 0)  
                    continue;

                // Find the property in the data object
                var property = data.GetType().GetProperty(propertyName);
                if (property is null)
                    throw new Exception("Template markup syntax error");

                // If it's a simple block (single object)
                if (!typeof(System.Collections.IEnumerable).IsAssignableFrom(property.PropertyType))
                {
                    var value = property.GetValue(data);
                    if (value != null)
                    {
                        ProcessRepeatableBlocks(blockContent, value);
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

                            ProcessRepeatableBlocks(blockContent, current);

                            while (enumerator.MoveNext()) 
                            {
                                current = enumerator.Current;

                                var lastBlockElement = blockContent.Last();
                                foreach (var element in clonedBlock)
                                {
                                    lastBlockElement.Parent.InsertBefore(element, lastBlockElement);
                                    lastBlockElement = element;
                                }

                                blockContent = clonedBlock;

                                clonedBlock = blockContent.Select(el => el.CloneNode(true)).ToList();

                                ProcessRepeatableBlocks(blockContent, current);
                            }
                        }
                    }
                }

                i = docElements.IndexOf(endTagParagraph);

                // Remove tag-paragraphs (start & end)
                docElement.Remove();
                endTagParagraph.Remove();
            }
            else
            {
                // Это table, либо это обычный paragraph (не tag-paragraph)
                ReplacePlaceholders(docElement, data);
            }
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

    /// <summary>
    /// Заполнение обычных тэгов и таблиц в множестве OpenXMLElement -ов
    /// </summary>
    /// <param name="elements"></param>
    /// <param name="data"></param>
    /// <exception cref="Exception"></exception>
    private static void ReplacePlaceholders(OpenXmlElement element, object data, bool allowRow = false)
    {
        Dictionary<string, string> replacements = CollectPropertyValues(data);
        Dictionary<string, int> tableCounters = new Dictionary<string, int>();

        //element либо Table либо Paragraph
        if (!(element is Paragraph) && !(element is Table) && !(allowRow && (element is TableRow)))
            throw new Exception("Неверное использование функции ReplacePlaceholders()");
        
        //Вставка обычных тэгов
        foreach (var key in replacements.Keys)
        {
            var keyTag = $"{{{key}}}";

            int i = 0, maxIterations = 1000; 
            
            if (element.InnerText.Contains("ConclusionsFromPreliminaryAnalysis"))
            {
                var y = 0;
            }

            for (i = 0; (i < maxIterations) && (element.InnerText.Contains(keyTag)); i++)
            {
                //{ключ} может оказаться разбитым по нескольким Text или даже Run
                //Поэтому явно вынесем его в один Text одного Run
                
                if(key == "ConclusionsFromPreliminaryAnalysis")
                {
                    var y = 0;
                }

                var run = DocxTemplateTools.MoveTagIntoSingleTextOfRun(element.Descendants<Run>().ToList(), keyTag);

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
                var matches = DocxTemplateTools.tableRowRegex.Matches(tableRow.InnerText);

                var tableTags = matches
                    .Select(match => match.Groups[2].Value)
                    .Distinct();

                if (tableTags.Count() > 1)
                {
                    throw new Exception("Нарушение разметки! Больше одно набора тэгов таблицы в одном ряду!");
                }
                else if (tableTags.Count() == 0)
                    continue;

                // Extract property name from the tag
                string propertyName = tableTags.Single(); //((Match)matchGroups.First().First()).Value.Trim('{', '}').Split(":")[1];

                if (tableCounters.ContainsKey(propertyName))
                    tableCounters[propertyName] += 1;
                else
                    tableCounters[propertyName] = 0;

                propertyName += ("List" + tableCounters[propertyName].ToString());

                var type = data.GetType();
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
                    ReplacePlaceholders(newRow, item, allowRow: true);
                    lastInsertedRow.InsertAfterSelf(newRow);
                    lastInsertedRow = newRow; // Update last inserted row reference
                }

                // Remove the original template row with placeholders
                tableRow.Remove();
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

            if (property.Name == "ConclusionsFromPreliminaryAnalysis")
            {
                var y = 0;
            }

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
}
