using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class DocxTemplateDtoGenerator
{
    public static string baseFolder = @"C:\Users\FeskovichAO\Documents\GitHub\gs-neoland-backend\src\Tap.Zis3.Domain.Services";//AppContext.BaseDirectory

    public static string filePath =
        Path.Combine(baseFolder, "Reports", "Templates", "InformationAnalyticCardForComplexTemplate.docx");

    public static string outputPath =
            Path.Combine(baseFolder, "Reports", "TemplateDtos", "InformationAnalyticCardForComplexTemplateDto.cs");

    public static void Main()
    {
        //DocxTemplateTools.NormalizeDocument(filePath);
        var parsedBlocks = ParseDocx(filePath, "");
        string generatedCode = GenerateClasses(parsedBlocks);

        File.WriteAllText(outputPath, generatedCode);
        Console.WriteLine("DTO class generated successfully!");
    }

    public static List<DocxTemplateBlockDefinition> ParseDocx(string filePath, string dtoName)
    {
        List<DocxTemplateBlockDefinition> blocks = new List<DocxTemplateBlockDefinition>();
        List<DocxTemplateBlockDefinition> tables = new List<DocxTemplateBlockDefinition>();
        Stack<DocxTemplateBlockDefinition> stack = new Stack<DocxTemplateBlockDefinition>();

        string fileName = String.IsNullOrEmpty(dtoName) ? Path.GetFileNameWithoutExtension(filePath) + "Dto" : dtoName;

        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
            var body = doc.MainDocumentPart.Document.Body;
            if (body == null) return blocks;

            var mainBlock = new DocxTemplateBlockDefinition(fileName, "", false);
            stack.Push(mainBlock);

            foreach (var element in body.Elements<OpenXmlElement>())
            {
                string text = element.InnerText.Trim();

                // Detect block start
                var startMatch = DocxTemplateTools.blockStartRegex.Match(text);
                if (startMatch.Success)
                {
                    string structName = startMatch.Groups[2].Value;
                    string propertyName = startMatch.Groups[3].Value;
                    bool isRepeatable = startMatch.Groups[1].Value == "RepeatableBlockStart";

                    DocxTemplateBlockDefinition newBlock = new DocxTemplateBlockDefinition(structName, propertyName, isRepeatable);
                    if (stack.Count > 0)
                        stack.Peek().Children.Add(newBlock);

                    stack.Push(newBlock);
                    continue;
                }

                // Detect block end
                var endMatch = DocxTemplateTools.blockEndRegex.Match(text);
                if (endMatch.Success)
                {
                    string structName = endMatch.Groups[2].Value;
                    string propertyName = endMatch.Groups[3].Value;
                    if (stack.Count > 0 && stack.Peek().Name == structName && stack.Peek().PropertyName == propertyName)
                    {
                        var block = stack.Pop();
                        block.Children.Reverse();
                        block.StringFields.Reverse();
                        blocks.Add(block);
                    }
                    continue;
                }

                // Detect table row properties
                var tableMatch = DocxTemplateTools.tableRowRegex.Match(text);
                if (tableMatch.Success)
                {
                    foreach (TableRow tableRow in ((Table)element).Elements<TableRow>())
                    {
                        //Каждый ряд таблицы может иметь набор тэгов таблицы (не более одного набора)
                        //Если таковой имеется - генерируем соответствующий класс и поле в материнской структуре

                        var rowText = tableRow.InnerText;
                        var tableRowMatches = DocxTemplateTools.tableRowRegex.Matches(rowText);

                        if(!tableRowMatches.Any())
                            continue;

                        var tableTags = tableRowMatches
                            .Select(match => match.Groups[2].Value)
                            .Distinct();

                        if(tableTags.Count() > 1)
                        {
                            throw new Exception("Нарушение разметки! Больше одно набора тэгов таблицы в одном ряду!");
                        }

                        string currentTableClassName = tableRowMatches[0].Groups[2].Value;
                        List<string> cellNames = tableRowMatches.Select(match => match.Groups[3].Value).ToList();

                        var existingTableClass = tables.SingleOrDefault(classDefinition => classDefinition.Name == currentTableClassName);

                        if (existingTableClass != null)
                        {
                            if (cellNames.Except(existingTableClass.StringFields).Any())
                                throw new Exception("Дублирующее описание TableRow с иным набором полей");
                        }

                        var parentBlock = stack.Peek();
                        var tablePropertiesOfThisType = parentBlock.Children.Where(b => b.Name == currentTableClassName);

                        //Добавим индекс таблицы в название на случай если нужно несколько таблиц одного типа в блоке
                        var newTableProperty = new DocxTemplateBlockDefinition(
                            currentTableClassName, 
                            currentTableClassName + "List" + tablePropertiesOfThisType.Count().ToString(),
                            true);

                        foreach (string cellName in tableRowMatches.Select(match => match.Groups[3].Value))
                            newTableProperty.StringFields.Add(cellName);

                        tables.Add(newTableProperty);

                        parentBlock.Children.Add(newTableProperty);
                    }
                }

                // Detect fields inside blocks
                var fieldMatches = DocxTemplateTools.fieldRegex.Matches(text);
                var photoListMatches = DocxTemplateTools.photoFieldRegex.Matches(text);
                if (stack.Count > 0)
                {
                    var fields = stack.Peek().StringFields;                    
                    foreach (Match match in fieldMatches.Reverse())
                    {
                        if (match.Success)
                        {
                            fields.Add(match.Groups[1].Value);
                        }
                    }

                    var photoFields = stack.Peek().PhotoListFields;
                    foreach (Match match in photoListMatches.Reverse())
                    {
                        if (match.Success)
                        {
                            photoFields.Add(match.Groups[2].Value);
                        }
                    }

                }
            }

            blocks.Add(stack.Pop());
        }

        blocks.Reverse();

        blocks = blocks.Concat(tables).ToList();

        foreach(var block in blocks)
        {
            block.StringFields = block.StringFields.Distinct().ToList();
        }

        return blocks;
    }

    public static string GenerateClasses(List<DocxTemplateBlockDefinition> blocks)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections.Generic;");

        foreach (var block in blocks)
        {
            GenerateClass(sb, block, 0);
        }

        return sb.ToString();
    }

    private static void GenerateClass(StringBuilder sb, DocxTemplateBlockDefinition block, int indentLevel)
    {
        sb.AppendLine();
        string indent = new string(' ', indentLevel * 4);
        sb.AppendLine($"{indent}public class {block.Name}");
        sb.AppendLine($"{indent}{{");

        // Fields (Simple text placeholders)
        foreach (var field in block.StringFields)
        {
            sb.AppendLine($"{indent}    public string {field} {{ get; set; }}");
        }

        // Fields (Simple text placeholders)
        foreach (var field in block.PhotoListFields)
        {
            sb.AppendLine($"{indent}    public List<PhotoDto> {field} {{ get; set; }}");
        }

        // Tables (Generate List<T> for table rows)
        foreach (var child in block.Children.Where(c => c.IsRepeatable && c.Name == c.PropertyName))
        {
            sb.AppendLine($"{indent}    public List<{child.Name}> {child.PropertyName} {{ get; set; }} = new List<{child.Name}>();");
        }

        // Nested Blocks (repeatable sections, not tables)
        foreach (var child in block.Children.Where(c => c.IsRepeatable && c.Name != c.PropertyName))
        {
            sb.AppendLine($"{indent}    public List<{child.Name}> {child.PropertyName} {{ get; set; }} = new List<{child.Name}>();");
        }

        // Non-repeatable child blocks
        foreach (var child in block.Children.Where(c => !c.IsRepeatable))
        {
            sb.AppendLine($"{indent}    public {child.Name} {child.PropertyName} {{ get; set; }}");
        }

        sb.AppendLine($"{indent}}}");
    }
}
