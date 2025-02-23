using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;

class DocxStructGenerator
{
    public static void Main()
    {
        string filePath = "C:\\Users\\FeskovichAO\\Documents\\GitHub\\gs-neoland-backend\\src\\Tap.Zis3.Domain.Services\\Reports\\Templates\\InformationAnalyticCardForComplexTemplate.docx"; // Change to your actual template path
        string outputPath = "C:\\Users\\FeskovichAO\\Documents\\GitHub\\gs-neoland-backend\\src\\Tap.Zis3.Domain.Services\\Reports\\Templates\\IACReportDtos.cs";

        var parsedBlocks = ParseDocx(filePath, "");
        string generatedCode = GenerateDtoClass(parsedBlocks);

        File.WriteAllText(outputPath, generatedCode);
        Console.WriteLine("DTO class generated successfully!");
    }

    public static List<BlockDefinition> ParseDocx(string filePath, string dtoName)
    {
        List<BlockDefinition> blocks = new List<BlockDefinition>();
        Stack<BlockDefinition> stack = new Stack<BlockDefinition>();

        string fileName = String.IsNullOrEmpty(dtoName) ? Path.GetFileNameWithoutExtension(filePath) + "Dto" : dtoName;

        using (WordprocessingDocument doc = WordprocessingDocument.Open(filePath, false))
        {
            var body = doc.MainDocumentPart.Document.Body;
            if (body == null) return blocks;

            Regex blockStartRegex = new Regex(@"\{#(RepeatableBlockStart|BlockStart):([\w\d_]+):([\w\d_]+)\}");
            Regex blockEndRegex = new Regex(@"\{#(RepeatableBlockEnd|BlockEnd):([\w\d_]+):([\w\d_]+)\}");
            Regex fieldRegex = new Regex(@"\{([\w\d_]+)\}");
            Regex tableRowRegex = new Regex(@"\{#TableRow:([\w\d_]+)\.([\w\d_]+)\}");

            var mainBlock = new BlockDefinition(fileName, "", false);
            stack.Push(mainBlock);

            foreach (var para in body.Descendants<Paragraph>())
            {
                string text = para.InnerText.Trim();

                // Detect block start with property name
                var startMatch = blockStartRegex.Match(text);
                if (startMatch.Success)
                {
                    string structName = startMatch.Groups[2].Value;
                    string propertyName = startMatch.Groups[3].Value;
                    bool isRepeatable = startMatch.Groups[1].Value == "RepeatableBlockStart";

                    BlockDefinition newBlock = new BlockDefinition(structName, propertyName, isRepeatable);
                    if (stack.Count > 0)
                        stack.Peek().Children.Add(newBlock);

                    stack.Push(newBlock);
                    continue;
                }

                // Detect block end with property name
                var endMatch = blockEndRegex.Match(text);
                if (endMatch.Success)
                {
                    string structName = endMatch.Groups[2].Value;
                    string propertyName = endMatch.Groups[3].Value;
                    if (stack.Count > 0 && stack.Peek().Name == structName && stack.Peek().PropertyName == propertyName)
                    {
                        var block = stack.Pop();
                        block.Children.Reverse();
                        block.Fields.Reverse();
                        blocks.Add(block);
                    }
                    continue;
                }

                // Detect table row
                var tableMatch = tableRowRegex.Match(text);
                if (tableMatch.Success)
                {
                    string structName = tableMatch.Groups[1].Value;
                    string cellName = tableMatch.Groups[2].Value;

                    var parentBlock = stack.Peek();
                    var tableBlock = parentBlock.Children.FirstOrDefault(b => b.Name == structName);
                    if (tableBlock == null)
                    {
                        tableBlock = new BlockDefinition(structName, structName + "List", true);
                        parentBlock.Children.Add(tableBlock);
                    }
                    tableBlock.Fields.Add(cellName);
                    continue;
                }

                // Detect fields inside blocks
                var fieldMatch = fieldRegex.Match(text);
                if (fieldMatch.Success && stack.Count > 0)
                {
                    stack.Peek().Fields.Add(fieldMatch.Groups[1].Value);
                }
            }

            blocks.Add(stack.Pop());
        }

        blocks.Reverse();

        return blocks;
    }

    public static string GenerateDtoClass(List<BlockDefinition> blocks)
    {
        StringBuilder sb = new StringBuilder();
        sb.AppendLine("using System;");
        sb.AppendLine("using System.Collections.Generic;");

        foreach (var block in blocks)
        {
            GenerateStruct(sb, block, 0);
        }

        return sb.ToString();
    }

    private static void GenerateStruct(StringBuilder sb, BlockDefinition block, int indentLevel)
    {
        sb.AppendLine();
        string indent = new string(' ', indentLevel * 4);
        sb.AppendLine($"{indent}public class {block.Name}");
        sb.AppendLine($"{indent}{{");

        // Fields
        foreach (var field in block.Fields)
        {
            sb.AppendLine($"{indent}    public string {field} {{ get; set; }}");
        }

        // Tables (nested structures for table rows)
        foreach (var child in block.Children.Where(c => c.IsRepeatable && c.Name == c.PropertyName))
        {
            sb.AppendLine($"{indent}    public List<{child.Name}> {child.PropertyName} {{ get; set; }} = new List<{child.Name}>();");
        }

        // Nested Blocks (non-table repeatable blocks)
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

        // Generate nested structs
        //foreach (var child in block.Children)
        //{
        //    GenerateStruct(sb, child, indentLevel);
        //}
    }
}

class BlockDefinition
{
    public string Name { get; }
    public string PropertyName { get; } // The field name in the parent DTO
    public bool IsRepeatable { get; }
    public List<string> Fields { get; } = new List<string>();
    public List<BlockDefinition> Children { get; } = new List<BlockDefinition>();

    public BlockDefinition(string name, string propertyName, bool isRepeatable)
    {
        Name = name;
        PropertyName = propertyName;
        IsRepeatable = isRepeatable;
    }
}
