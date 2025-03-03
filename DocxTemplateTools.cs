
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using DocumentFormat.OpenXml;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

public static class DocxTemplateTools
{
    public static Regex blockStartRegex = new Regex(@"\{#(RepeatableBlockStart|BlockStart):([\w\d_]+):([\w\d_]+)\}");
    public static Regex blockEndRegex = new Regex(@"\{#(RepeatableBlockEnd|BlockEnd):([\w\d_]+):([\w\d_]+)\}");
    public static Regex blockTagStartRegex = new Regex(@"\{#");
    public static Regex blockTagEndRegex = new Regex(@"\}");
    public static Regex tableRowRegex = new Regex(@"\{#(TableRow):([\w\d_]+):([\w\d_]+)\}");

    public static Regex fieldRegex = new Regex(@"\{([\w\d_]+)\}");
    public static Regex photoFieldRegex = new Regex(@"\{(Photos):([\w\d_]+)\}");

    public static void NormalizeDocument(string filePath)
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
            else if (firstRunIndex != -1)
            {
                break;
            }

            currentCharIndex += runLength;
        }

        lastRunIndex = i - 1;

        return true;
    }

    private static void NormalizeParagraph(Paragraph paragraph)
    {
        while (FindRunsContainingTag(paragraph, out int firstRunIndex, out int lastRunIndex, blockStartRegex, blockEndRegex)
                && ((firstRunIndex > 0)
                    || (lastRunIndex < (paragraph.Descendants<Run>().Count() - 1))))
        {
            ExtractTagParagraph(paragraph, firstRunIndex, lastRunIndex);
        }
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
    private static void ExtractTagParagraph(Paragraph originalParagraph, int firstRunIndex, int lastRunIndex)
    {
        List<Run> originalRuns;

        if (originalParagraph == null || firstRunIndex < 0 || firstRunIndex > lastRunIndex)
        {
            return;
        }
        else
        {
            originalRuns = originalParagraph.Elements<Run>().ToList();

            if (lastRunIndex >= originalRuns.Count())
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

    private static Paragraph? SplitParagraph(Paragraph originalParagraph, int splitRunIndex)
    {
        if (originalParagraph == null || splitRunIndex < 0)
            return null;

        Paragraph newParagraph = (Paragraph)originalParagraph.CloneNode(true);
        //newParagraph. = originalParagraph.Parent;

        var runs = originalParagraph.Elements<Run>().ToList();
        var newRuns = newParagraph.Elements<Run>().ToList();

        for (int i = 0; i < runs.Count; i++)
        {
            if (i < splitRunIndex)
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

}