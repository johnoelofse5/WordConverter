using System;
using System.IO;
using System.Linq;
using System.Text.RegularExpressions;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using HtmlAgilityPack;

namespace WordToHtmlConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            while (true)
            {
                Console.WriteLine("Enter the full path of the word document to convert:");
                string inputFilePath = Console.ReadLine();

                if (!File.Exists(inputFilePath))
                {
                    Console.WriteLine("File not found. Please check the path and try again.");
                    continue;
                }

                string outputDirectory = Path.GetDirectoryName(inputFilePath);
                string outputFileName = Path.Combine(outputDirectory, Path.GetFileNameWithoutExtension(inputFilePath) + ".html");
                string outputFilePath = Path.Combine(outputDirectory, outputFileName);
            
                ConvertWordToHtml(inputFilePath, outputFilePath);
                Console.WriteLine($"Conversion completed. The HTML File is saved at: {outputFilePath}");
            }
            
        }

        static void ConvertWordToHtml(string inputFilePath, string outputFilePath)
        {
            using (WordprocessingDocument doc = WordprocessingDocument.Open(inputFilePath, false))
            {
                Body body = doc.MainDocumentPart.Document.Body;
                HtmlDocument htmlDoc = new HtmlDocument();
                var root = HtmlNode.CreateNode("<html><body></body></html>");
                htmlDoc.DocumentNode.AppendChild(root);

                HtmlNode currentListNode = null;
                string currentListType = null;

                foreach (var element in body.Elements())
                {
                    HtmlNode htmlNode = null;

                    if (element is Paragraph paragraph)
                    {
                        // Handle Lists
                        var isListItem = paragraph.ParagraphProperties?.NumberingProperties != null;
                        if (isListItem)
                        {
                            var numId = paragraph.ParagraphProperties.NumberingProperties.NumberingId.Val.Value;
                            var listItemStyle = GetListItemStyle(doc, numId);

                            if (currentListType != listItemStyle)
                            {
                                if (currentListNode != null)
                                {
                                    root.SelectSingleNode("//body").AppendChild(currentListNode);
                                }

                                currentListType = listItemStyle;
                                currentListNode = HtmlNode.CreateNode(currentListType == "ordered" ? "<ol></ol>" : "<ul></ul>");
                            }

                            htmlNode = HtmlNode.CreateNode($"<li>{paragraph.InnerText}</li>");
                            currentListNode.AppendChild(htmlNode);
                        }
                        else
                        {
                            // If the paragraph is not a list item, close any open list
                            if (currentListNode != null)
                            {
                                root.SelectSingleNode("//body").AppendChild(currentListNode);
                                currentListNode = null;
                                currentListType = null;
                            }

                            var paragraphText = paragraph.InnerText;

                            // Detect Heading
                            if (paragraph.ParagraphProperties?.ParagraphStyleId?.Val != null)
                            {
                                var style = paragraph.ParagraphProperties.ParagraphStyleId.Val.Value;
                                if (style.StartsWith("Heading"))
                                {
                                    int level = int.Parse(style.Replace("Heading", ""));
                                    htmlNode = HtmlNode.CreateNode($"<h{level}>{paragraphText}</h{level}>");
                                }
                                else
                                {
                                    htmlNode = HtmlNode.CreateNode($"<p>{paragraphText}</p>");
                                }
                            }
                            else
                            {
                                htmlNode = HtmlNode.CreateNode($"<p>{paragraphText}</p>");
                            }

                            root.SelectSingleNode("//body").AppendChild(htmlNode);
                        }
                    }
                    else if (element is Table table)
                    {
                        htmlNode = HtmlNode.CreateNode("<table></table>");
                        foreach (var row in table.Elements<TableRow>())
                        {
                            var rowNode = HtmlNode.CreateNode("<tr></tr>");
                            foreach (var cell in row.Elements<TableCell>())
                            {
                                var cellNode = HtmlNode.CreateNode($"<td>{cell.InnerText}</td>");
                                rowNode.AppendChild(cellNode);
                            }
                            htmlNode.AppendChild(rowNode);
                        }
                        root.SelectSingleNode("//body").AppendChild(htmlNode);
                    }
                }

                // Append any remaining open list
                if (currentListNode != null)
                {
                    root.SelectSingleNode("//body").AppendChild(currentListNode);
                }
                
                string rawHtmlContent = htmlDoc.DocumentNode.OuterHtml;
                
                string pattern = "[A-Za-z0-9+/=]{30,}"; // Matches base64 strings of at least 30 characters
                string cleanedHtmlContent = Regex.Replace(rawHtmlContent, pattern, "");
                    
                File.WriteAllText(outputFilePath, cleanedHtmlContent);
            }
        }

        static string GetListItemStyle(WordprocessingDocument doc, int numId)
        {
            var numberingPart = doc.MainDocumentPart.NumberingDefinitionsPart;
            var abstractNumId = numberingPart.Numbering.Elements<NumberingInstance>()
                .FirstOrDefault(n => n.NumberID == numId)?.AbstractNumId?.Val?.Value;
            
            var abstractNum = numberingPart.Numbering.Elements<AbstractNum>()
                .FirstOrDefault(a => a.AbstractNumberId == abstractNumId);
            var listItemType = abstractNum?.MultiLevelType?.Val?.Value;

            if (listItemType.HasValue)
            {
                if(listItemType == MultiLevelValues.SingleLevel || listItemType == MultiLevelValues.HybridMultilevel)
                {
                    return "ul";
                }
            }
            return "ol";
        }
    }
}