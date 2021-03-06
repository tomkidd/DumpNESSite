﻿using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Xml;

namespace DumpNESSite
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        private void button1_Click(object sender, EventArgs e)
        {
            StringBuilder sbOutput = new StringBuilder();

            string codePoint = "00A0";
            int code = int.Parse(codePoint, System.Globalization.NumberStyles.HexNumber);
            string unicodeNBSP = char.ConvertFromUtf32(code).ToString();

            string tempImage = Path.Combine(System.Environment.CurrentDirectory, "temp.jpg");

            Microsoft.Office.Interop.Word.Application oWord;
            Microsoft.Office.Interop.Word.Document oDoc;

            object m = Type.Missing;

            oWord = new Microsoft.Office.Interop.Word.Application();
            oWord.Visible = true;

            if (File.Exists("test.docx"))
                File.Delete("test.docx");

            oDoc = oWord.Documents.Add(ref m, ref m, ref m, ref m);

            oWord.ActiveDocument.Range(ref m, ref m).NoProofing = 1;

            foreach (string url in txtLinks.Lines)
            {
                if (!string.IsNullOrWhiteSpace(url))
                {
                    HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
                    doc.OptionFixNestedTags = true;
                    doc.LoadHtml(new System.Net.WebClient().DownloadString(url));

                    // Title Node
                    HtmlAgilityPack.HtmlNode titleNode = doc.DocumentNode.SelectSingleNode("/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[4]/div/div/div/div/div/div/div/div/div/h3");
                    string titleNodeHTML = titleNode.OuterHtml;
                    
                    string title = titleNode.InnerText;
                    if (title.Split(new string[] { " - " }, StringSplitOptions.None).Length == 2)
                    {
                        title = title.Split(new string[] { " - " }, StringSplitOptions.None)[1].Trim();
                    }

                    // Content Node
                    HtmlAgilityPack.HtmlNode contentNode = doc.DocumentNode.SelectSingleNode("/html/body/div[3]/div[2]/div[2]/div[2]/div[2]/div[2]/div[2]/div/div[4]/div/div/div/div/div/div/div/div/div/div[2]");
                    string contentNodeHTML = contentNode.OuterHtml;

                    IEnumerable<HtmlAgilityPack.HtmlNode> imgNodes = contentNode.Descendants("img");

                    Dictionary<string, HtmlAgilityPack.HtmlNode> changedNodes = new Dictionary<string, HtmlAgilityPack.HtmlNode>();

                    foreach (HtmlAgilityPack.HtmlNode imgNode in imgNodes)
                    {
                        string width = imgNode.Attributes.Contains("width") ? imgNode.Attributes["width"].Value : null;
                        string height = imgNode.Attributes.Contains("height") ? imgNode.Attributes["height"].Value : null;

                        string src = imgNode.Attributes["src"].Value + ";" + width + ";" + height;
                        changedNodes.Add(src, imgNode);
                    }

                    IEnumerable<HtmlAgilityPack.HtmlNode> spanNodes = contentNode.Descendants("span");

                    foreach (HtmlAgilityPack.HtmlNode spanNode in spanNodes)
                    {
                        string src = spanNode.InnerText;
                        src = src.Replace("&nbsp;", "");
                        src = src.Replace("\n", " ");

                        if (!string.IsNullOrWhiteSpace(src))
                        {
                            if (!changedNodes.ContainsKey(src))
                            {
                                changedNodes.Add(src, spanNode);
                            }
                        }
                    }

                    foreach (KeyValuePair<string, HtmlAgilityPack.HtmlNode> kvp in changedNodes)
                    {
                        HtmlAgilityPack.HtmlNode newNode = HtmlAgilityPack.HtmlNode.CreateNode("<p>" + kvp.Key + "</p>");
                        kvp.Value.ParentNode.ReplaceChild(newNode, kvp.Value);
                    }

                    var scriptNodes = contentNode.Descendants("script").ToList();

                    foreach (var scriptNode in scriptNodes)
                        scriptNode.Remove();

                    string content = contentNode.InnerText;

                    string outputcontent = HtmlToText.ConvertHtml(content).Replace("\n", Environment.NewLine);

                    outputcontent = outputcontent.Replace(Environment.NewLine + Environment.NewLine + Environment.NewLine, Environment.NewLine);

                    sbOutput.AppendLine(title);
                                        
                    string[] outputsplit = outputcontent.Split(new[] { Environment.NewLine }, StringSplitOptions.None);

                    Microsoft.Office.Interop.Word.Paragraph oPara1;
                    oPara1 = oDoc.Content.Paragraphs.Add(ref m);
                    oPara1.Range.Text = title;
                    oPara1.Range.Font.Bold = 1;
                    oPara1.Range.Font.Size = 24;
                    oPara1.Range.InsertParagraphAfter();

                    string prevLine = "";
                    string currentLine = "";

                    var outputList = new List<string>();
                    foreach (string outputLine in outputsplit)
                    {
                        if ((prevLine != "") && (!outputLine.Trim().StartsWith("PLAYERS:") && (!outputLine.Trim().StartsWith("http://"))))
                        {
                            currentLine += " " + outputLine.Trim();
                        }
                        else
                        {
                            if ((outputLine != "") || (outputLine.Trim().StartsWith("http://")))
                            {
                                outputList.Add(currentLine.Trim());
                                currentLine = outputLine.Trim();
                            }
                        }
                         //)
                        if (outputLine.Trim().StartsWith("http://")|| (outputLine.Trim().StartsWith("RELEASE DATE:")))
                        {
                            prevLine = "";
                        }
                        else
                        {
                            prevLine = outputLine;
                        }
                    }
                    outputList.Add(currentLine); // man this is bad

                    bool previousLineImage = false;

                    bool firstImage = false;

                    foreach (string outputline in outputList)
                    {
                        if (!string.IsNullOrWhiteSpace(outputline.Trim()))
                        {
                            Microsoft.Office.Interop.Word.Paragraph oPara2;
                            oPara2 = oDoc.Content.Paragraphs.Add(ref m);

                            if (outputline.StartsWith("http"))
                            {
                                try
                                {
                                    byte[] tempImageData;

                                    using (System.Net.WebClient wc = new System.Net.WebClient())
                                    {
                                        if (File.Exists(tempImage)) File.Delete(tempImage);
                                        tempImageData = wc.DownloadData(outputline.Split(';')[0]);
                                    }

                                    Image tempImageOriginal;

                                    using (var ms = new MemoryStream(tempImageData))
                                    {
                                        tempImageOriginal = Image.FromStream(ms);
                                    }

                                    int width = !string.IsNullOrWhiteSpace(outputline.Trim().Split(';')[1]) ? int.Parse(outputline.Trim().Split(';')[1]) : tempImageOriginal.Width;
                                    int height = !string.IsNullOrWhiteSpace(outputline.Trim().Split(';')[2]) ? int.Parse(outputline.Trim().Split(';')[2]) : tempImageOriginal.Height;

                                    Image tempImageResized = (Image)(new Bitmap(tempImageOriginal, new Size(width, height)));
                                    tempImageResized.Save(tempImage);

                                    oPara2.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;

                                    oPara2.Range.InlineShapes.AddPicture(tempImage, false);

                                    previousLineImage = true;

                                    // Page break after first image, which is almost certainly the cover art
                                    if (firstImage)
                                    {
                                        oDoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                                        firstImage = false;
                                    }
                                }
                                catch { }
                            }
                            else
                            {
                                oPara2.Range.Font.Size = 10;

                                oPara2.Range.Text = outputline.Replace("  ", " ").Trim();

                                if (previousLineImage)
                                {
                                    oPara2.Range.Font.Italic = 1;
                                    oPara2.Range.Font.Size = 8;
                                    oPara2.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphCenter;
                                }
                                else
                                {
                                    oPara2.Range.Font.Italic = 0;
                                    oPara2.Alignment = Microsoft.Office.Interop.Word.WdParagraphAlignment.wdAlignParagraphLeft;
                                }

                                oPara2.Format.SpaceAfter = 10;
                                previousLineImage = false;
                            }

                            oPara2.Range.Font.Bold = 0;
                            oPara2.Range.InsertParagraphAfter();
                        }
                    }

                    oDoc.Words.Last.InsertBreak(Microsoft.Office.Interop.Word.WdBreakType.wdPageBreak);
                }

                txtOutput.Text = sbOutput.ToString();
            }
        }

        public static string FormatXML(string XML)
        {
            string Result = "";

            MemoryStream mStream = new MemoryStream();
            XmlTextWriter writer = new XmlTextWriter(mStream, Encoding.Unicode);
            XmlDocument document = new XmlDocument();

            try
            {
                // Load the XmlDocument with the XML.
                document.LoadXml(XML);

                writer.Formatting = Formatting.Indented;

                // Write the XML into a formatting XmlTextWriter
                document.WriteContentTo(writer);
                writer.Flush();
                mStream.Flush();

                // Have to rewind the MemoryStream in order to read
                // its contents.
                mStream.Position = 0;

                // Read MemoryStream contents into a StreamReader.
                StreamReader sReader = new StreamReader(mStream);

                // Extract the text from the StreamReader.
                String FormattedXML = sReader.ReadToEnd();

                Result = FormattedXML;
            }
            catch (XmlException x)
            {
                MessageBox.Show(x.Message);
            }

            mStream.Close();
            writer.Close();

            return Result;
        }

        private void Form1_Load(object sender, EventArgs e)
        {
            txtLinks.Text = File.ReadAllText("links.txt");
        }

    }

    public class HtmlToText
    {
        public HtmlToText()
        {
        }

        public static string Convert(string path)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.Load(path);

            StringWriter sw = new StringWriter();
            ConvertTo(doc.DocumentNode, sw);
            sw.Flush();
            return sw.ToString();
        }

        public static string ConvertHtml(string html)
        {
            HtmlAgilityPack.HtmlDocument doc = new HtmlAgilityPack.HtmlDocument();
            doc.LoadHtml(html);

            StringWriter sw = new StringWriter();
            ConvertTo(doc.DocumentNode, sw);
            sw.Flush();
            return sw.ToString();
        }

        private static void ConvertContentTo(HtmlAgilityPack.HtmlNode node, TextWriter outText)
        {
            foreach (HtmlAgilityPack.HtmlNode subnode in node.ChildNodes)
            {
                ConvertTo(subnode, outText);
            }
        }

        public static void ConvertTo(HtmlAgilityPack.HtmlNode node, TextWriter outText)
        {
            string html;
            switch (node.NodeType)
            {
                case HtmlAgilityPack.HtmlNodeType.Comment:
                    // don't output comments
                    break;

                case HtmlAgilityPack.HtmlNodeType.Document:
                    ConvertContentTo(node, outText);
                    break;

                case HtmlAgilityPack.HtmlNodeType.Text:
                    // script and style must not be output
                    string parentName = node.ParentNode.Name;
                    if ((parentName == "script") || (parentName == "style"))
                        break;

                    // get text
                    html = ((HtmlAgilityPack.HtmlTextNode)node).Text;

                    // is it in fact a special closing node output as text?
                    if (HtmlAgilityPack.HtmlNode.IsOverlappedClosingElement(html))
                        break;

                    // check the text is meaningful and not a bunch of whitespaces
                    if (html.Trim().Length > 0)
                    {
                        outText.Write(HtmlAgilityPack.HtmlEntity.DeEntitize(html));
                    }
                    break;

                case HtmlAgilityPack.HtmlNodeType.Element:
                    switch (node.Name)
                    {
                        case "p":
                            // treat paragraphs as crlf
                            outText.Write("\r\n");
                            break;
                    }

                    if (node.HasChildNodes)
                    {
                        ConvertContentTo(node, outText);
                    }
                    break;
            }
        }
    }
}
