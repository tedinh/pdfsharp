using System;
using System.IO;
using System.Linq;
using System.Windows.Forms;
using PdfSharp.Pdf;
using PdfSharp.Pdf.IO;
using PdfSharp.Drawing;
using MigraDoc.DocumentObjectModel;
using MigraDoc.DocumentObjectModel.Tables;
using MigraDoc.Rendering;
using MigraDoc.RtfRendering;
using System.Collections.Generic;
using System.Diagnostics;

namespace DynamicPDF
{
    public partial class Form1 : Form
    {

        public Form1()
        {
            InitializeComponent();
            checkedListBox.CheckOnClick = true;
            string[] files = LoadFiles();
            LoadListBox(files);
        }

        private void btnGetItem_Click(object sender, EventArgs e)
        {
            var selectedFiles = GetSelectedFiles();

            PdfDocument outputDocument = StitchPDFFiles(selectedFiles);

            string fileName = textBox2.Text;

            // Names the file
            string filename = $@"C:\Users\terry.dinh\source\repos\DynamicPDF\mergedfiles\{fileName}_tempfile.pdf";

            // Saves the PDF
            outputDocument.Save(filename);
        }

        private string[] GetSelectedFiles()
        {
            List<string> retVal = new List<string>();

            foreach (string itemChecked in checkedListBox.CheckedItems)
            {
                retVal.Add(itemChecked);
            }
            return retVal.ToArray();

            throw new NotImplementedException();
        }

        private void LoadListBox(string[] files)
        {
            // Pulling in PDFs and assigning their values respectively
            checkedListBox.Items.AddRange(files);
        }

        private static PdfDocument StitchPDFFiles(string[] files)
        {
            // Open the output document
            PdfDocument outputDocument = new PdfDocument();

            // Create first page
            var firstPage = outputDocument.AddPage();
            // Create second page
            var secondPage = outputDocument.AddPage();

            Document document = new Document();

            // Iterate files
            foreach (string file in files)
            {
                // Open the document to import pages from it.
                PdfDocument inputDocument = PdfReader.Open(file, PdfDocumentOpenMode.Import);

                // Iterate pages
                int count = inputDocument.PageCount;
                for (int idx = 0; idx < count; idx++)
                {
                    // Get the page from the external document
                    PdfPage page = inputDocument.Pages[idx];

                    // and add it to the output document.
                    outputDocument.AddPage(page);
                }
            }
            // Add title page to first page
            TitlePage(outputDocument);

            // Add page numbers
            addPageNumbers(outputDocument);

            // Add table of contents
            TableOfContents(outputDocument);

            return outputDocument;
        }

        static void TitlePage(PdfDocument document)
        {
            PdfPage page = document.Pages[0];
            XGraphics gfx = XGraphics.FromPdfPage(page);

            XFont font = new XFont("Verdana", 13, XFontStyle.Bold);

            gfx.DrawString("Title Page", font, XBrushes.Black,
              new XRect(100, 100, page.Width - 200, 300), XStringFormats.Center);

            // You always need a MigraDoc document for rendering.
            Document doc = new Document();
            Section sec = doc.AddSection();
            // Add a single paragraph with some text and format information.
            Paragraph para = sec.AddParagraph();
            para.Format.Alignment = ParagraphAlignment.Justify;
            para.Format.Font.Name = "Verdana";
            para.Format.Font.Size = 12;
            para.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            para.Format.Font.Color = MigraDoc.DocumentObjectModel.Colors.Black;
            para.AddText(" ipit iurero dolum zzriliquisis nit wis dolore vel et nonsequipit, velendigna " +
              "auguercilit lor se dipisl duismod tatem zzrit at laore magna feummod oloborting ea con vel " +
              "essit augiati onsequat luptat nos diatum vel ullum illummy nonsent nit ipis et nonsequis " +
              "niation utpat. Odolobor augait et non etueril landre min ut ulla feugiam commodo lortie ex " +
              "essent augait el ing eumsan hendre feugait prat augiatem amconul laoreet. ≤≥≈≠");
            para.Format.Borders.Distance = "5pt";
            para.Format.Borders.Color = Colors.Gold;

            // Create a renderer and prepare (=layout) the document
            MigraDoc.Rendering.DocumentRenderer docRenderer = new DocumentRenderer(doc);
            docRenderer.PrepareDocument();

            // Render the paragraph. You can render tables or shapes the same way.
            docRenderer.RenderObject(gfx, XUnit.FromCentimeter(5), XUnit.FromCentimeter(10), "12cm", para);

            gfx.Dispose();
        }

        static void TableOfContents(PdfDocument document)
        {
            // Puts the Table of contents on the second page
            PdfPage page = document.Pages[1];
            XGraphics gfx = XGraphics.FromPdfPage(page);
            gfx.MUH = PdfFontEncoding.Unicode;

            // Create MigraDoc document + Setup styles
            Document doc = new Document();
            Styles.DefineStyles(doc);

            // Add header
            Section section = doc.AddSection();
            Paragraph paragraph = section.AddParagraph("Table of Contents");
            paragraph.Format.Font.Size = 14;
            paragraph.Format.Font.Bold = true;
            paragraph.Format.SpaceAfter = 24;
            paragraph.Format.OutlineLevel = OutlineLevel.Level1;


            // Add links - these are the PdfSharp outlines/bookmarks added previously when concatinating the pages
            foreach (var bookmark in document.Outlines)
            {
                paragraph = section.AddParagraph();
                paragraph.Style = "TOC";
                //paragraph.AddBookmark(bookmark.Title);
                Hyperlink hyperlink = paragraph.AddHyperlink(bookmark.Title);
                hyperlink.AddText($"{bookmark.Title}\t");
                //hyperlink.AddPageRefField(bookmark.Title);
            }

            // Render document
            DocumentRenderer docRenderer = new DocumentRenderer(doc);
            docRenderer.PrepareDocument();
            docRenderer.RenderPage(gfx, 1);

            gfx.Dispose();
        }

        private static void addPageNumbers(PdfDocument outputDocument)
        {
            // Make a font and a brush to draw the page counter.
            XFont font = new XFont("Verdana", 8);
            XBrush brush = XBrushes.Black;

            // Need to pull in selected document names


            // Add the page counter.
            string noPages = outputDocument.Pages.Count.ToString();

            for (int i = 0; i < outputDocument.Pages.Count; ++i)
            {
                PdfPage page = outputDocument.Pages[i];

                // Make a layout rectangle.
                XRect layoutRectangle = new XRect(0/*X*/, page.Height - font.Height/*Y*/, page.Width/*Width*/, font.Height/*Height*/);

                using (XGraphics gfx = XGraphics.FromPdfPage(page))
                {
                    gfx.DrawString(
                        "Page " + (i + 1).ToString() + " of " + noPages,
                        font,
                        brush,
                        layoutRectangle,
                        XStringFormats.TopCenter);
                    gfx.Dispose();
                }

                // Table of contents file name
                var text = $"file name here";

                // Adding bookmarks for each page
                outputDocument.Outlines.Add(text, page, true);
            }
            
        }

        private static string[] LoadFiles()
        {
            return Directory.GetFiles(@"C:\Users\terry.dinh\source\repos\DynamicPDF\files", "*.pdf");
        }
    }
}
