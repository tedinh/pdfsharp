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

            // Adds title page (need to play around with xgraphics to format the text correctly
            addOverlay(outputDocument);

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
            PdfOutline outline = outputDocument.Outlines.Add("Root", firstPage, true, PdfOutlineStyle.Bold, XColors.Red);

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

                    // Page number text
                    var text = $"{file} Page {idx}";

                    // and add it to the output document.
                    outputDocument.AddPage(page);

                    // Adding bookmarks for each page
                    outline.Outlines.Add(text, firstPage, true);
                }
            }

            // Add page numbers
            addPageNumbers(outputDocument);

            return outputDocument;
        }

        private static void addPageNumbers(PdfDocument outputDocument)
        {
            // Make a font and a brush to draw the page counter.
            XFont font = new XFont("Verdana", 8);
            XBrush brush = XBrushes.Black;

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
                }
            }
        }

        private void addOverlay(PdfDocument combined_pdf)
        {
            PdfDocument document = combined_pdf;

            // Add Title page
            var titlePage = document.InsertPage(0);

            // Select the first page
            PdfPage page = document.Pages[0];

            page.Orientation = PdfSharp.PageOrientation.Portrait;

            XGraphics gfx = XGraphics.FromPdfPage(page, XPageDirection.Downwards);

            // Write on top of background with known colors
            gfx.DrawString("Title Page", new XFont("Helvetica", 12, XFontStyle.Regular), XBrushes.Black, 10, 0, XStringFormats.TopLeft);
        }

        private static string[] LoadFiles()
        {
            return Directory.GetFiles(@"C:\Users\terry.dinh\source\repos\DynamicPDF\files", "*.pdf");
        }
    }
}
