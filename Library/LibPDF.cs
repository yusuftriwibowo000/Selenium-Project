using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

using iText.Kernel.Pdf;
using iText.IO.Image;
using iText.Layout;
using iText.Layout.Element;
using iText.Layout.Properties;
using OpenQA.Selenium;
using iText.Kernel.Geom;
using iText.Kernel.Pdf.Canvas;
using iText.Kernel.Font;
using iText.Kernel.Colors;
using iText.IO.Font;
using iText.Layout.Borders;
using SystemPath = System.IO.Path;
using SeleniumNew;

namespace LibraryPDF
{
    public class LibPDF
    {
        static string projectDir;
        public static Document document;
        public static PdfDocument pdf;
        public static Color fontColor;
        public static Color fontColor2;
        public static Color fontColor3;
        // Create a transparent color with an alpha value
        public static Color transparentColor;
        private static PageSize pageSize;
        private static List<string> tableOfContent = new List<string>();
        private static List<string> startEndDate = new List<string>();
        private static float verticalPosition;
        private static float verticalPositionBottom;
        private static float centerX;
        private static float marginPage;
        // RGB Biru
        static int red      = 49;
        static int green    = 132;
        static int blue     = 155;
        // RGB Orange
        static int red2     = 227;
        static int green2   = 108;
        static int blue2    = 10;
        // RGB Biru Muda
        static int red3     = 146;
        static int green3   = 205;
        static int blue3    = 220;

        // Inisialisasi Dokumen PDF
        public static void InitializeDocument()
        {
            string tempProjectDir   = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
            projectDir              = tempProjectDir.Replace("\\", "/"); //Path project
            string reportDir        = SystemPath.Combine(tempProjectDir, "Report");
            string todayReportDir   = SystemPath.Combine(reportDir, DateTime.Now.ToString("yyyyMMdd"));
            string mainDir          = todayReportDir.Replace("\\", "/");

            if(!Directory.Exists(reportDir))
            {
                Directory.CreateDirectory(reportDir);
            }

            if(!Directory.Exists(todayReportDir))
            {
                Directory.CreateDirectory(todayReportDir);
            }
            
            PdfWriter writer    = new PdfWriter(mainDir + "/" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf");
            pdf                 = new PdfDocument(writer);
            document            = new Document(pdf, PageSize.A4); // document pdf
            verticalPosition    = pdf.GetDefaultPageSize().GetHeight() - 50; // get vertical position of page
            //verticalPosition    = pdf.GetDefaultPageSize().GetBottom() + 30; // get vertical position of page
            fontColor           = new DeviceRgb(red, green, blue); // blue
            fontColor2          = new DeviceRgb(red2, green2, blue2); //orange
            fontColor3          = new DeviceRgb(red3, green3, blue3); // biru muda
            transparentColor    = new DeviceCmyk(0,0,0,0); // warna transparan
            centerX             = pdf.GetDefaultPageSize().GetWidth() / 2; // get center X
            marginPage          = 70; // Set koordinat X halaman rata kiri (buat deskripsi gambar)
        }

        // Tambah halaman cover ke dokumen PDF
        public static void CreateCover()
        {
            pageSize        = pdf.GetDefaultPageSize();
            startEndDate.Add(DateTime.Now.ToString("dd MMM yyyy HH:mm:ss")); // start testing date
            float centerY   = (pageSize.GetHeight() / 2) + 90;
            string headerText   = "Automation";
            string headerText2  = "Regression Test";
            string headerText3  = "Project Code";
            string headerText4  = "Automation Test Execution Document";
            string headerText5  = "Prepared By Automation Team";
            string headerText6  = DateTime.Now.ToString("yyyy-MM-dd");
            string headerText7  = "COPYRIGHT NOTICE";
            string headerText8  = "Copyright "+DateTime.Now.ToString("yyyy")+" by Yusuf Tri Wibowo";

            Image logo = new Image(ImageDataFactory
            .Create(projectDir + "/Assets/MainImage.jpg"))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .ScaleToFit(75f, 50f)
            .SetFixedPosition(marginPage,centerY - 330);

            Paragraph header = FormattedParagraph(headerText,fontColor, TextAlignment.RIGHT,34, centerX, centerY, true);
            Paragraph header2 = FormattedParagraph(headerText2, fontColor2, TextAlignment.RIGHT, 22, centerX - 25, centerY - 25, true);
            Paragraph header3 = FormattedParagraph(headerText3, DeviceRgb.BLACK, TextAlignment.RIGHT, 20, centerX - 50, centerY - 50);
            Paragraph header4 = FormattedParagraph(headerText4, DeviceRgb.BLACK, TextAlignment.LEFT, 18, marginPage, centerY - 250);
            Paragraph header5 = FormattedParagraph(headerText5, DeviceRgb.BLACK, TextAlignment.LEFT, 12, marginPage, centerY - 265);
            Paragraph header6 = FormattedParagraph(headerText6, DeviceRgb.BLACK, TextAlignment.LEFT, 11, marginPage, centerY - 280);
            Paragraph header7 = FormattedParagraph(headerText7, new DeviceGray(0.5f), TextAlignment.LEFT, 11, marginPage, centerY - 350);
            Paragraph header8 = FormattedParagraph(headerText8, new DeviceGray(0.5f), TextAlignment.LEFT, 10, marginPage, centerY - 365, true);

            document.Add(header);
            document.Add(header2);
            document.Add(header3);
            document.Add(header4);
            document.Add(header5);
            document.Add(header6);
            document.Add(logo);
            document.Add(header7);
            document.Add(header8);
        }

        // Format Paragraf untuk Header
        private static Paragraph FormattedParagraph(string text, Color fontColor, TextAlignment alignment, float fontSize, float positionX, float positionY, bool bold = false)
        {
            Paragraph paragraph = new Paragraph(text)
                .SetFontColor(fontColor)
                .SetTextAlignment(alignment)
                .SetFontSize(fontSize);

            // Jika format paragraf ingin menggunakan bold
            if(bold)
            {
                paragraph.SetBold();
            }

            paragraph.SetFixedPosition(positionX,positionY,paragraph.GetWidth());
            return paragraph;
        }

        // Method Add Table of Contents
        public static void CreateTableOfContents()
        {
            PdfPage page2       = pdf.GetPage(2);
            PdfCanvas canvas    = new PdfCanvas(page2, true); // set canvas di halaman ke-2
            float xPosition     = marginPage; // Set koordinat X halaman rata kiri (buat deskripsi gambar)
            float xPosition2    = pdf.GetDefaultPageSize().GetRight() - marginPage; // Set koordinat X halaman rata kanan (buat numbers of page)
            float yPosition     = page2.GetPageSize().GetTop() - marginPage; // // Set koordinat Y halaman
            PdfFont boldFont    = PdfFontFactory.CreateFont(FontConstants.HELVETICA_BOLD);

            // Tambah table of contents
            string title = "Table of Contents";
            tableOfContent.Add($"{title}..2");    // Tambah ke list table of contents

            canvas.BeginText().SetFontAndSize(boldFont, 14)
                .SetColor(fontColor, true)
                .MoveText(centerX-50, yPosition)
                .ShowText(title)
                .EndText();

            yPosition -= 30; // Set koordinat Y agar bottom margin 30f

            // Data list "Table of Contents"
            string lastData = tableOfContent.Last();
            string[] splitLastData = lastData.Split(new string[] {".."}, StringSplitOptions.None);

            //Data list "Document Summary"
            int indexDocSum     = tableOfContent.Count() - 2;
            string dataDocSum   = tableOfContent[indexDocSum];
            string[] splitDocSum = dataDocSum.Split(new string[] {".."}, StringSplitOptions.None);

            if (tableOfContent.Any())
            {
                // Masukkan "Table of Contents" ke list Table of Contents
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                    .SetColor(DeviceRgb.BLACK, true)
                    .MoveText(xPosition, yPosition)
                    .ShowText(splitLastData[0])
                    .EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                    .MoveText(xPosition2, yPosition)
                    .ShowText(splitLastData[1])
                    .EndText();
                // Masukkan "Document Summary" ke list Table of Contents
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                    .SetColor(DeviceRgb.BLACK, true)
                    .MoveText(xPosition, yPosition-12)
                    .ShowText(splitDocSum[0])
                    .EndText();
                canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                    .MoveText(xPosition2, yPosition-12)
                    .ShowText(splitDocSum[1])
                    .EndText();
            }
            // Tambah data test step ke list Table of Contents
            foreach (string content in tableOfContent)
            {
                if (content != lastData && content != dataDocSum)
                {
                    if (content == tableOfContent.First())
                    {
                        yPosition = yPosition - 24;
                    }
                    string[] splitContents = content.Split(new string[] {".."}, StringSplitOptions.None);
                    canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                        .MoveText(xPosition, yPosition)
                        .ShowText(splitContents[0])
                        .EndText();
                    canvas.BeginText().SetFontAndSize(PdfFontFactory.CreateFont(), 10)
                        .MoveText(xPosition2, yPosition)
                        .ShowText(splitContents[1])
                        .EndText();

                    yPosition -= 12; // Set bottom margin 12f pada tiap line
                }
            }
            // End of tambah table of contents
            canvas.Release();
        }

        //Method Add Document Summary
        public static void CreateDocumentSummary()
        {
            PdfPage page3 = pdf.GetPage(3);
            float yPosition = page3.GetPageSize().GetTop() - marginPage; // // Set koordinat Y halaman 3

            string titleDocSum = "Document Summary";
            tableOfContent.Add($"{titleDocSum}..3");

            // Set Title Document Summary
            Paragraph titleDocumentSummary = new Paragraph(titleDocSum)
                .SetFontSize(14)
                .SetBold()
                .SetFontColor(fontColor)
                .SetFixedPosition(centerX - 50, yPosition, 500);

            int numRows = tableOfContent.Count - 1; // Jumlah row
            int numColumns = 5; // Jumlah coloumn

            // Buat Header Document Summary
            Table table = new Table(numColumns).SetFixedPosition(marginPage,yPosition-140,475).SetFontSize(8);
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("TC ID").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Scenario Name").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Test Case").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Procedure/Test Step").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Status").SetBold()));

            // Kolom yang tidak di lakukan rowspan
            Cell tcID           = new Cell(numRows, 1).Add(new Paragraph("Test Case"));
            Cell scenarioName   = new Cell(numRows, 1).Add(new Paragraph("Testing"));
            Cell testCase       = new Cell(numRows, 1).Add(new Paragraph("Step"));
            table.AddCell(tcID);
            table.AddCell(scenarioName);
            table.AddCell(testCase);

            // Set kolom Procedure/Test Step dan Status di tambahkan ke tabel sesuai step yang ada tiap TC
            for (int row = 0; row < numRows; row++)
            {
                for (int col = 3; col < numColumns; col++)
                {
                    string data = tableOfContent[row];
                    string[] splitData = data.Split(new string[] { ".." }, StringSplitOptions.None);
                    if (col == 3)
                    {
                        Cell cell = new Cell().Add(new Paragraph(splitData[0]));
                        table.AddCell(cell);
                    }
                    else if(col == 4)
                    {
                        Cell cell = new Cell().Add(new Paragraph(splitData[2]));
                        table.AddCell(cell);
                    } 
                }
            }
            // Tambah element ke dokumen PDF
            new Canvas(page3, pdf.GetDefaultPageSize())
                .Add(titleDocumentSummary)
                .Add(table);
        }

        // Method screenshot step by step
        public static void CaptureScreen(List<string> screenshotPaths, string descImage, string stepStatus)
        {
            IWebDriver driver   = SingletonDriver.GetDriver();
            Screenshot ss       = ((ITakesScreenshot)driver).GetScreenshot();
            String runName      = "ss" + screenshotPaths.Count;
            String fileName     = projectDir + "/Report/" + runName + ".png";
            ss.SaveAsFile(fileName);

            screenshotPaths.Add(fileName);
            descImage = screenshotPaths.Count +". "+ descImage; // Format deskripsi gambar misal : 1. Open Google.com 
            LibPDF.AddScreenshotToPDF(fileName, descImage, stepStatus); // Call 

            if (File.Exists(fileName))
            {
                File.Delete(fileName);
            }
        }
        // Method tambah konten ke PDF
        public static void AddScreenshotToPDF(string imagePath, string descImage, string stepStatus)
        {
            if (tableOfContent.Count == 0)
            {
                pdf.AddNewPage();
                pdf.AddNewPage();
                pdf.AddNewPage();
            }
            if (verticalPosition < 200f) // Jika sisa space halaman < 200f
            {
                pdf.AddNewPage(); // Go to halaman selanjutnya
                verticalPosition = pdf.GetDefaultPageSize().GetHeight() - 50f;
            }

            // Format element deskripsi gambar
            Paragraph description = new Paragraph(descImage)
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(10)
                .SetMarginBottom(5f)
                .SetFixedPosition(marginPage, verticalPosition, 500);

            // Format element gambar
            Image img = new Image(ImageDataFactory
                .Create(imagePath))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .ScaleAbsolute(450f, 210f)
                .SetFixedPosition(marginPage, verticalPosition - 210)
                .SetMarginBottom(20f);

            verticalPosition -= 250f; // Adjust this value based on the space you want between contents

            // Tambah contents ke dokumen PDF
            new Canvas(pdf.GetLastPage(), pdf.GetDefaultPageSize())
                .Add(description)
                .Add(img);

            tableOfContent.Add($"{descImage}..{pdf.GetNumberOfPages()}..{stepStatus}"); // Tambah list ke table of contents
        }

        // Method Page of Numbers
        public static void CreatePageOfNumbers()
        {
            //Page numbers
            int n = pdf.GetNumberOfPages();
            for (int i = 1; i <= n; i++)
            {
                //create a document for each page
                if (i != 1)
                {                 
                    document.ShowTextAligned(
                        new Paragraph().Add("Page " + i + " of " + n).SetItalic()
                        .SetFontSize(9),
                        530,    // koordinat X
                        30,     // koordinat Y
                        i,      // page number
                        TextAlignment.RIGHT,
                        VerticalAlignment.BOTTOM,
                        0       // rotation angle
                    );
                    document.ShowTextAligned(
                        new Paragraph().Add("Automation - Regression Test").SetItalic()
                        .SetFontSize(9),
                        centerX,    // koordinat X
                        30,     // koordinat Y
                        i,      // page number
                        TextAlignment.CENTER,
                        VerticalAlignment.BOTTOM,
                        0       // rotation angle
                    );
                    document.ShowTextAligned(
                        new Paragraph().Add("Confidential").SetItalic()
                        .SetFontSize(9),
                        marginPage, // koordinat X
                        30,     // koordinat Y
                        i,      // page number
                        TextAlignment.LEFT,
                        VerticalAlignment.BOTTOM,
                        0       // rotation angle
                    );
                }
            }
        }

        // Method Header of page
        public static void CreateHeaderOfPage()
        {
            startEndDate.Add(DateTime.Now.ToString("dd MMM yyyy HH:mm:ss")); // End testing date
            verticalPosition = pdf.GetDefaultPageSize().GetHeight() - 50f;
            int n = pdf.GetNumberOfPages();
            for (int i = 1; i <= n; i++)
            {
                //create a document for each page
                if (i != 1)
                {
                    // Format element deskripsi gambar
                    Paragraph titleHeader = new Paragraph()
                        .Add(new Text("Automation Test Execution Document").SetBold())
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetFontSize(11)
                        .SetBold()
                        .SetMarginBottom(5f)
                        .SetFixedPosition(marginPage, verticalPosition, 500);
                    
                    // Create a table
                    Table table = new Table(UnitValue.CreatePercentArray(4))
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetFontSize(8)
                        //.SetBorder(Border.NO_BORDER)
                        .AddCell("Project No.").SetBorder(Border.NO_BORDER)
                        .AddCell("Project Code").SetBorder(Border.NO_BORDER)
                        .AddCell("Tester").SetBorder(Border.NO_BORDER)
                        .AddCell("Automation Team").SetBorder(Border.NO_BORDER)
                        //.SetBorderBottom(new SolidBorder(DeviceRgb.BLACK, 1))
                        //.SetBorderTop(new SolidBorder(DeviceRgb.BLACK, 1))
                        //.SetBorderLeft(new SolidBorder(DeviceRgb.BLACK, 1))
                        //.SetBorderRight(new SolidBorder(DeviceRgb.BLACK, 1))
                        .SetFixedPosition(marginPage, verticalPosition - 15, 500);

                    new Canvas(pdf.GetPage(i), pdf.GetDefaultPageSize())
                        .Add(titleHeader)
                        .Add(table);
                }
            }
        }
       
        // Method Generate PDF
        public static void GeneratePDF()
        {
            //Call method document summary
            CreateDocumentSummary();
            //Call method table of contents
            CreateTableOfContents();
            //// Call method page number
            CreatePageOfNumbers();

            //CreateHeaderOfPage();

            document.Close();
        }
    }
}