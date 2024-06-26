﻿using System;
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
using LibraryExcel;
using SeleniumNew;
using iText.Layout.Layout;
using iText.Kernel.Pdf.Action;
using iText.Kernel.Pdf.Navigation;
//using DocumentFormat.OpenXml.Spreadsheet;

namespace LibraryPDF
{
    public class LibPDF
    {
        public static Document document;
        public static PdfDocument pdf;
        public static Color fontColor;
        public static Color fontColor2;
        public static Color fontColor3;
        public static Color fontColor4;
        // Create a transparent color with an alpha value
        public static Color transparentColor;
        private static PageSize pageSize;
        // Project Path
        public static string tempProjectDir = Directory.GetParent(Directory.GetCurrentDirectory()).Parent.Parent.FullName;
        public static string projectDir = tempProjectDir.Replace("\\", "/"); //Path project
        public static string globalExcelFilePath = projectDir + "/Excel/Global_Report.xlsx";
        private static List<string> tableOfContent = new List<string>();
        private static List<string> startEndDate = new List<string>();
        private static float verticalPosition;
        private static float centerX;
        private static float marginLeftRight;
        private static float marginTop;
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
        public static void InitializeDocument(string excelPath, string excelSheet)
        {
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
            
            PdfWriter writer    = new PdfWriter(mainDir + "/" + LibExcel.GetDataExcel(excelPath, "TC_ID", excelSheet)+ "_" + DateTime.Now.ToString("yyyyMMdd_HHmmss") + ".pdf");
            pdf                 = new PdfDocument(writer);
            document            = new Document(pdf, PageSize.A4); // document pdf
            marginLeftRight     = 70; // Set koordinat X halaman rata kiri (buat deskripsi gambar)
            marginTop           = 130;
            verticalPosition    = pdf.GetDefaultPageSize().GetHeight() - marginTop; // get vertical position of page
            fontColor           = new DeviceRgb(red, green, blue); // blue
            fontColor2          = new DeviceRgb(red2, green2, blue2); //orange
            fontColor3          = new DeviceRgb(red3, green3, blue3); // biru muda
            fontColor4          = new DeviceRgb(0,128,0); // green
            transparentColor    = new DeviceCmyk(0,0,0,0); // warna transparan
            centerX             = pdf.GetDefaultPageSize().GetWidth() / 2; // get center X
        }

        // Tambah halaman cover ke dokumen PDF
        public static void CreateCover()
        {
            pageSize            = pdf.GetDefaultPageSize();
            startEndDate.Add(DateTime.Now.ToString("dd MMM yyyy HH:mm:ss")); // start testing date
            float centerY       = (pageSize.GetHeight() / 2) + 90;
            string headerText   = LibExcel.GetDataExcel(globalExcelFilePath, "COVER_TITLE", "Global");
            string headerText2  = LibExcel.GetDataExcel(globalExcelFilePath, "COVER_SUBTITLE","Global");
            string headerText3  = LibExcel.GetDataExcel(globalExcelFilePath, "PROJECT_CODE", "Global");
            string headerText4  = LibExcel.GetDataExcel(globalExcelFilePath, "HEADER_DESCRIPTION", "Global");
            string headerText5  = "Prepared By " + LibExcel.GetDataExcel(globalExcelFilePath, "AUTHOR", "Global");
            string headerText6  = DateTime.Now.ToString("yyyy-MM-dd");
            string headerText7  = "COPYRIGHT NOTICE";
            string headerText8  = "Copyright " + DateTime.Now.ToString("yyyy") + " by " + LibExcel.GetDataExcel(globalExcelFilePath, "CREATOR", "Global");

            Image logo = new Image(ImageDataFactory
            .Create(projectDir + "/Assets/MainImage.jpg"))
            .SetTextAlignment(TextAlignment.LEFT)
            .SetHorizontalAlignment(HorizontalAlignment.LEFT)
            .ScaleToFit(75f, 50f)
            .SetFixedPosition(marginLeftRight,centerY - 330);

            Paragraph header = FormattedParagraph(headerText,fontColor, TextAlignment.RIGHT,34, centerX, centerY, true);
            Paragraph header2 = FormattedParagraph(headerText2, fontColor2, TextAlignment.RIGHT, 22, centerX - 25, centerY - 25, true);
            Paragraph header3 = FormattedParagraph(headerText3, DeviceRgb.BLACK, TextAlignment.RIGHT, 20, centerX - 50, centerY - 50);
            Paragraph header4 = FormattedParagraph(headerText4, DeviceRgb.BLACK, TextAlignment.LEFT, 18, marginLeftRight, centerY - 250);
            Paragraph header5 = FormattedParagraph(headerText5, DeviceRgb.BLACK, TextAlignment.LEFT, 12, marginLeftRight, centerY - 265);
            Paragraph header6 = FormattedParagraph(headerText6, DeviceRgb.BLACK, TextAlignment.LEFT, 11, marginLeftRight, centerY - 280);
            Paragraph header7 = FormattedParagraph(headerText7, new DeviceGray(0.5f), TextAlignment.LEFT, 11, marginLeftRight, centerY - 350);
            Paragraph header8 = FormattedParagraph(headerText8, new DeviceGray(0.5f), TextAlignment.LEFT, 10, marginLeftRight, centerY - 365, true);

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
            float xPosition     = marginLeftRight; // Set koordinat X halaman rata kiri (buat deskripsi gambar)
            float xPosition2    = pdf.GetDefaultPageSize().GetRight() - marginLeftRight; // Set koordinat X halaman rata kanan (buat numbers of page)
            float yPosition     = page2.GetPageSize().GetTop() - marginTop; // // Set koordinat Y halaman
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
        public static void CreateDocumentSummary(string excelPath, string excelSheet)
        {
            PdfPage page3 = pdf.GetPage(3);
            float yPosition = page3.GetPageSize().GetTop() - marginTop; // // Set koordinat Y halaman 3

            string titleDocSum = "Document Summary";
            tableOfContent.Add($"{titleDocSum}..3");

            // Set Title Document Summary
            Paragraph titleDocumentSummary = new Paragraph(titleDocSum)
                .SetFontSize(14)
                .SetBold()
                .SetFontColor(fontColor);
                //.SetFixedPosition(centerX - 50, yPosition, 500);

            // Buat Header Tabel Total Document Summary
            Table tableTotal = new Table(4).SetFontSize(8).SetTextAlignment(TextAlignment.CENTER);
            tableTotal.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Total Passed").SetBold()));
            tableTotal.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Total Failed").SetBold()));
            tableTotal.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Total Done").SetBold()));
            tableTotal.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Total").SetBold()));

            int numRows = tableOfContent.Count - 1; // Jumlah row
            int numColumns = 5; // Jumlah coloumn

            // Buat Header Tabel Test Step Document Summary
            Table table = new Table(numColumns).SetFontSize(8);
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("TC ID").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Scenario Name").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Test Case").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Procedure/Test Step").SetBold()));
            table.AddHeaderCell(new Cell().SetBackgroundColor(fontColor3).Add(new Paragraph("Status").SetBold()));

            int jmlFailed = 0;
            for(int i = 0; i < numRows; i++)
            {
                string data = tableOfContent[i];
                string[] splitData = data.Split(new string[] { ".." }, StringSplitOptions.None);
                if (splitData[2] == "Failed")
                {
                    jmlFailed++;
                }
            }
            // Tabel Total Passed, Total Failed, Total Done, Total (Passed)
            Cell totalPassed = new Cell(numRows, 1).Add(new Paragraph((jmlFailed == 0) ? "1" : "0").SetFontColor(fontColor4).SetFontSize(10));
            Cell totalFailed = new Cell(numRows, 1).Add(new Paragraph((jmlFailed > 0) ? "1" : "0").SetFontColor(DeviceRgb.RED).SetFontSize(10));
            Cell totalDone = new Cell(numRows, 1).Add(new Paragraph("0").SetFontSize(10)); // Assuming you need to set this accordingly
            Cell total = new Cell(numRows, 1).Add(new Paragraph("1").SetFontSize(10));
            tableTotal.AddCell(totalPassed);
            tableTotal.AddCell(totalFailed);
            tableTotal.AddCell(totalDone);
            tableTotal.AddCell(total);

            // Kolom yang di lakukan rowspan di Tabel Test Step (Passed or Failed)
            Cell tcID = new Cell(numRows, 1).Add(new Paragraph(LibExcel.GetDataExcel(excelPath, "TC_ID", excelSheet)).SetFontColor((jmlFailed > 0) ? DeviceRgb.RED : fontColor4));
            Cell scenarioName = new Cell(numRows, 1).Add(new Paragraph(LibExcel.GetDataExcel(excelPath, "SCENARIO_NAME", excelSheet)).SetFontColor((jmlFailed > 0) ? DeviceRgb.RED : fontColor4));
            Cell testCase = new Cell(numRows, 1).Add(new Paragraph(LibExcel.GetDataExcel(excelPath, "TEST_CASE", excelSheet)).SetFontColor((jmlFailed > 0) ? DeviceRgb.RED : fontColor4));
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
                        // Jika step Passed
                        if (splitData[2] == "Passed")
                        {
                            Cell cell = new Cell().Add(new Paragraph(splitData[2]).SetFontColor(fontColor4));
                            table.AddCell(cell);
                        }
                        // Jika step Failed
                        else if (splitData[2] == "Failed")
                        {
                            Cell cell = new Cell().Add(new Paragraph(splitData[2]).SetFontColor(DeviceRgb.RED));
                            table.AddCell(cell);
                        }
                        // Jika step Done
                        else
                        {
                            Cell cell = new Cell().Add(new Paragraph(splitData[2]));
                            table.AddCell(cell);
                        }
                    }
                }
            }
            // Tambah element ke dokumen PDF
            new Canvas(page3, pdf.GetDefaultPageSize())
                .Add(titleDocumentSummary.SetFixedPosition(centerX - 50, yPosition, 500));

            new Canvas(page3, pdf.GetDefaultPageSize())
                .Add(tableTotal.SetFixedPosition(marginLeftRight, yPosition - 40, 460));

            // Mengatur layout untuk menghitung tinggi tanpa menambahkan ke dalam dokumen
            LayoutResult layoutResult = table.CreateRendererSubTree().SetParent(new Document(new PdfDocument(new PdfWriter(new System.IO.MemoryStream()))).GetRenderer()).Layout(new LayoutContext(new LayoutArea(1, PageSize.A4)));

            // Mengambil tinggi hasil layout
            float tableHeight = layoutResult.GetOccupiedArea().GetBBox().GetHeight();

            new Canvas(page3, pdf.GetDefaultPageSize())
                .Add(table.SetFixedPosition(marginLeftRight, yPosition - 100 - tableHeight, 460));
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
            if (verticalPosition < 230f) // Jika sisa space halaman < 200f
            {
                pdf.AddNewPage(); // Go to halaman selanjutnya
                verticalPosition = pdf.GetDefaultPageSize().GetHeight() - marginTop;
            }

            // Format element deskripsi gambar
            Paragraph description = new Paragraph(descImage)
                .SetTextAlignment(TextAlignment.LEFT)
                .SetFontSize(10)
                .SetMarginBottom(5f)
                .SetFixedPosition(marginLeftRight, verticalPosition, 500);

            // Format element gambar
            Image img = new Image(ImageDataFactory
                .Create(imagePath))
                .SetTextAlignment(TextAlignment.CENTER)
                .SetHorizontalAlignment(HorizontalAlignment.CENTER)
                .ScaleAbsolute(450f, 210f)
                .SetFixedPosition(marginLeftRight, verticalPosition - 210)
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
                    PdfCanvas lineFooter = new PdfCanvas(pdf.GetPage(i));
                    lineFooter.SetLineWidth(0.5f);
                    lineFooter.MoveTo(marginLeftRight, 40)
                        .LineTo(530,40)
                        .Stroke();

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
                        new Paragraph().Add(LibExcel.GetDataExcel(globalExcelFilePath, "PROJECT_NAME", "Global")).SetItalic()
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
                        marginLeftRight, // koordinat X
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
                        .Add(new Text(LibExcel.GetDataExcel(globalExcelFilePath, "HEADER_DESCRIPTION", "Global")).SetBold())
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetFontSize(11)
                        .SetBold()
                        .SetMarginBottom(5f)
                        .SetFixedPosition(marginLeftRight, verticalPosition, 500);
                    
                    // Create a table
                    Table table = new Table(4)
                        .SetTextAlignment(TextAlignment.LEFT)
                        .SetFontSize(8)
                        .SetFixedPosition(marginLeftRight, verticalPosition - 50, 460);

                    Cell cell1 = new Cell(3, 1).Add(new Paragraph("Project No.")).SetBorderRight(Border.NO_BORDER).SetBorderLeft(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER);
                    Cell cell2 = new Cell(3, 1).Add(new Paragraph(LibExcel.GetDataExcel(globalExcelFilePath, "PROJECT_CODE", "Global"))).SetBorderLeft(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER);
                    Cell cell3 = new Cell(3, 1).Add(new Paragraph("Tester")).SetBorderRight(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER);
                    Cell cell4 = new Cell(3, 1).Add(new Paragraph(LibExcel.GetDataExcel(globalExcelFilePath, "AUTHOR", "Global"))).SetBorderRight(Border.NO_BORDER).SetBorderLeft(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER);
                    Cell cell5 = new Cell(3, 1).Add(new Paragraph("Project Type")).SetBorderRight(Border.NO_BORDER).SetBorderLeft(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER);
                    Cell cell6 = new Cell(3, 1).Add(new Paragraph(LibExcel.GetDataExcel(globalExcelFilePath, "PROJECT_TYPE", "Global"))).SetBorderLeft(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER);
                    Cell cell7 = new Cell(3, 1).Add(new Paragraph("Start Date")).SetBorderRight(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER);
                    Cell cell8 = new Cell(3, 1).Add(new Paragraph(startEndDate[0])).SetBorderRight(Border.NO_BORDER).SetBorderLeft(Border.NO_BORDER).SetBorderBottom(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER);
                    Cell cell9 = new Cell(3, 1).Add(new Paragraph("Short Description")).SetBorderRight(Border.NO_BORDER).SetBorderLeft(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER).SetBorderBottom(new SolidBorder(DeviceRgb.BLACK, 1));
                    Cell cell10 = new Cell(3, 1).Add(new Paragraph(LibExcel.GetDataExcel(globalExcelFilePath, "HEADER_DESCRIPTION", "Global"))).SetBorderLeft(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER).SetBorderBottom(new SolidBorder(DeviceRgb.BLACK, 1));
                    Cell cell11 = new Cell(3, 1).Add(new Paragraph("End Date")).SetBorderRight(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER).SetBorderBottom(new SolidBorder(DeviceRgb.BLACK, 1));
                    Cell cell12 = new Cell(3, 1).Add(new Paragraph(startEndDate[1])).SetBorderRight(Border.NO_BORDER).SetBorderLeft(Border.NO_BORDER).SetBorderTop(Border.NO_BORDER).SetBorderBottom(new SolidBorder(DeviceRgb.BLACK, 1));
                    table.AddCell(cell1);
                    table.AddCell(cell2);
                    table.AddCell(cell3);
                    table.AddCell(cell4);
                    table.AddCell(cell5);
                    table.AddCell(cell6);
                    table.AddCell(cell7);
                    table.AddCell(cell8);
                    table.AddCell(cell9);
                    table.AddCell(cell10);
                    table.AddCell(cell11);
                    table.AddCell(cell12);

                    new Canvas(pdf.GetPage(i), pdf.GetDefaultPageSize())
                        .Add(titleHeader)
                        .Add(table);
                }
            }
        }
       
        // Method Generate PDF
        public static void GeneratePDF(string excelPath, string excelSheet)
        {
            //Call method document summary
            CreateDocumentSummary(excelPath, excelSheet);
            //Call method table of contents
            CreateTableOfContents();
            // Call method page number
            CreatePageOfNumbers();
            // Call method header of pages
            CreateHeaderOfPage();

            document.Close();
        }
    }
}