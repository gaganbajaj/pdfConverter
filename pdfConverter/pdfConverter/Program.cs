using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.IO;
using iTextSharp.text;
using iTextSharp.text.pdf;
using System.Collections;
using msExcel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.Word;
//using Microsoft.Office.Interop.PowerPoint;
//using Microsoft.Office.Core;

namespace pdfConverter
{
    class Program
    {
        public static object missing = System.Reflection.Missing.Value;
        private static string sourcefolder;
        private static string destinationfile;
        private static IList fileList = new ArrayList();
        public string SourceFolder
        {
            get { return sourcefolder; }
            set { sourcefolder = value; }
        }
        public string DestinationFile
        {
            get { return destinationfile; }
            set { destinationfile = value; }
        }
        static void Main(string[] args)
        {
            string[] array1 = Directory.GetFiles("D:\\converter");
            for (int i = 0; i < array1.Length; i++)
            {
                ConvertExcelToPdf(array1[i], array1[i].Substring(0, array1[i].LastIndexOf("\\")));
            }

        }

        public static void ConvertExcelToPdf(string excelFileIn, string pdfFileOut)
        {

            msExcel.Application excel = new msExcel.Application();
            try
            {
                excel.Visible = false;
                excel.ScreenUpdating = false;
                excel.DisplayAlerts = false;

                FileInfo excelFile = new FileInfo(excelFileIn);

                string filename = excelFile.FullName;

                msExcel.Workbook wbk = excel.Workbooks.Open(filename, missing,
                missing, missing, missing, missing, missing,
                missing, missing, missing, missing, missing,
                missing, missing, missing);
                wbk.Activate();

                object outputFileName = wbk.FullName.Replace(".xslx", ".pdf"); ;
                msExcel.XlFixedFormatType fileFormat = msExcel.XlFixedFormatType.xlTypePDF;

                // Save document into PDF Format
                wbk.ExportAsFixedFormat(fileFormat, outputFileName,
                missing, missing, missing,
                missing, missing, missing,
                missing);

                object saveChanges = msExcel.XlSaveAction.xlDoNotSaveChanges;
                ((msExcel._Workbook)wbk).Close(saveChanges, missing, missing);
                wbk = null;
            }
            catch (Exception e)
            {
                Console.WriteLine(e);
            }
            finally
            {
                ((msExcel._Application)excel).Quit();
                excel = null;
            }


        }

        public static void AddFile(string pathnname)
        {
            fileList.Add(pathnname);

        }
        public static void Execute()
        {
            MergeDocs();
        }


        public static void MergeDocs()
        {
            //Step 1: Create a Docuement-Object
            iTextSharp.text.Document document = new iTextSharp.text.Document();
            try
            {
                //Step 2: we create a writer that listens to the document
                PdfWriter writer = PdfWriter.GetInstance(document, new FileStream("D:\\Test1\\Final.pdf", FileMode.Create));
                //Step 3: Open the document
                document.Open();
                PdfContentByte cb = writer.DirectContent; PdfImportedPage page;
                int n = 0;
                int rotation = 0;
                string[] array2 = Directory.GetFiles("D:\\Test");
                for (int i = 0; i < array2.Length; i++)
                {
                    if (array2[i].Contains(".pdf"))
                    {
                        AddFile(array2[i]);
                    }
                }

                //Loops for each file that has been listed
                foreach (string filename in fileList)
                {

                    //The current file path
                    string filePath = sourcefolder + filename;
                    // we create a reader for the document
                    PdfReader reader = new PdfReader(filePath);
                    //Gets the number of pages to process
                    n = reader.NumberOfPages; int i = 0;
                    while (i < n)
                    {
                        i++; document.SetPageSize(reader.GetPageSizeWithRotation(1));
                        document.NewPage();
                        //Insert to Destination on the first page
                        if (i == 1)
                        {
                            Chunk fileRef = new Chunk(" ");
                            fileRef.SetLocalDestination(filename); document.Add(fileRef);
                        }
                        page = writer.GetImportedPage(reader, i);
                        rotation = reader.GetPageRotation(i);
                        if (rotation == 90 || rotation == 270)
                        { cb.AddTemplate(page, 0, -1f, 1f, 0, 0, reader.GetPageSizeWithRotation(i).Height); }
                        else
                        {
                            cb.AddTemplate(page, 1f, 0, 0, 1f, 0, 0);
                        }
                    }

                }
            }
            catch (Exception e) { throw e; }
            finally { document.Close(); }
        }


        #region Converting word to PDF

        // C# doesn’t have optional arguments so we’ll need a dummy value
        // object oMissing = System.Reflection.Missing.Value;
        public static void ConvertWordToPdf(String FilePath, String FileName)
        {
            Microsoft.Office.Interop.Word.Application word = new Microsoft.Office.Interop.Word.Application();
            try
            {

                // Get a Word file
                FileInfo wordFile = new FileInfo((FileName));

                word.Visible = false;
                word.ScreenUpdating = false;

                // Cast as Object for word Open method
                Object filename = (Object)FilePath;
                //object filename = (Object)wordFile.FullName;

                // Use the dummy value as a placeholder for optional arguments
                Microsoft.Office.Interop.Word.Document doc = word.Documents.Open(ref filename, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);
                doc.Activate();

                object outputFileName = FilePath.Replace(".docx", ".pdf");
                object fileFormat = WdSaveFormat.wdFormatPDF;

                // Save document into PDF Format
                doc.SaveAs(ref outputFileName,
                ref fileFormat, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing,
                ref missing, ref missing, ref missing, ref missing);

                // Close the Word document, but leave the Word application open.
                // doc has to be cast to type _Document so that it will find the
                // correct Close method.
                object saveChanges = WdSaveOptions.wdDoNotSaveChanges;
                ((_Document)doc).Close(ref saveChanges, ref missing, ref missing);
                doc = null;

                // word has to be cast to type _Application so that it will find
                // the correct Quit method.
                ((Microsoft.Office.Interop.Word._Application)word).Quit(ref missing, ref missing, ref missing);
                word = null;
            }
            catch (Exception e)
            {
            }

            #endregion
        }

        //public static void ConvertPPTXToPDF(string FileName)
        //{
        //    Microsoft.Office.Interop.PowerPoint.Application app = new Microsoft.Office.Interop.PowerPoint.Application();
        //    string sourcePptx = FileName;
        //    string targetPpt = sourcePptx.Replace(“.pptx”, “.pdf”);
        //    object missing = Type.Missing;
        //    Microsoft.Office.Interop.PowerPoint.Presentation pptx = app.Presentations.Open(sourcePptx, MsoTriState.msoFalse, MsoTriState.msoTrue, MsoTriState.msoFalse);
        //    pptx.SaveAs(targetPpt, Microsoft.Office.Interop.PowerPoint.PpSaveAsFileType.ppSaveAsPDF, MsoTriState.msoTrue);
        //    app.Quit();
        //}
    }

 }
