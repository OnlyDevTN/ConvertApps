using Microsoft.Office.Interop.Excel;
using System;
using System.IO;
using System.Windows.Forms;
using System.Linq;
using System.Data.OleDb;
using System.Data;

namespace ConvertApp
{
    public partial class frmConvert : Form
    {
        public frmConvert()
        {
            InitializeComponent();
        }

        /// <summary>
        /// Convert word document to PDF
        /// </summary>
        /// <param name="Source">Excel document path</param>x
        /// <param name="Target">Converted PDF document path</param>
        /// <history>
        ///   <para>[Anis Zouari] 01/01/2012 Créé</para>
        /// </history>
        public static void excel2PDF(object Source, object Target)
        {
            // Make sure the source document exists.
            if (!System.IO.File.Exists(Source.ToString()))
                throw new Exception("The specified source workbook does not exist.");

            // Create an instance of the Excel ApplicationClass object.          
            Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();

            // Declare a variable to hold the reference to the workbook.
            Workbook excelWorkBook = null;

            object paramMissing = Type.Missing;

            XlFixedFormatType paramExportFormat = XlFixedFormatType.xlTypePDF;
            XlFixedFormatQuality paramExportQuality = XlFixedFormatQuality.xlQualityStandard;
            bool paramOpenAfterPublish = false;
            bool paramIncludeDocProps = true;
            bool paramIgnorePrintAreas = true;
            object paramFromPage = Type.Missing;
            object paramToPage = Type.Missing;

            try
            {
                // Open the source workbook.
                excelWorkBook = excelApplication.Workbooks.Open(Source.ToString(), paramMissing, paramMissing, paramMissing,
                    paramMissing, paramMissing, paramMissing, paramMissing, paramMissing, paramMissing,
                    paramMissing, paramMissing, paramMissing, paramMissing, paramMissing);

                // Save it in the target format.
                if (excelWorkBook != null)
                    excelWorkBook.ExportAsFixedFormat(paramExportFormat, Target.ToString(), paramExportQuality, paramIncludeDocProps,
                        paramIgnorePrintAreas, paramFromPage, paramToPage, paramOpenAfterPublish, paramMissing);
            }
            catch (Exception ex)
            {
                // Respond to the error.
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Close the workbook object.
                if (excelWorkBook != null)
                {
                    excelWorkBook.Close(false, paramMissing, paramMissing);
                    excelWorkBook = null;
                }

                // Close the ApplicationClass object.
                if (excelApplication != null)
                {
                    excelApplication.Quit();
                    excelApplication = null;
                }

                GC.Collect();
                GC.WaitForPendingFinalizers();
                GC.Collect();
                GC.WaitForPendingFinalizers();
            }

        }

        /// <summary>
        /// Convert word document to PDF
        /// </summary>
        /// <param name="Source">Word document path</param>
        /// <param name="Target">Converted PDF document path</param>
        /// <history>
        ///   <para>[Anis Zouari] 01/01/2012 Créé</para>
        /// </history>
        public static void word2PDF(object Source, object Target)
        {
            Microsoft.Office.Interop.Word.Application MSdoc = null;
            //Use for the parameter whose type are not known or say Missing
            object Unknown = Type.Missing;
            //Creating the instance of Word Application           
            if (MSdoc == null) MSdoc = new Microsoft.Office.Interop.Word.Application();

            try
            {
                MSdoc.Visible = false;
                MSdoc.Documents.Open(ref Source, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown,
                     ref Unknown, ref Unknown, ref Unknown, ref Unknown, ref Unknown);
                MSdoc.Application.Visible = false;
                MSdoc.WindowState = Microsoft.Office.Interop.Word.WdWindowState.wdWindowStateMinimize;

                object format = Microsoft.Office.Interop.Word.WdSaveFormat.wdFormatPDF;

                MSdoc.ActiveDocument.SaveAs(ref Target, ref format,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                        ref Unknown, ref Unknown, ref Unknown,
                       ref Unknown, ref Unknown);
            }
            catch (Exception)
            {
                //MessageBox.Show(e.Message);
            }
            finally
            {
                if (MSdoc != null)
                {
                    MSdoc.Documents.Close(ref Unknown, ref Unknown, ref Unknown);
                }
                // for closing the application
                MSdoc.Application.Quit(ref Unknown, ref Unknown, ref Unknown);
            }
        }

        //Convert Excel files to PDF
        private void btnCnvExcel_Click(object sender, EventArgs e)
        {
            string sourceFilePath = "";
            string targetFilePath = "";
            string[] allFiles = Directory.GetFiles(System.IO.Directory.GetCurrentDirectory());
            //Loop on files on convert each of them
            foreach (var file in allFiles)
            {
                sourceFilePath = file;
                targetFilePath = file.Substring(0, file.IndexOf(".")) + ".pdf";
                try
                {
                    if (sourceFilePath.EndsWith(".xls") || sourceFilePath.EndsWith(".xlsx"))
                    {
                        excel2PDF(sourceFilePath, targetFilePath);
                        System.IO.File.Delete(sourceFilePath);
                    }
                }
                catch (Exception) { }
            }
        }

        //Convert Word files to PDF
        private void btnCnvWord_Click(object sender, EventArgs e)
        {
            string sourceFilePath = "";
            string targetFilePath = "";
            string[] allFiles = Directory.GetFiles(System.IO.Directory.GetCurrentDirectory(), "*.*", SearchOption.AllDirectories).Where(s => s.EndsWith(".doc") || s.EndsWith(".docx")).Where(s => (!s.Contains("~$"))).ToArray();
            //Loop on files on convert each of them
            foreach (var file in allFiles)
            {
                sourceFilePath = file;
                targetFilePath = file.Substring(0, file.IndexOf(".")) + ".pdf";
                try
                {
                    if (sourceFilePath.EndsWith(".doc") || sourceFilePath.EndsWith(".docx"))
                    {
                        word2PDF(sourceFilePath, targetFilePath);
                        System.IO.File.Delete(sourceFilePath);
                    }
                }
                catch (Exception) { }
            }
        }
    }
}
