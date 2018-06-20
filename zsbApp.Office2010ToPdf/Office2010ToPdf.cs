/*
 * 2018年4月27日 20:11:24 郑少宝
 */
using System;
using excel = Microsoft.Office.Interop.Excel;
using word = Microsoft.Office.Interop.Word;

namespace zsbApp
{
    public class Office2010ToPdf : zsbApp.OfficeToPdf.BaseOfficeToPdf, zsbApp.OfficeToPdf.IOfficeToPdf
    {
        [System.Runtime.InteropServices.DllImport("user32.dll")]
        private static extern uint GetWindowThreadProcessId(IntPtr hWnd, out uint lpdwProcessId);

        /// <summary>
        /// 本接口的标识
        /// </summary>
        string zsbApp.OfficeToPdf.IOfficeToPdf.Tag 
        {
            get { return "office2010"; }
        }

        /// <summary>
        /// office 转成 pdf 文件
        /// </summary>
        /// <param name="officeFileName">office 文件绝对路径</param>
        /// <param name="pdfFileName">输出的 pdf 文件绝对路径</param>
        /// <returns>错误信息，null 代表正常</returns>
        string zsbApp.OfficeToPdf.IOfficeToPdf.Convert(string officeFileName, string pdfFileName)
        {
            return base.DoWork(officeFileName, pdfFileName);
        }

        /// <summary>
        /// Word 转成 pdf 文件
        /// </summary>
        /// <returns>错误信息，null 代表正常</returns>
        override protected string WordToPDF()
        {
            string result = null;
            var application = new word.Application();
            word.Document document = null;
            try
            {
                application.Visible = false;
                document = application.Documents.Open(base._officeFileName);
                document.ExportAsFixedFormat(base._pdfFileName, word.WdExportFormat.wdExportFormatPDF);
                document.Save();
            }
            catch (Exception e)
            {
                result = e.Message;
            }
            finally
            {
                document.Close();
                application.Quit();
            }
            return result;
        }

        /// <summary>
        /// Excel 转成 pdf 文件
        /// </summary>
        /// <returns>错误信息，null 代表正常</returns>
        override protected string ExcelToPDF()
        {
            string result = null;
            var application = new excel.Application();
            excel.Workbook document = null;
            try
            {
                application.Visible = false;
                document = application.Workbooks.Open(base._officeFileName);
                document.ExportAsFixedFormat(excel.XlFixedFormatType.xlTypePDF, base._pdfFileName);
                document.Save();
            }
            catch (Exception e)
            {
                result = e.Message;
            }
            finally
            {
                document.Close();
                uint processID;
                GetWindowThreadProcessId((IntPtr)application.Application.Hwnd, out processID);
                System.Diagnostics.Process.GetProcessById((int)processID).Kill();
            }
            return result;
        }
    }
}
