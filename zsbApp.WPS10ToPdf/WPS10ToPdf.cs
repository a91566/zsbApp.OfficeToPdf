/*
 * 2018年4月28日 16:53:19 郑少宝
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;
using Microsoft.Office.Interop.Excel;

namespace zsbApp
{
    public class WPS10ToPdf : zsbApp.OfficeToPdf.BaseOfficeToPdf, zsbApp.OfficeToPdf.IOfficeToPdf
    {
        /// <summary>
        /// 本接口的标识
        /// </summary>
        string zsbApp.OfficeToPdf.IOfficeToPdf.Tag
        {
            get { return "WPS10"; }
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
            try
            {
                Type type = Type.GetTypeFromProgID("KWps.Application");
                if (type == null)
                    return "没有安装 WPS";
                dynamic wps = Activator.CreateInstance(type);
                dynamic document = wps.Documents.Open(this._officeFileName, Visible: false);
                document.ExportAsFixedFormat(this._pdfFileName, WdExportFormat.wdExportFormatPDF);//doc  转pdf 
                document.Close();
                wps.Quit();
            }
            catch (Exception ex)
            {
                result = ex.Message;
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
            try
            {
                Type type = Type.GetTypeFromProgID("KWps.Application");
                if (type == null)
                    return "没有安装 WPS";
                dynamic wps = Activator.CreateInstance(type);
                dynamic document = wps.Workbooks.Open(this._officeFileName, Visible: false);
                document.ExportAsFixedFormat(this._pdfFileName, XlFixedFormatType.xlTypePDF);
                document.Close();
                wps.Quit();
            }
            catch (Exception ex)
            {
                result = ex.Message;
            }
            return result;
        }

        
    }
}
