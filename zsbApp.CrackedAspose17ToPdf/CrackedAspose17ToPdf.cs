/*
 * 2018年4月28日 10:36:20 郑少宝
 */
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace zsbApp
{
    public class CrackedAspose17ToPdf : zsbApp.OfficeToPdf.BaseOfficeToPdf, zsbApp.OfficeToPdf.IOfficeToPdf
    {
        /// <summary>
        /// 本接口的标识
        /// </summary>
        string zsbApp.OfficeToPdf.IOfficeToPdf.Tag
        {
            get { return "CrackedAspose17"; }
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
            Aspose.Words.Document doc = new Aspose.Words.Document(this._officeFileName);
            doc.Save(this._pdfFileName);
            return result;
        }

        /// <summary>
        /// Excel 转成 pdf 文件
        /// </summary>
        /// <returns>错误信息，null 代表正常</returns>
        override protected string ExcelToPDF()
        {
            string result = null;
            Aspose.Cells.Workbook workbook = new Aspose.Cells.Workbook(this._officeFileName);
            workbook.Save(this._pdfFileName);
            return result;
        }
    }
}

