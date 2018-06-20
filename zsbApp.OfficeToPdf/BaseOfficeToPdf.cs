/*
 * 2018年4月27日 19:42:19 郑少宝
 */

using System;
namespace zsbApp.OfficeToPdf
{
    public class BaseOfficeToPdf
    {
        /// <summary>
        /// office 文件绝对路径
        /// </summary>
        protected string _officeFileName;
        /// <summary>
        /// 输出的 pdf 文件绝对路径
        /// </summary>
        protected string _pdfFileName;

        /// <summary>
        /// office 转成 pdf 文件
        /// </summary>
        /// <param name="officeFileName">office 文件绝对路径</param>
        /// <param name="pdfFileName">输出的 pdf 文件绝对路径</param>
        /// <returns>错误信息，null 代表正常</returns>
        protected string DoWork(string officeFileName, string pdfFileName)
        {
            if (!System.IO.File.Exists(officeFileName))
                return "文件不存在";
            this._officeFileName = officeFileName;
            this._pdfFileName = pdfFileName;
            if (officeFileName.ToLower().EndsWith("doc") || officeFileName.ToLower().EndsWith("docx"))
            {
                return WordToPDF();
            }
            else if (officeFileName.ToLower().EndsWith("xls") || officeFileName.ToLower().EndsWith("xlsx"))
            {
                return ExcelToPDF();
            }
            return "未能识别的文件格式";
        }

        /// <summary>
        /// Word 转成 pdf 文件
        /// </summary>
        /// <returns>错误信息，null 代表正常</returns>
        virtual protected string WordToPDF()
        {
            throw new NotImplementedException();
        }

        /// <summary>
        /// Excel 转成 pdf 文件
        /// </summary>
        /// <param name="officeFileName">office 文件绝对路径</param>
        /// <param name="pdfFileName">输出的 pdf 文件绝对路径</param>
        /// <returns>错误信息，null 代表正常</returns>
        virtual protected string ExcelToPDF()
        {
            throw new NotImplementedException();
        }
    }
}
