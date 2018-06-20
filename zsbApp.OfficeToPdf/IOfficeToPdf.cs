/*
 * 2018年4月27日 19:42:19 郑少宝
 */
namespace zsbApp.OfficeToPdf
{
    public interface IOfficeToPdf
    {
        /// <summary>
        /// 标识
        /// </summary>
        string Tag { get; }

        /// <summary>
        /// office 转成 pdf 文件
        /// </summary>
        /// <param name="officeFileName">office 文件绝对路径</param>
        /// <param name="pdfFileName">输出的 pdf 文件绝对路径</param>
        /// <returns>错误信息，null 代表正常</returns>
        string Convert(string officeFileName, string pdfFileName);
    }
}
