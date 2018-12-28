using HtmlAgilityPack;
using System.Text;

namespace MijiGenerator
{
    /// <summary>
    /// html 转换器
    /// 印象笔记
    /// </summary>
    class MijiHtmlParser
    {
        /// <summary>
        /// 获取内部字符串
        /// </summary>
        /// <param name="htmlPath"></param>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public string GetInnerText(string htmlPath, string xpath)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlPath, Encoding.UTF8);

            var node = doc.DocumentNode.SelectSingleNode(xpath);
            var innerText = node.InnerText;

            return innerText;
        }

        /// <summary>
        /// 获取内部 html
        /// </summary>
        /// <param name="htmlPath"></param>
        /// <param name="xpath"></param>
        /// <returns></returns>
        public string GetInnerHtml(string htmlPath, string xpath)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlPath, Encoding.UTF8);

            var node = doc.DocumentNode.SelectSingleNode(xpath);

            var innerText = node.InnerHtml;

            return innerText;
        }

    }
}
