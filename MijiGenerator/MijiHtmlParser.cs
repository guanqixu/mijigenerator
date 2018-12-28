using HtmlAgilityPack;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MijiGenerator
{
    class MijiHtmlParser
    {

        public string GetInnerText(string htmlPath, string xpath)
        {
            var doc = new HtmlDocument();
            doc.Load(htmlPath, Encoding.UTF8);

            var node = doc.DocumentNode.SelectSingleNode(xpath);
            var innerText = node.InnerText;

            return innerText;
        }

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
