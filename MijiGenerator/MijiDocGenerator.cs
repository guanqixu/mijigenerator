using Spire.Doc;
using Spire.Doc.Documents;
using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Text.RegularExpressions;

namespace MijiGenerator
{
    /// <summary>
    /// word 生成器
    /// </summary>
    class MijiDocGenerator
    {
        public string TemplateFile { get; set; }

        public MijiDocGenerator(string template)
        {
            TemplateFile = template;
        }

        private string _htmlPath;

        public bool Gen(string htmlPath)
        {
            _htmlPath = htmlPath;

            var doc = new Document(TemplateFile);

            Console.WriteLine($"Gen {htmlPath}");
            var parser = new MijiHtmlParser();
            var text = parser.GetInnerHtml(htmlPath, "//span");
            text = text.Replace("<div><br></div>", "\r\n").Replace("<div>", string.Empty).Replace("</div>", "\r\n").Trim();

            //之后出现好多这种 br 标签，统一替换成回车换行符
            text = text.Replace("<br clear=\"none\">", string.Empty)
                .Replace("<span style=\"font-weight: bold;\">", string.Empty)
                .Replace("</span>", string.Empty)
                .Replace("<font style=\"font-size: 10pt;\">", string.Empty)
                .Replace("<br>", string.Empty)
                .Replace("</font>", string.Empty)
                .Replace("<b>", string.Empty)
                .Replace("</b>", string.Empty);
            string imgPattern = "<img src=\\\"(?<imgPath>\\d{8}.+?)\\\".*>";

            List<string> imgPaths = new List<string>();

            foreach (Match m in Regex.Matches(text, imgPattern))
            {
                var path = m.Groups["imgPath"].Value;
                imgPaths.Add(path);
                text = text.Replace(m.Value, string.Empty);
            }

            ChangeBuiltin(doc);

            ImportText(doc, text);
            ImportImg(doc, imgPaths);
            //ImportImgCommend(doc, imgCommends);

            Directory.CreateDirectory("output_doc");
            doc.SaveToFile($".\\output_doc\\{doc.BuiltinDocumentProperties.Title}.docx", FileFormat.Docx2013);
            Console.WriteLine($"End {$".\\output_doc\\{doc.BuiltinDocumentProperties.Title}.docx"}");
            return true;

        }

        private void ChangeBuiltin(Document doc)
        {
            var title = Path.GetFileNameWithoutExtension(_htmlPath);
            var subject = "觅记";
            var author = "觅哥带盐人";
            var comments = CalcBornDay(title);

            doc.BuiltinDocumentProperties.Title = title;
            doc.BuiltinDocumentProperties.Subject = subject;
            doc.BuiltinDocumentProperties.Author = author;
            doc.BuiltinDocumentProperties.Comments = comments;
        }

        private void ImportText(Document doc, string text)
        {
            doc.Replace("内容", text, true, false);
        }

        private void ImportImg(Document doc, IEnumerable<string> imagePaths)
        {
            BookmarksNavigator bn = new BookmarksNavigator(doc);
            bn.MoveToBookmark("插图", true, true);
            Section section = doc.AddSection();

            foreach (var path in imagePaths)
            {
                var graph = section.AddParagraph();
                Image img = Image.FromFile("html/" + path);
                var pic = graph.AppendPicture(img);
                pic.HeightScale = 40f;
                pic.WidthScale = 40f;
                pic.VerticalAlignment = ShapeVerticalAlignment.Center;
                bn.InsertParagraph(graph);
            }

            doc.Sections.Remove(section);
        }

        private void ImportImgCommend(Document doc, List<string> commends)
        {
            string importText = string.Join(",", commends);
            doc.Replace("图片描述", importText, true, false);
        }

        /// <summary>
        /// 计算出生的
        /// </summary>
        /// <param name="title"></param>
        /// <returns></returns>
        private string CalcBornDay(string title)
        {
            var str_date = title.Split(' ')[0];
            DateTime.TryParseExact(str_date, "yyyyMMdd", System.Globalization.CultureInfo.InvariantCulture, System.Globalization.DateTimeStyles.None, out DateTime date);
            var miBirthday = new DateTime(2017, 5, 6);

            int day = date.Day - miBirthday.Day;

            bool noenoughmonth = false;
            if (date.Day < miBirthday.Day)
            {
                var span = date - date.AddMonths(-1);
                day += span.Days;
                noenoughmonth = true;
            }

            int month = (date.Year - miBirthday.Year) * 12 + (date.Month - miBirthday.Month) - (noenoughmonth ? 1 : 0);

            var str_day = day != 0 ? $"{day}天" : string.Empty;

            //小于两岁则只显示月数
            if (month < 24)
            {
                return $"{month}个月{str_day}";
            }
            else
            {
                var year = month / 12;
                month = month % 12;
                var str_month = month != 0 ? $"{month}个月" : string.Empty;
                return $"{year}岁{str_month}{str_day}";
            }
        }
    }
}
