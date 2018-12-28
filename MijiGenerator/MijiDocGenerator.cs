using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MijiGenerator
{
    class MijiDocGenerator
    {
        public string TemplateFile { get; set; }

        public MijiDocGenerator(string template)
        {
            TemplateFile = template;
        }

        public bool Gen(string htmlPath)
        {
            Console.WriteLine($"Gen {htmlPath}");
            var parser = new MijiHtmlParser();
            var text = parser.GetInnerHtml(htmlPath, "//span");
            text = text.Replace("<div><br></div>", "\r\n").Replace("<div>", string.Empty).Replace("</div>", "\r\n").Trim();

            var doc = new Document(TemplateFile);


            var title = Path.GetFileNameWithoutExtension(htmlPath);
            var subject = "觅记";
            var author = "觅哥带盐人";
            var comments = CalcBornDay(title);
            doc.Replace("内容", text, true, false);

            doc.BuiltinDocumentProperties.Title = title;
            doc.BuiltinDocumentProperties.Subject = subject;
            doc.BuiltinDocumentProperties.Author = author;
            doc.BuiltinDocumentProperties.Comments = comments;

            Directory.CreateDirectory("Output");
            doc.SaveToFile($".\\Output\\{title}.docx", FileFormat.Docx2013);
            Console.WriteLine($"End {$".\\Output\\{title}.docx"}");
            return true;

        }

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

            return $"{month}个月{day}天";
        }
    }
}
