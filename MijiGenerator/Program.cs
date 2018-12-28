using Spire.Doc;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;

namespace MijiGenerator
{
    class Program
    {
        static void Main(string[] args)
        {
            Program p = new Program();
            string input;
            do
            {
                p.Test();

                input = Console.ReadLine();
            }
            while (input != "");
            //var parser = new MijiHtmlParser();
            //var text = parser.GetInnerHtml(@"F:\觅记 test\20181224 平安夜咯 觅觅继续自己吃.html", "//span");
            //text = text.Replace("<div><br></div>", "\r\n").Replace("<div>", string.Empty).Replace("</div>", "\r\n").Trim();
            //Console.WriteLine(text);

            //Document doc = new Document("Test.docx");
            //doc.Replace("内容", text, true, false);
            //doc.SaveToFile("Test1.docx", FileFormat.Docx2013);

            //DateTime birthday = new DateTime(2017, 2, 28);
            //DateTime now = new DateTime(2018, 3, 2);

            //int day = now.Day - birthday.Day;

            //bool noenoughmonth = false;
            //if (now.Day < birthday.Day)
            //{
            //    var span = now - now.AddMonths(-1);
            //    day += span.Days;
            //    noenoughmonth = true;
            //}

            //int month = (now.Year - birthday.Year) * 12 + (now.Month - birthday.Month) - (noenoughmonth ? 1 : 0);

            //Console.WriteLine($"{month}个月{day}天");



        }


        public void Test()
        {
            MijiDocGenerator gen = new MijiDocGenerator("Test.docx");
            Regex regex = new Regex(@"\d{8}");
            foreach (var file in Directory.EnumerateFiles(@".\html", "*.html"))
            {
                if (regex.IsMatch(file))
                {
                    gen.Gen(file);
                }
            }
        }
    }
}
