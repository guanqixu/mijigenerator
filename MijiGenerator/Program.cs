using System;
using System.IO;
using System.Text.RegularExpressions;

/// <summary>
/// MIJI 生成器
/// </summary>
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
                p.Gen();

                input = Console.ReadLine();
            }
            while (input != "");
        }


        public void Gen()
        {
            MijiDocGenerator gen = new MijiDocGenerator(@"..\..\Template.docx");
            Regex regex = new Regex(@"\d{8}");
            foreach (var file in Directory.EnumerateFiles(@".\html", "*.html"))
            {
                //由日期开头的文件
                if (regex.IsMatch(file))
                {
                    gen.Gen(file);
                }
            }
        }
    }
}
