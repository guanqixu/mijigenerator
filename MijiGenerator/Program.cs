using Spire.Doc;
using System;
using System.IO;
using System.Linq;
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
            string tips = @"Please enter the number:
1. gen;
2. merge;";

            while (true)
            {
                Console.WriteLine(tips);
                string input = Console.ReadLine();

                switch (input)
                {
                    case "1":
                        Console.WriteLine("gen doc");
                        p.Gen();
                        break;
                    case "2":
                        Console.WriteLine("merge doc");
                        p.Merge();
                        break;
                    case "":
                        return;
                    default:
                        Console.WriteLine("input error");
                        break;
                }
            }
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

        public void Merge()
        {
            MijiDocMerger.MergePdf("output_doc");
            //MijiDocMerger.MergePdf(Directory.EnumerateFiles("output_pdf", "*.pdf").ToArray(), "merge.pdf");
        }
    }
}
