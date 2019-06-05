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
        /// <summary>
        /// 主函数
        /// 
        /// 1. html 转换成 docx
        /// 2. docx 转换成 pdf
        /// 3. 合并 pdf
        /// </summary>
        /// <param name="args"></param>
        static void Main(string[] args)
        {
            Program p = new Program();
            string tips = @"Please enter the number:
1. gen;
2. convert;
3. merge;";

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
                        Console.WriteLine("Please enter docx dir path which to convert:");
                        var path = Console.ReadLine();
                        p.Convert(path);
                        break;
                    case "3":
                        Console.WriteLine("Please enter pdf dir path which to merge:");
                        path = Console.ReadLine();
                        p.Merge(path);
                        break;
                    case "":
                        return;
                    default:
                        Console.WriteLine("input error");
                        break;
                }
            }
        }

        /// <summary>
        /// 生成
        /// </summary>
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

        /// <summary>
        /// 组合
        /// </summary>
        /// <param name="inputDir"></param>
        public void Merge(string inputDir)
        {
            inputDir = inputDir.Trim('"');
            var inputFiles = Directory.GetFiles(inputDir, "20*.pdf");
            var merge = new PdfFileMerge(inputFiles);
            var outputPath = Path.Combine(inputDir, $"merge_{DateTime.Now.ToString("yyyyMMdd_hhmmss")}.pdf");
            merge.Merge(outputPath);
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <param name="inputDir"></param>
        public void Convert(string inputDir)
        {
            inputDir = inputDir.Trim('"');
            var inputFiles = Directory.GetFiles(inputDir, "*.docx");
            var outputDir = inputDir + "_pdf";

            var converter = new PdfConvert(inputFiles);
            converter.Convert(outputDir);
        }

    }
}
