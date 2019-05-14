using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using Microsoft.Office.Interop.Word;

namespace MijiGenerator
{
    class PdfConvert
    {
        /// <summary>
        /// 文件集合
        /// </summary>
        private string[] _files;

        /// <summary>
        /// Word 运行程序
        /// </summary>
        private Application _app;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="files"></param>
        public PdfConvert(string[] files)
        {
            _files = files;
        }

        /// <summary>
        /// 转换
        /// </summary>
        /// <param name="outDir"></param>
        public void Convert(string outDir)
        {
            Directory.CreateDirectory(outDir);
            _app = new Application();
            try
            {
                foreach (var file in _files)
                {
                    var input = file;
                    var fileName = Path.GetFileNameWithoutExtension(input);
                    fileName += ".pdf";
                    var output = Path.Combine(outDir, fileName);
                    Convert(input, output, WdSaveFormat.wdFormatPDF);
                    Console.WriteLine($"Converted {output}");
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                _app.Quit();
            }
        }

        /// <summary>
        /// word 转换为 pdf
        /// </summary>
        /// <param name="input">输入</param>
        /// <param name="output">输出</param>
        /// <param name="format">转换后的格式</param>
        private void Convert(string input, string output, WdSaveFormat format)
        {
            object oInput = input;
            object oOutput = output;
            object oFormat = format;
            var doc = _app.Documents.Open(ref oInput);
            doc.Activate();
            doc.SaveAs2(ref oOutput, ref oFormat);
            doc.Close();
        }
    }
}
