using iTextSharp.text;
using iTextSharp.text.pdf;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;

namespace MijiGenerator
{
    class PdfFileMerge
    {
        private string[] _files;

        /// <summary>
        /// 构造函数
        /// </summary>
        /// <param name="files"></param>
        public PdfFileMerge(string[] files)
        {
            _files = files;
        }

        /// <summary>
        /// 组合
        /// </summary>
        /// <param name="outputPath"></param>
        /// <returns></returns>
        public bool Merge(string outputPath)
        {
            List<PdfReader> readers = new List<PdfReader>();

            PdfWriter pw;
            Document document;

            PdfContentByte pcb;
            PdfImportedPage pip;
            document = new Document(PageSize.A4);
            var fs = new FileStream(outputPath, FileMode.Create, FileAccess.Write);
            pw = PdfWriter.GetInstance(document, fs);
            document.Open();
            pcb = pw.DirectContent;

            foreach (var fileName in _files)
            {
                PdfReader pr = new PdfReader(fileName);
                readers.Add(pr);

                for (int i = 1; i <= pr.NumberOfPages; i++)
                {
                    document.SetPageSize(pr.GetPageSizeWithRotation(i));
                    document.NewPage();
                    pip = pw.GetImportedPage(pr, i);
                    pcb.AddTemplate(pip, 0, 0);
                }

                //奇数页增加空白页
                if (pr.NumberOfPages % 2 != 0)
                {
                    document.NewPage();
                    document.Add(new Chunk());
                }
            }
            document.CloseDocument();

            foreach (var pr in readers)
            {
                pr.Close();
                pr.Dispose();
            }

            readers.Clear();

            //document.Close();
            Console.WriteLine($"Merged {outputPath}");
            return true;
        }
    }
}
