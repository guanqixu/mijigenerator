using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Spire.Doc;
using Spire.Pdf;

namespace MijiGenerator
{
    class MijiDocMerger
    {
        public static bool MergeDoc(string[] docFiles, string outputFile)
        {
            Document doc = new Document();

            foreach (var file in docFiles)
            {
                Console.WriteLine(file);
                Document temp = new Document(file);
                foreach (Section sec in temp.Sections)
                {
                    doc.Sections.Add(sec.Clone());
                }

            }

            //doc.SaveToFile(outputFile, FileFormat.Docx);
            return true;
        }

        public static bool MergePdf(string[] pdfFiles, string outputFile)
        {
            var pdf = PdfDocument.MergeFiles(pdfFiles);
            pdf.Save(outputFile);
            return true;
        }

        public static bool MergePdf(string folder)
        {
            foreach (var item in Directory.EnumerateFiles(folder, "*.docx"))
            {
                Document doc = new Document(item);
                if (doc.PageCount % 2 == 0)
                {
                    doc.SaveToFile($"output_pdf\\{doc.BuiltinDocumentProperties.Title}.pdf", Spire.Doc.FileFormat.PDF);
                    Console.WriteLine($"Convert {$"output_pdf\\{doc.BuiltinDocumentProperties.Title}.pdf"}");
                }
            }

            var pdf = PdfDocument.MergeFiles(Directory.EnumerateFiles("output_pdf", "*.pdf").ToArray());
            pdf.Save("output_pdf\\merge.pdf");

            return true;

        }

    }
}
