using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using iText.Kernel.Pdf;

namespace MergePDFs
{
    internal class Program
    {
        public static void MergePDFs(List<string> inputFiles, string outputFile)
        {
            // Create a PdfDocument to write the merged output
            using (PdfWriter writer = new PdfWriter(outputFile))
            using (PdfDocument pdf = new PdfDocument(writer))
            {
                PdfDocument pdfDoc;

                foreach (var file in inputFiles)
                {
                    // Open each PDF file for reading
                    pdfDoc = new PdfDocument(new PdfReader(file));
                    pdfDoc.CopyPagesTo(1, pdfDoc.GetNumberOfPages(), pdf);  // Copy all pages to the output document
                    pdfDoc.Close();  // Close the current file
                }
            }

            Console.WriteLine("PDFs merged successfully!");
        }

        public static void Main(string[] args)
        {
            if (args.Length != 2)
            {
                Console.WriteLine("Usage: MergePDFs <input_folder> <output_file>");
                return;
            }

            string inputFolder = args[0];  // Input folder path
            string outputFile = args[1];   // Output file path

            if (!Directory.Exists(inputFolder))
            {
                Console.WriteLine("Error: Input folder does not exist.");
                return;
            }

            // Get all PDFs in the folder and sort by modification timestamp (newest first)
            List<string> pdfFiles = Directory.GetFiles(inputFolder, "*.pdf")
                                             .OrderByDescending(f => File.GetLastWriteTime(f)) // Sort by modification timestamp
                                             .ToList();

            if (pdfFiles.Count == 0)
            {
                Console.WriteLine("No PDF files found in the folder.");
                return;
            }

            // Merge the PDFs
            MergePDFs(pdfFiles, outputFile);
        }
    }
}
