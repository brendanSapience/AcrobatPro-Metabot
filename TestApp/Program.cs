using AutomationAnywhere;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace TestApp
{
    class Program
    {
        static void Main(string[] args)
        {
            String PDFFilePath = @"C:\Users\brendan.sapience\Google Drive\AutomationAnywhere\Customers\GSK\Aesica\2017IN000001091.pdf";
            BasicIac bi = new BasicIac();
          //  int PageNumber = bi.StartAcrobatIac(PDFFilePath);
          //  Console.Write("Number of Pages:" + PageNumber);

            //String PnForWord = bi.GetPageNumforWord(PDFFilePath, "manufacturing",0,1);
            //Console.Write("Debug:" + PnForWord);

           //   String something = bi.getTextFromPdf(PDFFilePath);
            //  Console.Write("Text:" + something);

            Boolean b = bi.OCRDocumentAndSave(PDFFilePath);
            Console.Write("OCR PDF Done: "+b);

             String Range = bi.getPageRangeBetweenStrings(PDFFilePath,  "CERTIFICATE OF QUALITY","Not more than", true, true);
             Console.Write("Range Found:" + Range);

            Console.ReadKey();
        }
    }
}
