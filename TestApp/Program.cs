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
            String PDFFilePath = @"C:\Users\brendan.sapience\Google Drive\AutomationAnywhere\Customers\GSK\Aesica\2017IN000000957.pdf";
            BasicIac bi = new BasicIac();
            //int PageNumber = bi.StartAcrobatIac(PDFFilePath);
            //Console.Write("Debug:" + PageNumber);

            //String PnForWord = bi.GetPageNumforWord(PDFFilePath, "manufacturing",0,1);
            //Console.Write("Debug:" + PnForWord);

            //  String something = bi.getTextFromPdf(PDFFilePath);
            //  Console.Write("Debug:" + something);

            //String Range = bi.getPageRangeBetweenStrings(PDFFilePath,  "ELECTRONIC WITHDRAWALS","OTHER WITHDRAWALS, FEES & CHARGES", true, false);
            //Console.Write("Range Found:" + Range);

            bi.OCRDocumentAndSave(PDFFilePath);
            Console.Write("Debug: Done");

            Console.ReadKey();
        }
    }
}
