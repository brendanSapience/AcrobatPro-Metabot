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
            BasicIac bi = new BasicIac();
            int PageNumber = bi.StartAcrobatIac("C:\\Users\\brendan.sapience\\Google Drive\\AutomationAnywhere\\Customers\\hess\\Document Samples\\0057_486301224 June-2018.pdf");
            Console.Write("Debug:" + PageNumber);

            String PnForWord = bi.GetPageNumforWord("C:\\Users\\brendan.sapience\\Google Drive\\AutomationAnywhere\\Customers\\hess\\Document Samples\\0057_486301224 June-2018.pdf","ELECTRONIC WITHDRAWALS",1,1);
            Console.Write("Debug:" + PnForWord);
            Console.ReadKey();
        }
    }
}
