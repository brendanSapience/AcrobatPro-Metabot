
using System;
using System.Collections;
using System.ComponentModel;
using System.Collections.Generic;
using System.Data;
using Acrobat;
using System.Reflection;
using System.Threading;
using System.Linq;

namespace AutomationAnywhere
{
    /// <summary>
    /// Summary description for BasicIac.
    /// </summary>
    public class BasicIac
    {

        /// <summary>
        /// Clean up any resources being used.
        /// </summary>
        void Dispose(bool disposing)
        {
            if (disposing)
            {

            }

        }

        /// <summary>
        /// The main entry point for the application.
        /// </summary>

        //Function to open the PDF and get the number of pages
        public int StartAcrobatIac(String szPdfPathConst)
        {
            //variables
            int iNum = 0;

            try
            {
                //IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;
                CAcroApp avApp;


                //set AVApp Project
                avApp = new AcroAppClass();

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //open the PDF
                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                    iNum = pdDoc.GetNumPages();
                }
                else
                {
                    iNum = 0;
                }
            }
            catch (Exception e)
            {
                iNum = 0;
                Console.Write("Error, Could not open File: "+ szPdfPathConst);
                Console.Write("\nMessage: " + e.Message);
            }
            return iNum;
        }

        // test
        public Boolean OCRDocumentAndSave(String szPdfPathConst)
        {
            CAcroAVDoc avDoc;
            CAcroApp avApp;

            avApp = new AcroAppClass();
            avDoc = new AcroAVDocClass();
            avApp.Show();

            //open the PDF
            if (avDoc.Open(szPdfPathConst, ""))
            {
                //set the pdDoc object and get some data
                CAcroPDDoc pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                try
                {
                    avApp.MenuItemExecute("TouchUp:EditDocument");
                    //System.Threading.Thread.Sleep(1000);
                    avApp.MenuItemExecute("Save");
                }catch (Exception e)
                {
                    return false;
                }
                avApp.CloseAllDocs();
                avApp.Exit();
                
                return true;
                
            }
            avApp.Exit();
            return false;

        }

        // Function to Check if Word is present or not, returns true or false
        public bool IsWordPresent(string szPdfPathConst, string searchword)
        {
            //variables
            bool TextCheck;

            try
            {
                //IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;
                CAcroApp avApp;

                //set AVApp Project
                avApp = new AcroAppClass();

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //Show Acrobat
                avApp.Show();

                //open the PDF if it isn't already open

                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                    //Checking if word is present or not
                    TextCheck = avDoc.FindText(searchword, 0, 0, 0);
                }
                else
                {
                    TextCheck = false;
                }
            }

            catch (Exception)
            {
                TextCheck = false;
            }
            return TextCheck;
        }

        public String getPageRangeBetweenStrings(String szPdfPathConst, String HeaderStr, String FooterStr, Boolean includeFirstPageInRange, Boolean includeLastPageInRange)
        {
            CAcroApp avApp;
            CAcroAVDoc avDoc;
            CAcroAVPageView avPage;
            avApp = new AcroAppClass();
            avDoc = new AcroAVDocClass();
            avDoc.Open(szPdfPathConst, "");
            CAcroPDDoc pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
            List<int> PageListHeader = new List<int>();
            List<int> PageListFooter = new List<int>();
            //AcroPDDoc pdDoc = getPDDoc(szPdfPathConst);
            int TotalNumberOfPages = pdDoc.GetNumPages();
            AcroPDPage page;
            //set AVPage View object
            avPage = (CAcroAVPageView)avDoc.GetAVPageView();
            avApp.Show();
            for (int i = 0; i < TotalNumberOfPages; i++)
            {
                page = (AcroPDPage)pdDoc.AcquirePage(i);
                Boolean TextCheck = avDoc.FindText(HeaderStr,1,1,0);
                if (TextCheck == true)
                {
                    int PageNum = avPage.GetPageNum();
                    PageListHeader.Add(PageNum);
                }
            }
          
            
            List<int> PagesWithHeaderWords = DeDuplicateArray(PageListHeader);

            

            for (int i = 0; i < TotalNumberOfPages; i++)
            {
                page = (AcroPDPage)pdDoc.AcquirePage(i);
                Boolean TextCheck = avDoc.FindText(FooterStr,1,1,0);
                if (TextCheck == true)
                {
                    int PageNum = avPage.GetPageNum();
                    PageListFooter.Add(PageNum);
                }
            }
            List<int> PagesWithFooterWords = DeDuplicateArray(PageListFooter);
            int MinimumFooterRange = 0;
            int MinimumHeaderRange = 0;
            if (PagesWithFooterWords.Count == 0 || PagesWithHeaderWords.Count == 0)
            {
                return "No Range Found";
            }

                MinimumFooterRange = PagesWithFooterWords.Min();
                MinimumHeaderRange = PagesWithFooterWords.Min();
            


            int HeaderFinalPageNumber = MinimumHeaderRange + 1;
            int FooterFinalPageNumber = MinimumFooterRange + 1;

            if (!includeFirstPageInRange)
            {
                HeaderFinalPageNumber++;
            }
            if (!includeLastPageInRange)
            {
                FooterFinalPageNumber--;
            }

            return HeaderFinalPageNumber+"-"+ FooterFinalPageNumber;
        }

            // Returns the list of page numbers on which the word or words can be found (separated by commas, ex: 7,8,9,10)
            // bCaseSensitive: 0 = false, 1 = true, bWholeWordsOnly: 0 = false, 1 = true
            public string GetPageNumforWord(string szPdfPathConst, string searchword, int bCaseSensitive, int bWholeWordsOnly)
        {
            //Initializing variables
            int iNum = 0;
            bool TextCheck;
            int PageNum;
            bool GoToStatus;
            string PageNumConsol = "";
            int ScanPage;
            List<int> PageList = new List<int>();
            List<string> PageListString = new List<string>();

            try
            {
                //Declaring relevant IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;
                CAcroApp avApp;
                CAcroAVPageView avPage;

                //set AVApp Project
                avApp = new AcroAppClass();

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //open the PDF if it isn't already open

                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();

                    //Getting Total Number of Pages in the PDF
                    iNum = pdDoc.GetNumPages();

                    //set AVPage View object
                    avPage = (CAcroAVPageView)avDoc.GetAVPageView();

                    //Navigating to Page 1 to initiate search
                    ScanPage = 0;
                    GoToStatus = avPage.GoTo(ScanPage);

                    //Checking if word is present or not
                    TextCheck = avDoc.FindText(searchword, bCaseSensitive, bWholeWordsOnly, 0);

                    //Declaring variable for storing the previous page number 
                    int PageNumPrev = 0;

                    if (TextCheck == true)
                    {
                        PageNum = avPage.GetPageNum();
                        //First Page is 0 and thus offset is being taken care of
                        PageNum = PageNum + 1;
                        PageList.Add(PageNum);

                        //Incrementing Page numbers and searching for more instances
                        while (TextCheck == true)
                        {
                            //Going to the page next to the previous search result - Not incremented by 1 since PageNum was already incremented for recording.
                            ScanPage = PageNum;
                            if (ScanPage == iNum)
                            {
                                TextCheck = false;
                                break;
                            }
                            GoToStatus = avPage.GoTo(ScanPage);
                            TextCheck = avDoc.FindText(searchword, bCaseSensitive, bWholeWordsOnly, 0);
                            PageNum = avPage.GetPageNum();

                            //Exit loop in case the previous page number is bigger than the current
                            if (PageNumPrev > PageNum)
                            {
                                break;
                            }
                            //Assigning the page number for this search iteration to a previous variable
                            PageNumPrev = PageNum;

                            //First Page is 0 and thus offset is being taken care of
                            PageNum = PageNum + 1;
                            PageList.Add(PageNum);

                        }
                    }
                    else
                    {
                        PageNum = 0;
                        PageList.Add(PageNum);
                    }
                }
                else
                {
                    PageNum = 0;
                    PageList.Add(PageNum);
                }

                //Removing Duplicates in the list due to multiple occurences of word on the same page
                List<int> PageListFilter = new List<int>();
                foreach (int i in PageList)
                {
                    if (!PageListFilter.Contains(i))
                    {
                        PageListFilter.Add(i);
                    }
                }

                //Converting Integer List for Page List to String List
                PageListString = PageListFilter.ConvertAll<string>(delegate (int i)
                {
                    return i.ToString();
                });

                //Converting String List to Comma Delimited List
                PageNumConsol = string.Join(",", PageListString.ToArray());
            }
            catch(Exception)
            {
                PageNumConsol = "Unknown Exception";
            }

            return PageNumConsol;
        }

        private List<int> DeDuplicateArray(List<int> list)
        {
            //Removing Duplicates in the list due to multiple occurences of word on the same page
            List<int> PageListFilter = new List<int>();
            foreach (int i in list)
            {
                if (!PageListFilter.Contains(i))
                {
                    PageListFilter.Add(i);
                }
            }
            return PageListFilter;
        }

        // Saving PDF file
        public bool SavePDF(string szPdfPathConst, string sFullPath)
        {
            //Declaring Variables
            bool SaveAs;

            try
            {
                //IAC objects
                CAcroPDDoc pdDoc;
                CAcroAVDoc avDoc;

                //set AVDoc object
                avDoc = new AcroAVDocClass();

                //open the PDF
                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //set the pdDoc object and get some data
                    pdDoc = (CAcroPDDoc)avDoc.GetPDDoc();
                    SaveAs = pdDoc.Save(1, sFullPath);
                }
                else
                {
                    SaveAs = false;
                }
            }
            catch
            {
                SaveAs = false;
            }

            //Returning output var
            return SaveAs;

        }

        // Closing PDF File
        public bool ClosePDFNoChanges(string szPdfPathConst)
        {
            //Initializing Variables
            bool CloseCheck;

            try
            {
                //SettingObject
                CAcroApp avApp;
                CAcroAVDoc avDoc;

                //set AVApp Project
                avApp = new AcroAppClass();
                //set AVDoc object
                avDoc = new AcroAVDocClass();

                if (avDoc.Open(szPdfPathConst, ""))
                {
                    //Checking if word is present or not
                    CloseCheck = avDoc.Close(1);
                }
                else
                {
                    CloseCheck = false;
                }

                avApp.CloseAllDocs();
                avApp.Exit();
            }
            catch
            {
                CloseCheck = false;
            }

            return CloseCheck;

        }



        public string getTextFromPdf(String szPdfPathConst)
        {
            AcroPDDoc pddoc = getPDDoc(szPdfPathConst);
            String myText = GetTextInPdf(pddoc);
            return myText;
        }

        private AcroPDDoc getPDDoc(String szPdfPathConst)
        {
            //Declaring relevant IAC objects
            AcroPDDoc pdDoc;
            CAcroAVDoc avDoc;
            CAcroApp avApp;
            avApp = new AcroAppClass();
            avDoc = new AcroAVDocClass();
            if (avDoc.Open(szPdfPathConst, ""))
            {
                pdDoc = (AcroPDDoc)avDoc.GetPDDoc();
                return pdDoc;
            }
            return null;
        }

        private string GetTextInPdf(AcroPDDoc pdDoc)
        {
            AcroPDPage page;
            int TotalNumberOfPages = pdDoc.GetNumPages();
            string pageText = "";
            for (int i = 0; i < TotalNumberOfPages; i++)
            {
                page = (AcroPDPage)pdDoc.AcquirePage(i);
                object jso, jsNumWords, jsWord;
                List<string> words = new List<string>();
                try
                {
                    jso = pdDoc.GetJSObject();
                    if (jso != null)
                    {
                        object[] args = new object[] { i };
                        jsNumWords = jso.GetType().InvokeMember("getPageNumWords", BindingFlags.InvokeMethod, null, jso, args, null);
                        int numWords = Int32.Parse(jsNumWords.ToString());
                        for (int j = 0; j <= numWords; j++)
                        {
                            object[] argsj = new object[] { i, j, false };
                            jsWord = jso.GetType().InvokeMember("getPageNthWord", BindingFlags.InvokeMethod, null, jso, argsj, null);
                            words.Add((string)jsWord);
                        }
                    }
                    foreach (string word in words)
                    {
                        pageText += word;
                    }
                }
                catch
                {
                }
            }
            return pageText;
        }
    }


}
