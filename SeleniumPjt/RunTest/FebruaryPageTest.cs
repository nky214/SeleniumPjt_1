using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumPjt.RunTest
{
    internal class FebruaryPageTest : RunTest
    {
        Excel.Worksheet TestSheet;
        public FebruaryPageTest(Excel.Worksheet TestSheet)
        {
            this.TestSheet = TestSheet;
        }

        public void ExecuteTest()
        {
            int exeCount = 2;
            while (true)
            {
                string testCase = (string)TestSheet.Cells[exeCount, 1].Value;

                if (testCase.Equals("CheckBrand"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckBrand(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckName"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckName(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckPrice"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckPrice(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckDescription"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckDescription(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckReleaseMonth"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckReleaseMonth(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }

                if ((string)TestSheet.Cells[exeCount + 1, 1].Value == null)
                {
                    break;
                }
                exeCount++;
            }
        }



        public string CheckBrand(string SortOrder)
        {
            string returnText = ""; 
            int s = Int32.Parse(SortOrder);
            IReadOnlyCollection<IWebElement> brandElements = sUtil.FindElements(po.GetShoeBrand());

            for(int i = 0; i < brandElements.Count; i++)
            {
                if(s == i)
                {
                    returnText = brandElements.ElementAt(i).Text;   
                }
            }
            return returnText;
        }
        public string CheckName(string SortOrder)
        {
            string returnText = "";
            int s = Int32.Parse(SortOrder);
            IReadOnlyCollection<IWebElement> nameElements = sUtil.FindElements(po.GetShoeName());

            for (int i = 0; i < nameElements.Count; i++)
            {
                if (s == i)
                {
                    returnText = nameElements.ElementAt(i).Text;
                }
            }
            return returnText;
        }
        public string CheckPrice(string SortOrder)
        {
            string returnText = "";
            int s = Int32.Parse(SortOrder);
            IReadOnlyCollection<IWebElement> priceElements = sUtil.FindElements(po.GetShoePrice());

            for (int i = 0; i < priceElements.Count; i++)
            {
                if (s == i)
                {
                    returnText = priceElements.ElementAt(i).Text;
                }
            }
            return returnText;
        }
        public string CheckDescription(string SortOrder)
        {
            string returnText = "";
            int s = Int32.Parse(SortOrder);
            IReadOnlyCollection<IWebElement> descElements = sUtil.FindElements(po.GetShoeDescription());

            for (int i = 0; i < descElements.Count; i++)
            {
                if (s == i)
                {
                    returnText = descElements.ElementAt(i).Text;
                }
            }
            return returnText;
        }
        public string CheckReleaseMonth(string SortOrder)
        {
            string returnText = "";
            int s = Int32.Parse(SortOrder);
            IReadOnlyCollection<IWebElement> releaseElements = sUtil.FindElements(po.GetShoeReleaseMonth());

            for (int i = 0; i < releaseElements.Count; i++)
            {
                if (s == i)
                {
                    returnText = releaseElements.ElementAt(i).Text;
                }
            }
            return returnText;
        }
    }
}
