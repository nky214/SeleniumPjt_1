using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;
using OpenQA.Selenium;

namespace SeleniumPjt.RunTest
{
    internal class JanuaryPageTest : RunTest
    {
        Excel.Worksheet TestSheet;
        public JanuaryPageTest(Excel.Worksheet TestSheet)
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
                    string cBehavior = NewTestMethod(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckName"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = NewTestMethod(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckPrice"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = NewTestMethod(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckDescription"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = NewTestMethod(SortOrder);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckReleaseMonth"))
                {
                    string SortOrder = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = NewTestMethod(SortOrder);
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

        public string NewTestMethod(string SortOrder)
        {

            return null;
        }
    }
}
