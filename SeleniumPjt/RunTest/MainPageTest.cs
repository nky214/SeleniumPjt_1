using OpenQA.Selenium;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumPjt.RunTest
{
    internal class MainPageTest : RunTest
    {
        Excel.Worksheet TestSheet;
        public MainPageTest(Excel.Worksheet TestSheet)
        {
            this.TestSheet = TestSheet;
        }

        public void ExecuteTest()
        {
            //Run TestCases Till End
            int exeCount = 2;
            while (true)
            {
                string testCase = (string)TestSheet.Cells[exeCount, 1].Value;

                if (testCase.Equals("CheckTitleText"))
                {
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckTitleText();
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckEmailSubmitButton"))
                {
                    string testData = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckEmailSubmitButton(testData);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckPromotionalCodeButton"))
                {
                    string testData = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckPromotionalCodeButton(testData);
                    TestSheet.Cells[exeCount, 4].Value = cBehavior;
                    WriteResult(exeCount, cBehavior.Equals(eBehavior), TestSheet);
                }
                else if (testCase.Equals("CheckWelcomeMessage"))
                {
                    string testData = (string)TestSheet.Cells[exeCount, 2].Value;
                    string eBehavior = (string)TestSheet.Cells[exeCount, 3].Value;
                    string cBehavior = CheckWelcomeMessage();
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



        public string CheckTitleText()
        {
            sUtil.GetTitle();
            return sUtil.GetTitle();
        }

        public string CheckEmailSubmitButton(string testData)
        {
            sUtil.FindElement(po.GetEmailInputBox()).SendKeys(testData);
            sUtil.FindElement(po.GetEmailSubmitButton()).Click();
            return sUtil.FindElement(po.GetEmailReturnMessageBox()).Text;
        }

        public string CheckPromotionalCodeButton(string testData)
        {
            sUtil.FindElement(po.GetPromoCodeInputBox()).SendKeys(testData);
            sUtil.FindElement(po.GetPromoSubmitButton()).Click();
            return sUtil.FindElement(po.GetEmailReturnMessageBox()).Text;
        }

        public string CheckWelcomeMessage()
        {
            return sUtil.FindElement(po.GetWelcomeMessage()).Text;
        }



    }
}
