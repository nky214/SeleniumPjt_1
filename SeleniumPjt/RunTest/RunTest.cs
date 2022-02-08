using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumPjt.RunTest
{
    internal class RunTest
    {
        protected PageObjectRepository po;
        protected SeleniumUtil sUtil;

        public RunTest()
        {
            this.po = PageObjectRepository.GetInstance();
            this.sUtil = SeleniumUtil.GetInstance();
        }

        public void WriteResult(int exeCount, Boolean result, Excel.Worksheet TestSheet)
        {
            if (result)
            {
                TestSheet.Cells[exeCount, 5].Value = "PASS";
                TestSheet.Cells[exeCount, 5].Interior.Color = 0x00ff00;
            }
            else
            {
                TestSheet.Cells[exeCount, 5].Value = "FAIL";
                TestSheet.Cells[exeCount, 5].Interior.Color = 0x0000ff;
            }
        }
    }
}
