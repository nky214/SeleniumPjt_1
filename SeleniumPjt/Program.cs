using System;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;

namespace SeleniumPjt
{
    internal class Program
    {
        static string cDirectory = Directory.GetParent(Directory.GetParent(Directory.GetCurrentDirectory()).ToString()).ToString();
        static string TestDataExcel = cDirectory + @"\TestData\TestDataExcel.xlsx";
        static PageObjectRepository po = PageObjectRepository.GetInstance();
        static SeleniumUtil sUtil = SeleniumUtil.GetInstance();
       
        static void Main(string[] args)
        {
            Console.WriteLine("SeleniumTest Start");
            
            //Open Test Data Excel
            Excel.Application configExcelAppMaster = new Excel.Application();
            Excel.Workbook configTestCaseMaster = configExcelAppMaster.Workbooks.Open(System.IO.Path.GetFullPath(TestDataExcel));
            Excel.Worksheet configDataSheet = (Excel.Worksheet)configExcelAppMaster.Sheets["TestConfiguration"];
            
            //Load Chrome Driver and GoToURL
            sUtil.LoadChromeDriver();
            sUtil.GoToTargetURL(po.GetURL());

                int exeCount = 2;
                while (true)
                {
                    try
                    {
                        if ((string)configDataSheet.Cells[1, 1].Value == "Execute")
                        {
                            string execute = (string)configDataSheet.Cells[exeCount, 1].Value;
                            string testItem = (string)configDataSheet.Cells[exeCount, 2].Value;
                            Console.WriteLine("-----"+testItem+"-----Execute: " + execute);

                            if (execute.ToUpper().Equals("Y"))
                            {
                                Excel.Worksheet TestSheet = (Excel.Worksheet)configExcelAppMaster.Sheets[testItem];

                                if (testItem.Equals("MainPageTest"))
                                {
                                    new RunTest.MainPageTest(TestSheet).ExecuteTest();
                                }
                                else if (testItem.Equals("MarchPageTest"))
                                {
                                    new RunTest.MarchPageTest(TestSheet).ExecuteTest();
                                }
                            }
                        }
                    }
                    catch (Exception ex)
                    {
                        Console.WriteLine("-----Exception-----");
                        Console.WriteLine(ex);
                    }
                    if ((string)configDataSheet.Cells[exeCount + 1, 1].Value == null)
                    {
                        break;
                    }
                    exeCount++;
                }
                
                Console.WriteLine("-----Quit Chrome Driver-----");
                sUtil.QuitChromeDriver();
                Console.WriteLine("-----Save Result Excel-----");
                string cResultDirectory = cDirectory + @"\TestResult\";
                configTestCaseMaster.SaveAs(cResultDirectory+"TestResult_"+DateTime.Now.ToString().Replace(":","_").Replace("/","_")+".xlsx");
                configTestCaseMaster.Close(0);
                configExcelAppMaster.Quit();
                Console.WriteLine("-----Test End-----");
                


        }
    }
}
