using Microsoft.VisualStudio.TestTools.UnitTesting;
using OfficeOpenXml;
using OpenQA.Selenium;
using OpenQA.Selenium.Chrome;
using System;
using System.Collections;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text.RegularExpressions;
using System.Threading;

namespace DataDrivenTesting
{
    [TestClass]
    public class UnitTest1
    {
        [TestMethod]
        [DataRow("Narendra", "Modi", "01/01/2019")]
        [DataRow("donald", "trump", "07/01/2020")]
        [DataRow("BORIS", "JOHNSON", "12/31/2021")]
        public void DataDrivenTestingUsingDataRow(string fName, string lName, string eDate)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "http://uitestpractice.com/Students/Create";

            driver.FindElement(By.Id("FirstName")).SendKeys(fName);
            driver.FindElement(By.Id("LastName")).SendKeys(lName);
            driver.FindElement(By.Id("EnrollmentDate")).SendKeys(eDate);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();
            driver.Quit();

        }

        [DynamicData(nameof(GetData), DynamicDataSourceType.Method)]
        [TestMethod]
        public void DataDrivenTestingUsingDynamicData(string fName, string lName, string eDate)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "http://uitestpractice.com/Students/Create";

            driver.FindElement(By.Id("FirstName")).SendKeys(fName);
            driver.FindElement(By.Id("LastName")).SendKeys(lName);
            driver.FindElement(By.Id("EnrollmentDate")).SendKeys(eDate);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();

            driver.Quit();
        }

        public static IEnumerable <object[]> ReadExcel()
        {
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            // create worksheet object
            //string path = Path.Combine(Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location), @"Libro.xlsx");
            byte[] file = File.ReadAllBytes("Libro.xlsx");
            using (MemoryStream ms = new MemoryStream(file))
            using (ExcelPackage package = new ExcelPackage(ms))
            {
                //get the first worksheet in the workbook
                ExcelWorksheet worksheet = package.Workbook.Worksheets["Sheet1"];
                int rowCount = worksheet.Dimension.End.Row;     //get row count
                for (int row = 2; row <= rowCount; row++)
                {
                    yield return new object[] {
                        worksheet.Cells[row, 1].Value?.ToString().Trim(), // First name
                        worksheet.Cells[row, 2].Value?.ToString().Trim(), // Last name
                        worksheet.Cells[row, 3].Value?.ToString().Trim()  // Enrollment date
                    };
                }
            }
        }


        [DynamicData(nameof(ReadExcel), DynamicDataSourceType.Method)]
        [TestMethod]
        public void DataDrivenTestingUsingExcelSheet(string fName, string lName, string eDate)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "http://uitestpractice.com/Students/Create";

            driver.FindElement(By.Id("FirstName")).SendKeys(fName);
            driver.FindElement(By.Id("LastName")).SendKeys(lName);
            driver.FindElement(By.Id("EnrollmentDate")).SendKeys(eDate);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();

            driver.Quit();
        }

        [DynamicData(nameof(ReadCsv), DynamicDataSourceType.Method)]
        [TestMethod]
        public void DataDrivenTestingUsingCSVFile(string fName, string lName, string eDate)
        {
            IWebDriver driver = new ChromeDriver();
            driver.Manage().Window.Maximize();
            driver.Url = "http://uitestpractice.com/Students/Create";

            driver.FindElement(By.Id("FirstName")).SendKeys(fName);
            driver.FindElement(By.Id("LastName")).SendKeys(lName);
            driver.FindElement(By.Id("EnrollmentDate")).SendKeys(eDate);
            driver.FindElement(By.XPath("//input[@type='submit']")).Click();


            Thread.Sleep(2000);
            driver.Quit();
        }
        private static string[] SplitCsv(string input)
        {
            var csvSplit = new Regex("(?:^|,)(\"(?:[^\"]+|\"\")*\"|[^,]*)", RegexOptions.Compiled);
            var list = new List<string>();
            foreach (Match match in csvSplit.Matches(input))
            {
                string value = match.Value;
                if (value.Length == 0)
                {
                    list.Add(string.Empty);
                }

                list.Add(value.TrimStart(','));
            }
            return list.ToArray();
        }
        private static IEnumerable<string[]> ReadCsv()
        {
            IEnumerable<string> rows = System.IO.File.ReadAllLines("data.csv").Skip(1);
            foreach (string row in rows)
            {
                yield return SplitCsv(row);
            }
        }
        public static IEnumerable<object[]> GetData()
        {
            yield return new object[] { "Narendra", "Modi", "01/01/2019" };
            yield return new object[] { "donald", "trump", "07/01/2020" };
            yield return new object[] { "BORIS", "JOHNSON", "12/31/2021" };
        }
    }

}
