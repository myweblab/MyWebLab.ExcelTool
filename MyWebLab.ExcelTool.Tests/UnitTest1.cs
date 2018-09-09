using Microsoft.VisualStudio.TestTools.UnitTesting;
using MyWebLab.ExcelTool.Data;
using MyWebLab.ExcelTool.Writer;
using System.Collections.Generic;
using System.IO;
using System.Reflection;
using System.Text;

namespace MyWebLab.ExcelTool.Tests
{
    [TestClass]
    public class ExcelWriteTest
    {
        [TestMethod]
        public void Verify_FileWrite()
        {
            FileStream stream = new FileStream("FO.xlsx",FileMode.Create);
            var assembly = Assembly.GetExecutingAssembly();
            var resourceName = "MyWebLab.ExcelTool.Tests.Template.TemplateTest1.xlsx";
            Stream templateStram = assembly.GetManifestResourceStream(resourceName);
            var testContext = new DataContext();
            testContext.SheetName = "Sheet1";
            testContext.Items = new List<DataItem>();
            DataItem item = new DataItem();
            item.Value = 100;
            item.Range = "A1";
            testContext.Items.Add(item);
            item = new ArrayDataItem();
            item.Value = new List<object[]>() { new object[] { 1, 2, 3 }, new object[] { 4, 5, 6 } };
            item.Range = "B2:B3";
            testContext.Items.Add(item);

            ExcelWriter excelWriter = new ExcelWriter(stream, templateStram, testContext );
            excelWriter.WriteContext();
            excelWriter.Save();
            Assert.IsTrue(File.Exists("FO.xlsx"));
            
        }
    }
}
