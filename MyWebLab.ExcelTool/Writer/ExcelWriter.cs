using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;
using System.IO;
using MyWebLab.ExcelTool.Data;
using MyWebLab.ExcelTool.Interface;

namespace MyWebLab.ExcelTool.Writer
{
    public class ExcelWriter : IExcelWriter
    {
        private readonly DataContext context;
        ExcelPackage _excelPacakge;
        public ExcelWriter(string filePath, DataContext context)
        {
            _excelPacakge = new ExcelPackage(new FileInfo(filePath));
            this.context = context;
        }

        public ExcelWriter(string filePath, string templatePath, DataContext context)
        {
            _excelPacakge = new ExcelPackage(new FileInfo(filePath), new FileInfo(templatePath));
            this.context = context;
        }

        public ExcelWriter(Stream filePath,Stream content, DataContext context)
        {
            _excelPacakge = new ExcelPackage(filePath,content);
            this.context = context;
        }

        public void WriteContext()
        {
            var ws = _excelPacakge.Workbook.Worksheets[context.SheetName];
            if (ws == null)
                throw new ArgumentException("Invalid Worksheet Name");
            foreach (var item in this.context.Items)
            {
                item.WriteContent(ws);
            }
        }

        public void Save()
        {
            _excelPacakge.Save();

        }

        public void SaveAs(string filePath)
        {
            _excelPacakge.SaveAs(new FileInfo(filePath));
        }
    }
}
