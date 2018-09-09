using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;
using MyWebLab.ExcelTool.Data;
using MyWebLab.ExcelTool.Interface;

namespace MyWebLab.ExcelTool.Data
{
    public class DataItem
    {
    
        public dynamic Value { get; set; }

        public ContextStyle Style { get; set; }

        public string Range { get; set; }

        public virtual void WriteContent(ExcelWorksheet ws)
        {
            ws.Cells[Range].Value = Value;
        }
    }
}
