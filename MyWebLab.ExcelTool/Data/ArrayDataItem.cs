using MyWebLab.ExcelTool.Data;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Text;

namespace MyWebLab.ExcelTool
{
    public class ArrayDataItem : DataItem
    {
        public override void WriteContent(ExcelWorksheet ws)
        {
            ws.Cells[Range].LoadFromArrays(Value);
        }
    }
}
