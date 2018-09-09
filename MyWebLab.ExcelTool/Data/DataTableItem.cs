using System;
using System.Collections.Generic;
using System.Text;
using OfficeOpenXml;

namespace MyWebLab.ExcelTool.Data
{
    public class DataTableItem : DataItem
    {
        public override void WriteContent(ExcelWorksheet ws)
        {
            ws.Cells[Range].LoadFromDataTable(Value, false);
        }
    }
}
