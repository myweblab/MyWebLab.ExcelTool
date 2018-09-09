using System;
using System.Collections.Generic;
using System.Text;

namespace MyWebLab.ExcelTool.Data
{
    public class DataContext
    {
        public string SheetName { get; set; }
        public IList<DataItem> Items { get; set; }
    }
}
