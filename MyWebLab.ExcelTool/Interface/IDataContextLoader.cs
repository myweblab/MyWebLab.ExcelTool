using MyWebLab.ExcelTool.Data;
using System;
using System.Collections.Generic;
using System.Text;

namespace MyWebLab.ExcelTool.Interface
{
    public interface IDataContextLoader
    {
        DataContext GetDataContext();
    }
}
