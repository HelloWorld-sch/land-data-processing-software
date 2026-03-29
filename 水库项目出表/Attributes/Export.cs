using System;
using System.Collections.Generic;
using System.Data;
using System.Diagnostics;
using System.IO;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using 水库项目出表.Entity;
using 水库项目出表.Util;
using OfficeOpenXml;
using OfficeOpenXml.Style;

namespace 水库项目出表.Attributes
{
    public partial class Export
    {
        private Logger log = Logger.GetLogger("Export");
        private string _saveDir;
        private DataTable _table;
        private string _reservoirName;
        private string _zbdw;
        private List<Landuse> _landuses;
        //private List<XZQ> _xzqs;
        private string _sylx;
        private Dictionary<string, int> _dlwzDic;
        private List<Sum> _sums;
        private int _digit;

        public Export(ExportParameter parameter, decimal digit)
        {
            _landuses = parameter.Landuses;
            //_xzqs = Tools.GetXZQ(parameter.DataSource, "");
            _saveDir = parameter.SaveDirectory;
            _table = parameter.DataSource.Select("使用类型='" + parameter.UseType + "'").CopyToDataTable();
            _reservoirName = parameter.ReservoirName;
            _zbdw = parameter.Unit;
            _sylx = parameter.UseType;
            _dlwzDic = Tools.GetDLWZ(ExcelTypeEnum.LandClassify);
            _sums = Tools.GetSumEntiyList(ExcelTypeEnum.LandClassify);
            _digit = (int)digit;

        }
        private ExcelWorksheet GetWorksheet(ExcelPackage package, string sheetName)
        {
            ExcelWorkbook workbook = package.Workbook;
            ExcelWorksheets worksheets = workbook.Worksheets;
            bool c = worksheets.Select(b => b.Name).Contains(sheetName, StringComparer.OrdinalIgnoreCase);
            if (!c)
                throw new Exception(package.File.Name + "读取错误");
            ExcelWorksheet worksheet = worksheets[sheetName];
            return worksheet;
        }

        private void SetBorderStyle(ExcelRange range)
        {
            range.Style.Border.Bottom.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Left.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Right.Style = ExcelBorderStyle.Thin;
            range.Style.Border.Top.Style = ExcelBorderStyle.Thin;
        }
    }
}
