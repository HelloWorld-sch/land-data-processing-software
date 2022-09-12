using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Text;

namespace 水库项目出表.Entity
{
    public class ExportParameter
    {
        public DataTable DataSource { get; set; }
        public string SaveDirectory { get; set; }
        public string ReservoirName { get; set; }
        public string Unit { get; set; }
        public string UseType { get; set; }
        public List<Landuse> Landuses { get; set; }
    }
}
