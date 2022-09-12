using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace 水库项目出表.Entity
{
   public class XZQ
    {
        public string City { get; set; }
        public string County { get; set; }
        public string Town { get; set; }
        public string Village { get; set; }
        public string Group { get; set; }

        public XZQ()
        {
            
        }
        public XZQ(string city,string county,string town,string village,string group)
        {
            City = city;
            County = county;
            Town = town;
            Village = village;
            Group = group;
        }

        public XZQ GetVillage()
        {
            return new XZQ(this.City,this.County,this.Town,this.Village,"");
        }
        public XZQ GetTown()
        {
            return new XZQ(this.City, this.County, this.Town, "", "");
        }
        public XZQ GetCounty()
        {
            return new XZQ(this.City, this.County, "", "", "");
        }
    }
}
