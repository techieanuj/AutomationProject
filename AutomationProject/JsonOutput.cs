using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AutomationProject
{
    public class JsonOutput
    {
        public string Domain { get; set; }
        public string Market { get; set; }
        public List<Coupon> Coupons { get; set; }

        public JsonOutput()
        {
            Coupons = new List<Coupon>();
        }
    }
}
