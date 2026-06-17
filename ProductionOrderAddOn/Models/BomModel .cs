using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ProductionOrderAddOn.Models
{
    class BomModel
    {
        public string PartNo { get; set; }
        public string PartName { get; set; }
        public double Quantity { get; set; }
    }
}
