using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Models
{
    public class UserClusteringModel
    {
        public Guid Id { get; set; }

        public int? Cluster { get; set; }

        public double Value { get; set; }
    }
}
