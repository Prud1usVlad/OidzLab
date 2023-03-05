using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace DataAcquisition.Models.DataModels
{
    public class EventViewModel
    {
        public Guid Udid { get; set; }
        public DateOnly Date { get; set; }
        public int Event_id { get; set; }
        public Dictionary<string, string> Parameters { get; set; } 
    }
}
