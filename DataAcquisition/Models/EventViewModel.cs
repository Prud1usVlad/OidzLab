namespace DataAcquisition.Models
{
    public class EventViewModel
    {
        public int Event_id { get; set; }
        public Guid Udid { get; set; }
        public DateTime Date { get; set; }
        public Dictionary<string, string> Parameters { get; set; }
    }
}
