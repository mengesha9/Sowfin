namespace Sowfin.API.ViewModels
{
    public class sensitivity
    {
        public List<string> variableList { get; set; }
        public long ProjectId { get; set; }
        public string linItem { get; set; }
        public double difference { get; set; }
        public bool intervalFlag { get; set; }
    }
}
  