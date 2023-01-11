namespace TestApp.Models
{
    public class TestModel
    {
        public int Id { get; set; }
        public string Name { get; set; }
        public string Surname { get; set; }
        public DateTime StartDate { get; set; }
        public DateTime FinishDate { get; set; }
        public string DealerList { get; set; }
        public string FuelList { get; set; }
        public string FuelCode { get; set; }
        public int? Control { get; set; }
        public string TableControl { get; set; }
        public int? Explanation { get; set; }
        public string ExplanationTypes { get; set; }
        public string PageType { get; set; }
        public string DataCompany { get; set; } //Entegrasyon yapılan firmalar
        public string DataSelectCompany { get; set; } //Listede olan firmalar
    }
}
