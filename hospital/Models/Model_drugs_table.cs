using System.ComponentModel.DataAnnotations;

namespace hospital.Models
{
    public class New_NHI_Drugs_item
    {
        [Key]
        public double id { get; set; }
        public string ATC_CODE { get; set; }
        public string Drug_code { get; set; }
        public string License { get; set; }
        public string Change { get; set; }
        public string Drug_name_eng { get; set; }
        public string Drug_name_Ch { get; set; }
        public double? Specification { get; set; }
        public string Unit { get; set; }
        public string Single_compound { get; set; }
        public double? Reference_Price { get; set; }
        public double? Effective_date_start { get; set; }
        public double? Effective_date_end { get; set; }
        public string Manufacturer { get; set; }
        public string Dosage { get; set; }
        public string Ingredients { get; set; }
        public string ATC_CODE1 { get; set; }
        public string Ingredients1 { get; set; }
        public string unit1 { get; set; }
        public string Ingredients2 { get; set; }
        public string unit2 { get; set; }
        public string Ingredients3 { get; set; }
        public string unit3 { get; set; }
        public string Ingredients4 { get; set; }
        public string unit4 { get; set; }
        public string Ingredients5 { get; set; }
        public string unit5 { get; set; }
        public string Ingredients6 { get; set; }
        public string unit6 { get; set; }
        public string Ingredients7 { get; set; }
        public string unit7 { get; set; }
        public string Ingredients8 { get; set; }
        public string unit8 { get; set; }
        public string Ingredients9 { get; set; }
        public string unit9 { get; set; }
        public string Ingredients10 { get; set; }
        public string unit10 { get; set; }
        public string Ingredients11 { get; set; }
        public string unit11 { get; set; }
        public string Ingredients12 { get; set; }
    }
}
