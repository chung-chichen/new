
namespace hospital.Models
{
    public class Disease
    {
        public string Id { get; set; }
        public string disease_name { get; set; }
    }
    public class Case_relapse
    {
        public string Id { get; set; }
        public string did { get; set; }
        public string relapse_code { get; set; }
    }
    public class Case_treatment
    {
        public string Id { get; set; }
        public string did { get; set; }
        public string treatment_code { get; set; }
    }
    public class TOTF_comorbidity
    {
        public string Id { get; set; }
        public string did { get; set; }
        public string comorbidity_code { get; set; }
        public string comorbidity_icd9 { get; set; }
        public string comorbidity_icd10 { get; set; }
    }
    public class LABH_laboratory
    {
        public string Id { get; set; }
        public string did { get; set; }
        public string laboratory_name { get; set; }
        public string laboratory_code { get; set; }
    }
    public class TOTF_medicine
    {
        public string Id { get; set; }
        public string did { get; set; }
        public string medicine_name { get; set; }
        public string medicine_code { get; set; }
    }
}
