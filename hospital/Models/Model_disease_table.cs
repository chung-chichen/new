using System.ComponentModel.DataAnnotations;

namespace hospital.Models
{
    public class ICD9
    {
        [Key]
        public string Code { get; set; }
        public string Eng { get; set; }
        public string Ch { get; set; }
    }
    public class ICD10
    {
        [Key]
        public string Code { get; set; }
        public string Eng { get; set; }
        public string Ch { get; set; }
    }
}
