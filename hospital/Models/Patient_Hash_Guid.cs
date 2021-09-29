using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.Linq;
using System.Web;

namespace hospital.Models
{
    public class Patient_Hash_Guid
    {
        [Key]
        public string Patient_Id { get; set; }
        [Required]
        public string Patient_Birth { get; set; }
    }
}