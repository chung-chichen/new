using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("DEATH")]
    public class DEATH
    {
        [Key, Column(Order = 0)]
        public string SSN { get; set; } //身分證字號
        public string d2 { get; set; } //性別
        [Key, Column(Order = 1)]
        public string d3 { get; set; } //出生年月
        public string d4 { get; set; } //死亡日期
        public string d5 { get; set; } //死因分類一
        public string d6 { get; set; } //死因分類二
        public string d7 { get; set; } //死因分類三
    }
}