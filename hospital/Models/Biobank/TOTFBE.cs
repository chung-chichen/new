using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("TOTFBE")]
    public class TOTFBE
    {
        [Key, Column(Order = 0)]
        public string t2 { get; set; } //服務機構代號
        [Key, Column(Order = 1)]
        public string t3 { get; set; } //費用年月
        [Key, Column(Order = 2)]
        public string t5 { get; set; } //申報類別
        [Key, Column(Order = 3)]
        public string t6 { get; set; } //申報日期
        [Key, Column(Order = 4)]
        public string d1 { get; set; } //案件分類
        [Key, Column(Order = 5)]
        public string d2 { get; set; } //流水編號
        public string d3 { get; set; } //身分證統一編號
        public string d6 { get; set; } //出生年月日
        public string d9 { get; set; } //就醫科別
        public string d10 { get; set; } //入院年月日
        public string d11 { get; set; } //出院年月日
        public string d14 { get; set; } //急性病床天數
        public string d15 { get; set; } //慢性病床天數
        public string d18 { get; set; } //Tw-DRG碼
        public string d21 { get; set; } //DRGs碼
        public string d24 { get; set; } //轉歸代碼
        public string d25 { get; set; } //主診斷
        public string d26 { get; set; } //次診斷代碼(一)
        public string d27 { get; set; } //次診斷代碼(二)
        public string d28 { get; set; } //次診斷代碼(三)
        public string d29 { get; set; } //次診斷代碼(四)
        public string d45 { get; set; } //主手術(處置)代碼
        public string d46 { get; set; } //次手術(處置)代碼一
        public string d47 { get; set; } //次手術(處置)代碼二
        public string d48 { get; set; } //次手術(處置)代碼三
        public string d49 { get; set; } //次手術(處置)代碼四

        public virtual ICollection<TOTFBO1> TOTFBO1 { get; set; }
        public virtual ICollection<TOTFBO2> TOTFBO2 { get; set; }
    }
}