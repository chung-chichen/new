using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("TOTFAE")]
    public class TOTFAE
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
        public string d4 { get; set; } //特定治療項目代號(一)
        public string d5 { get; set; } //特定治療項目代號(二)
        public string d6 { get; set; } //特定治療項目代號(三)
        public string d7 { get; set; } //特定治療項目代號(四)
        public string d8 { get; set; } //就醫科別
        public string d9 { get; set; } //就醫日期
        public string d10 { get; set; } //治療結束日期
        public string d11 { get; set; } //出生年月日
        public string d3 { get; set; } //身分證統一編號
        public string d19 { get; set; } //主診斷代碼
        public string d20 { get; set; } //次診斷代碼(一)
        public string d21 { get; set; } //次診斷代碼(二)
        public string d22 { get; set; } //次診斷代碼(三)
        public string d23 { get; set; } //次診斷代碼(四)
        public string d24 { get; set; } //主手術(處置)代碼
        public string d25 { get; set; } //次手術(處置)代碼(一)
        public string d26 { get; set; } //次手術(處置)代碼(二)
        public string d27 { get; set; } //給藥日份
        public string d28 { get; set; } //處方調劑方式

        public virtual ICollection<TOTFAO1> TOTFAO1 { get; set; }
        public virtual ICollection<TOTFAO2> TOTFAO2 { get; set; }
    }
}