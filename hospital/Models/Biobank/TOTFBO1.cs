using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("TOTFBO1")]
    public class TOTFBO1
    {
        [Index]
        public int Id { get; set; }
        [ForeignKey("TOTFBE"), Column(Order = 0)]
        public string t2 { get; set; } //服務機構代號
        [ForeignKey("TOTFBE"), Column(Order = 1)]
        public string t3 { get; set; } //費用年月
        [ForeignKey("TOTFBE"), Column(Order = 2)]
        public string t5 { get; set; } //申報類別
        [ForeignKey("TOTFBE"), Column(Order = 3)]
        public string t6 { get; set; } //申報日期
        [ForeignKey("TOTFBE"), Column(Order = 4)]
        public string d1 { get; set; } //案件分類
        [ForeignKey("TOTFBE"), Column(Order = 5)]
        public string d2 { get; set; } //流水編號
        public string p1 { get; set; } //醫令序
        public string p2 { get; set; } //醫令類別
        public string p3 { get; set; } //醫令代碼 (各院內的代碼)
        public string p5 { get; set; } //藥品用量
        public string p6 { get; set; } //(藥品)使用頻率
        public string p7 { get; set; } //給藥途徑/作用部位
        public string p8 { get; set; } //會診科別
        public string p14 { get; set; } //執行時間-起
        public string p15 { get; set; } //執行時間-迄
        public string p16 { get; set; } //總量

        public virtual TOTFBE TOTFBE { get; set; }
    }
}