using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("CASE")]
    public class CASE
    {
        [Key, Column(Order = 0)]
        public string SSN { get; set; } //身分證統一編號
        [Key, Column(Order = 1)]
        public string d3 { get; set; } //出生年月日
        public string m2 { get; set; } //輸入日期
        public string m3 { get; set; } //追蹤日期
        public string m4 { get; set; } //追蹤方式
        public string m5 { get; set; } //復發狀態
        public string m6 { get; set; } //治療反應
    }
}