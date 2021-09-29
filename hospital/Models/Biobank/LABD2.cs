using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("LABD2")]
    public class LABD2
    {
        public string h1 { get; set; } //報告類別
        [Key, Column(Order = 0)]
        public string h2 { get; set; } //醫事機構代碼
        public string h3 { get; set; } //醫事類別
        [Key, Column(Order = 1)]
        public string h4 { get; set; } //執行年月
        [Key, Column(Order = 2)]
        public string h5 { get; set; } //健保卡刷卡日期時間
        [Key, Column(Order = 3)]
        public string h6 { get; set; } //就醫類別 
        [Key, Column(Order = 4)]
        public string h7 { get; set; } //就醫序號
        public string h8 { get; set; } //補卡註記
        public string h9 { get; set; } //身分證統一編號
        public string h10 { get; set; } //出生日期
        public string h11 { get; set; } //就醫日期
        public string h12 { get; set; } //治療結束日期
        public string h13 { get; set; } //入院年月日
        public string h14 { get; set; } //出院年月日
        public string h15 { get; set; } //醫令代碼  (各院內的代碼)
        public string h19 { get; set; } //醫囑日期時間
        public string h20 { get; set; } //採檢/實際檢查/手術日期時間
        public string h22 { get; set; } //檢體採檢方法/來源/類別
        public string r1 { get; set; } //報告序號
        public string r2 { get; set; } //檢驗項目名稱
        public string r3 { get; set; } //檢驗方法
        public string r4 { get; set; } //檢驗報告結果值
        public string r5 { get; set; } //單位
        public string r6_1 { get; set; } //參考值下限
        public string r6_2 { get; set; } //參考值上限
        public string r7 { get; set; } //報告結果
        public string r8_1 { get; set; } //病理發現及診斷
        public string r10 { get; set; } //報告日期時間
        public string r12 { get; set; } //檢驗（查）結果值註記
    }
}