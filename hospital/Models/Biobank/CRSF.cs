using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("CRSF")]
    public class CRSF
    {
        [Key, Column(Order = 0)]
        public string SF1_1 { get; set; } // 申報醫院代碼
        public string SF1_2 { get; set; } // 病歷號碼
        public string SF1_3 { get; set; } // 姓名(此欄位不需要)
        [Key, Column(Order = 1)]
        public string SF1_4 { get; set; } // 身分證統一編號
        public string SF1_5 { get; set; } // 性別
        [Key, Column(Order = 2)]
        public string SF1_6 { get; set; } // 出生日期
        public string SF1_7 { get; set; } // 戶籍地代碼
        public string SF2_1 { get; set; } // 診斷年齡
        public string SF2_2 { get; set; } // 癌症發生順序號碼
        public string SF2_3 { get; set; } // 個案分類
        public string SF2_3_1 { get; set; } // 診斷狀態分類
        public string SF2_3_2 { get; set; } // 治療狀態分類
        [Key, Column(Order = 3)]
        public string SF2_4 { get; set; } // 首次就診日期
        public string SF2_5 { get; set; } // 最初診斷日期
        public string SF2_6 { get; set; } // 原發部位
        public string SF2_7 { get; set; } // 側性
        public string SF2_8 { get; set; } // 組織類型
        public string SF2_9 { get; set; } // 性態碼
        public string SF2_10_1 { get; set; } // 臨床分級/分化
        public string SF2_10_2 { get; set; } // 病理分級/分化
        public string SF2_11 { get; set; } // 癌症確診方式
        public string SF2_12 { get; set; } // 首次顯微鏡檢證實日期
        public string SF4_1_1 { get; set; } // 首次手術日期
        public string SF4_1_4 { get; set; } // 申報醫院原發部位手術方式
        public string SF4_2_1_3 { get; set; } // 放射治療開始日期
        public string SF4_2_1_7 { get; set; } // 放射治療機構
        public string SF4_3_3 { get; set; } // 申報醫院化學治療
        public string SF4_3_4 { get; set; } // 申報醫院化學治療開始日期
        public string SF4_3_6 { get; set; } // 申報醫院賀爾蒙/類固醇治療
        public string SF4_3_7 { get; set; } // 申報醫院賀爾蒙/類固醇治療開始日期
        public string SF4_3_9 { get; set; } // 申報醫院免疫治療
        public string SF4_3_10 { get; set; } // 申報醫院免疫治療開始日期
        public string SF4_3_11 { get; set; } // 骨髓/幹細胞移植或內分泌處置
        public string SF4_3_12 { get; set; } // 申報醫院骨髓/幹細胞移植或內分泌處置開始日期
        public string SF4_3_14 { get; set; } // 申報醫院標靶治療
        public string SF4_3_15 { get; set; } // 申報醫院標靶治療開始日期
        public string SF4_4 { get; set; } // 申報醫院緩和照護
        public string SF4_5_1 { get; set; } // 其他治療
        public string SF4_5_2 { get; set; } // 其他治療開始日期
        public string SF6_1 { get; set; } // 摘錄者(此欄位沒有被寫入)
        public string SF7_1 { get; set; } // 身高
        public string SF7_2 { get; set; } // 體重
        public string SF7_3 { get; set; } // 吸菸行為
        public string SF7_4 { get; set; } // 嚼檳榔行為
        public string SF7_5 { get; set; } // 喝酒行為
    }
}