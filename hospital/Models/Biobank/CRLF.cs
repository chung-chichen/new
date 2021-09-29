using System;
using System.Collections.Generic;
using System.ComponentModel.DataAnnotations;
using System.ComponentModel.DataAnnotations.Schema;
using System.Linq;
using System.Web;

namespace hospital.Models.Biobank
{
    [Table("CRLF")]
    public class CRLF
    {
        [Key, Column(Order = 0)]
        public string LF1_1 { get; set; } // 申報醫院代碼
        public string LF1_2 { get; set; } // 病歷號碼
        public string LF1_3 { get; set; } // 姓名(此欄位不需要)
        [Key, Column(Order = 1)]
        public string LF1_4 { get; set; } // 身分證統一編號
        public string LF1_5 { get; set; } // 性別
        [Key, Column(Order = 2)]
        public string LF1_6 { get; set; } // 出生日期
        public string LF1_7 { get; set; } // 戶籍地代碼
        public string LF2_1 { get; set; } // 診斷年齡
        public string LF2_2 { get; set; } // 癌症發生順序號碼
        public string LF2_3 { get; set; } // 個案分類
        public string LF2_3_1 { get; set; } // 診斷狀態分類
        public string LF2_3_2 { get; set; } // 治療狀態分類
        [Key, Column(Order = 3)]
        public string LF2_4 { get; set; } // 首次就診日期
        public string LF2_5 { get; set; } // 最初診斷日期
        public string LF2_6 { get; set; } // 原發部位
        public string LF2_7 { get; set; } // 側性
        public string LF2_8 { get; set; } // 組織類型
        public string LF2_9 { get; set; } // 性態碼
        public string LF2_10_1 { get; set; } // 臨床分級/分化
        public string LF2_10_2 { get; set; } // 病理分級/分化
        public string LF2_11 { get; set; } // 癌症確診方式
        public string LF2_12 { get; set; } // 首次顯微鏡檢證實日期
        public string LF2_13 { get; set; } // 腫瘤大小
        public string LF2_13_1 { get; set; } // 神經侵襲
        public string LF2_13_2 { get; set; } // 淋巴管或血管侵犯
        public string LF2_14 { get; set; } // 區域淋巴結檢查數目
        public string LF2_15 { get; set; } // 區域淋巴結侵犯數目
        public string LF3_1 { get; set; } // 診斷性及分期性手術處置日期
        public string LF3_2 { get; set; } // 外院診斷性及分期性手術處置
        public string LF3_3 { get; set; } // 申報醫院診斷性及分期性手術處置
        public string LF3_4 { get; set; } // 臨床T
        public string LF3_5 { get; set; } // 臨床N
        public string LF3_6 { get; set; } // 臨床M
        public string LF3_7 { get; set; } // 臨床期別組合
        public string LF3_8 { get; set; } // 臨床分期字根/字首
        public string LF3_10 { get; set; } // 病理T
        public string LF3_11 { get; set; } // 病理N
        public string LF3_12 { get; set; } // 病理M
        public string LF3_13 { get; set; } // 病理期別組合
        public string LF3_14 { get; set; } // 病理分期字根/字首
        public string LF3_16 { get; set; } // AJCC癌症分期版本與章節
        public string LF3_17 { get; set; } // 其他分期系統
        public string LF3_19 { get; set; } // 其他分期系統期別(臨床分期)
        public string LF3_21 { get; set; } // 其他分期系統期別(病理分期)
        public string LF4_1 { get; set; } // 首次療程開始日期
        public string LF4_1_1 { get; set; } // 首次手術日期
        public string LF4_1_2 { get; set; } // 原發部位最確切的手術切除日期
        public string LF4_1_3 { get; set; } // 外院原發部位手術方式
        public string LF4_1_4 { get; set; } // 申報醫院原發部位手術方式
        public string LF4_1_4_1 { get; set; } // 微創手術
        public string LF4_1_5 { get; set; } // 原發部位手術邊緣
        public string LF4_1_5_1 { get; set; } // 原發部位手術切緣距離
        public string LF4_1_6 { get; set; } // 外院區域淋巴結手術範圍
        public string LF4_1_7 { get; set; } // 申報醫院區域淋巴結手術範圍
        public string LF4_1_8 { get; set; } // 外院其他部位手術方式
        public string LF4_1_9 { get; set; } // 申報醫院其他部位手術方式
        public string LF4_1_10 { get; set; } // 原發部位未手術原因
        public string LF4_2_1_1 { get; set; } // 放射治療臨床標靶體積摘要
        public string LF4_2_1_2 { get; set; } // 放射治療儀器
        public string LF4_2_1_3 { get; set; } // 放射治療開始日期
        public string LF4_2_1_4 { get; set; } // 放射治療結束日期
        public string LF4_2_1_5 { get; set; } // 放射治療與手術順序
        public string LF4_2_1_6 { get; set; } // 區域治療與全身性治療順序
        public string LF4_2_1_8 { get; set; } // 放射治療執行狀態
        public string LF4_2_2_1 { get; set; } // 體外放射治療技術
        public string LF4_2_2_2_1 { get; set; } // 最高放射劑量臨床標靶體積
        public string LF4_2_2_2_2 { get; set; } // 最高放射劑量臨床標靶體積劑量
        public string LF4_2_2_2_3 { get; set; } // 最高放射劑量臨床標靶體積治療次數
        public string LF4_2_2_3_1 { get; set; } // 較低放射劑量臨床標靶體積
        public string LF4_2_2_3_2 { get; set; } // 最低放射劑量臨床標靶體積劑量
        public string LF4_2_2_3_3 { get; set; } // 最低放射劑量臨床標靶體積治療次數
        public string LF4_2_3_1 { get; set; } // 其他放射治療儀器
        public string LF4_2_3_2 { get; set; } // 其他放射治療技術
        public string LF4_2_3_3_1 { get; set; } // 其他放射治療臨床標靶體積
        public string LF4_2_3_3_2 { get; set; } // 其他放射劑量臨床標靶體積劑量
        public string LF4_2_3_3_3 { get; set; } // 其他放射劑量臨床標靶體積治療次數
        public string LF4_3_1 { get; set; } // 全身性治療開始日期
        public string LF4_3_2 { get; set; } // 外院化學治療
        public string LF4_3_3 { get; set; } // 申報醫院化學治療
        public string LF4_3_4 { get; set; } // 申報醫院化學治療開始日期
        public string LF4_3_5 { get; set; } // 外院賀爾蒙/類固醇治療
        public string LF4_3_6 { get; set; } // 申報醫院賀爾蒙/類固醇治療
        public string LF4_3_7 { get; set; } // 申報醫院賀爾蒙/類固醇治療開始日期
        public string LF4_3_8 { get; set; } // 外院免疫治療
        public string LF4_3_9 { get; set; } // 申報醫院免疫治療
        public string LF4_3_10 { get; set; } // 申報醫院免疫治療開始日期
        public string LF4_3_11 { get; set; } // 骨髓/幹細胞移植或內分泌處置
        public string LF4_3_12 { get; set; } // 申報醫院骨髓/幹細胞移植或內分泌處置開始日期
        public string LF4_3_13 { get; set; } // 外院標靶治療
        public string LF4_3_14 { get; set; } // 申報醫院標靶治療
        public string LF4_3_15 { get; set; } // 申報醫院標靶治療開始日期
        public string LF4_4 { get; set; } // 申報醫院緩和照護
        public string LF4_5_1 { get; set; } // 其他治療
        public string LF4_5_2 { get; set; } // 其他治療開始日期
        public string LF5_1 { get; set; } // 首次復發日期
        public string LF5_2 { get; set; } // 首次復發型式
        public string LF5_3 { get; set; } // 最後聯絡或死亡日期
        public string LF5_4 { get; set; } // 生存狀態
        public string LF6_1 { get; set; } // 摘錄者(此欄位沒有被寫入)
        public string LF7_1 { get; set; } // 身高
        public string LF7_2 { get; set; } // 體重
        public string LF7_3 { get; set; } // 吸菸行為
        public string LF7_4 { get; set; } // 嚼檳榔行為
        public string LF7_5 { get; set; } // 喝酒行為
        public string LF7_6 { get; set; } // 首次治療前生活功能狀態評估
        public string LF8_1 { get; set; } // 癌症部位特定因子1
        public string LF8_2 { get; set; } // 癌症部位特定因子2
        public string LF8_3 { get; set; } // 癌症部位特定因子3
        public string LF8_4 { get; set; } // 癌症部位特定因子4
        public string LF8_5 { get; set; } // 癌症部位特定因子5
        public string LF8_6 { get; set; } // 癌症部位特定因子6
        public string LF8_7 { get; set; } // 癌症部位特定因子7
        public string LF8_8 { get; set; } // 癌症部位特定因子8
        public string LF8_9 { get; set; } // 癌症部位特定因子9
        public string LF8_10 { get; set; } // 癌症部位特定因子10
    }
}