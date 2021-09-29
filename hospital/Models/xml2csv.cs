
namespace hospital.Models
{
    //各表單儲存身份證與生日欄位
    public class Hash_table
    {
        public string Type_name { get; set; }
        public string Patient_Id_col { get; set; }
        public string Patient_Birth_col { get; set; }
    }

    public class data_result
    {
        public string alert { get; set; }
        public string result { get; set; }
        public string fileZip { get; set; }
        public string fileError { get; set; }
    }
    public class TOTFA_Tdata
    {
        public string t2 { get; set; } // 服務機構代號
        public string t3 { get; set; } // 費用年月
        public string t5 { get; set; } // 申報類別
        public string t6 { get; set; } // 申報日期
    }

    public class TOTFA_Dhead
    {
        public string d1 { get; set; } // 案件分類
        public string d2 { get; set; } // 流水編號
    }


    public class TOTFA_Dbody
    {
        public string d4 { get; set; } // 特定治療項目代號(一)
        public string d5 { get; set; } // 特定治療項目代號(二)
        public string d6 { get; set; } // 特定治療項目代號(三)
        public string d7 { get; set; } // 特定治療項目代號(四)
        public string d8 { get; set; } // 就醫科別
        public string d9 { get; set; } // 就醫日期
        public string d10 { get; set; } // 治療結束日期
        public string d11 { get; set; } // 出生年月日
        public string d3 { get; set; } // 身分證統一編號
        public string d19 { get; set; } // 主診斷代碼
        public string d20 { get; set; } // 次診斷代碼(一)
        public string d21 { get; set; } // 次診斷代碼(二)
        public string d22 { get; set; } // 次診斷代碼(三)
        public string d23 { get; set; } // 次診斷代碼(四)
        public string d24 { get; set; } // 主手術(處置)代碼
        public string d25 { get; set; } // 次手術(處置)代碼(一)
        public string d26 { get; set; } // 次手術(處置)代碼(二)
        public string d27 { get; set; } // 給藥日份
        public string d28 { get; set; } // 處方調劑方式
    }

    public class TOTFA_Pdata
    {
        public string p1 { get; set; } // 藥品給藥日份
        public string p2 { get; set; } // 醫令調劑方式
        public string p3 { get; set; } // 醫令類別
        public string p4 { get; set; } // 藥品(項目)代號 / 各院內的代碼
        public string p5 { get; set; } // 藥品用量
        public string p6 { get; set; } // 診療之部位
        public string p7 { get; set; } // 藥品使用頻率
        public string p9 { get; set; } // 給藥途徑/作用部位
        public string p10 { get; set; } // 總量
        public string p13 { get; set; } // 醫令序
        public string p14 { get; set; } // 執行時間-起
        public string p15 { get; set; } // 執行時間-迄
        public string p17 { get; set; } // 慢性病連續處方箋、同一療程及排程檢查案件註記
    }

    public class TOTFB_Tdata
    {
        public string t2 { get; set; } // 服務機構代號
        public string t3 { get; set; } // 費用年月
        public string t5 { get; set; } // 申報類別
        public string t6 { get; set; } // 申報日期
    }

    public class TOTFB_Dhead
    {
        public string d1 { get; set; } // 案件分類
        public string d2 { get; set; } // 流水編號
    }


    public class TOTFB_Dbody
    {
        public string d3 { get; set; } // 身分證統一編號
        public string d6 { get; set; } // 出生年月日
        public string d9 { get; set; } // 就醫科別
        public string d10 { get; set; } // 入院年月日 
        public string d11 { get; set; } // 出院年月日 
        public string d14 { get; set; } // 急性病床天數
        public string d15 { get; set; } // 慢性病床天數
        public string d18 { get; set; } // Tw-DRG碼
        public string d21 { get; set; } // DRGs碼
        public string d24 { get; set; } // 轉歸代碼
        public string d25 { get; set; } // 主診斷
        public string d26 { get; set; } // 次診斷代碼(一)
        public string d27 { get; set; } // 次診斷代碼(二)
        public string d28 { get; set; } // 次診斷代碼(三)
        public string d29 { get; set; } // 次診斷代碼(四)
        public string d45 { get; set; } // 主手術(處置)代碼
        public string d46 { get; set; } // 次手術(處置)代碼一
        public string d47 { get; set; } // 次手術(處置)代碼二
        public string d48 { get; set; } // 次手術(處置)代碼三
        public string d49 { get; set; } // 次手術(處置)代碼四
    }

    public class TOTFB_Pdata
    {
        public string p1 { get; set; } // 醫令序
        public string p2 { get; set; } // 醫令類別
        public string p3 { get; set; } // 醫令代碼
        public string p5 { get; set; } // 藥品用量
        public string p6 { get; set; } // (藥品)使用頻率
        public string p7 { get; set; } // 給藥途徑/作用部位
        public string p8 { get; set; } // 會診科別
        public string p14 { get; set; } // 執行時間-起
        public string p15 { get; set; } // 執行時間-訖
        public string p16 { get; set; } // 總量
    }

    //檢驗檢查每日申報
    public class LABD_Hdata
    {
        public string h1 { get; set; } // 報告類別
        public string h2 { get; set; } // 醫事機構代碼
        public string h3 { get; set; } // 醫事類別
        public string h4 { get; set; } // 執行年月
        public string h5 { get; set; } // 健保卡刷卡日期時間
        public string h6 { get; set; } // 就醫類別
        public string h7 { get; set; } // 就醫序號
        public string h8 { get; set; } // 補卡註記
        public string h9 { get; set; } // 身分證統一編號
        public string h10 { get; set; } // 出生日期
        public string h11 { get; set; } // 就醫日期
        public string h12 { get; set; } // 治療結束日期
        public string h13 { get; set; } // 入院年月日
        public string h14 { get; set; } // 出院年月日
        public string h15 { get; set; } // 醫令代碼
        public string h19 { get; set; } // 醫囑日期時間
        public string h20 { get; set; } // 採檢/實際檢查/手術日期時間
        public string h22 { get; set; } // 檢體採檢方法/來源/類別
    }

    public class LABD_Rdata
    {
        public string r1 { get; set; } // 報告序號
        public string r2 { get; set; } // 檢驗項目名稱
        public string r3 { get; set; } // 檢驗方法
        public string r4 { get; set; } // 檢驗報告結果值
        public string r5 { get; set; } // 單位
        public string r6_1 { get; set; } // 參考值下限
        public string r6_2 { get; set; } // 參考值上限
        public string r7 { get; set; } // 報告結果
        public string r8_1 { get; set; } // 病理發現及診斷
        public string r10 { get; set; } // 報告日期時間
        public string r12 { get; set; } // 檢驗（查）結果值註記
    }

    //檢驗檢查每月申報
    public class LABM_Hdata
    {
        public string h1 { get; set; } // 報告類別
        public string h2 { get; set; } // 醫事機構代碼
        public string h3 { get; set; } // 醫事類別
        public string h4 { get; set; } // 費用年月
        public string h5 { get; set; } // 申報類別
        public string h6 { get; set; } // 申報日期
        public string h7 { get; set; } // 案件分類
        public string h8 { get; set; } // 流水編號
        public string h9 { get; set; } // 身分證統一編號
        public string h10 { get; set; } // 出生日期
        public string h11 { get; set; } // 就醫日期
        public string h12 { get; set; } // 治療結束日期
        public string h13 { get; set; } // 入院年月日
        public string h14 { get; set; } // 出院年月日
        public string h17 { get; set; } // 醫令序
        public string h18 { get; set; } // 醫令代碼
        public string h22 { get; set; } // 醫囑日期時間
        public string h23 { get; set; } // 採檢/實際檢查/手術日期時間
        public string h25 { get; set; } // 檢體採檢方法/來源/類別
    }

    public class LABM_Rdata
    {
        public string r1 { get; set; } // 報告序號
        public string r2 { get; set; } // 檢驗項目名稱
        public string r3 { get; set; } // 檢驗方法
        public string r4 { get; set; } // 檢驗報告結果值
        public string r5 { get; set; } // 單位
        public string r6_1 { get; set; } // 參考值下限
        public string r6_2 { get; set; } // 參考值上限
        public string r7 { get; set; } // 報告結果
        public string r8_1 { get; set; } // 病理發現及診斷
        public string r10 { get; set; } // 報告日期時間
        public string r12 { get; set; } // 檢驗（查）結果值註記
    }

    //public class CRLF_Mdata
    //{
    //    public string LF1_1 { get; set; } // 申報醫院代碼
    //    public string LF1_2 { get; set; } // 病歷號碼(此欄位不需要)
    //    public string LF1_3 { get; set; } // 姓名(此欄位不需要)
    //    public string LF1_4 { get; set; } // 身分證統一編號
    //    public string LF1_5 { get; set; } // 性別
    //    public string LF1_6 { get; set; } // 出生日期
    //    public string LF1_7 { get; set; } // 戶籍地代碼
    //    public string LF2_1 { get; set; } // 診斷年齡
    //    public string LF2_2 { get; set; } // 癌症發生順序號碼
    //    public string LF2_3 { get; set; } // 個案分類
    //    public string LF2_3_1 { get; set; } // 診斷狀態分類
    //    public string LF2_3_2 { get; set; } // 治療狀態分類
    //    public string LF2_4 { get; set; } // 首次就診日期
    //    public string LF2_5 { get; set; } // 最初診斷日期
    //    public string LF2_6 { get; set; } // 原發部位
    //    public string LF2_7 { get; set; } // 側性
    //    public string LF2_8 { get; set; } // 組織類型
    //    public string LF2_9 { get; set; } // 性態碼
    //    public string LF2_10_1 { get; set; } // 臨床分級/分化
    //    public string LF2_10_2 { get; set; } // 病理分級/分化
    //    public string LF2_11 { get; set; } // 癌症確診方式
    //    public string LF2_12 { get; set; } // 首次顯微鏡檢證實日期
    //    public string LF2_13 { get; set; } // 腫瘤大小
    //    public string LF2_13_1 { get; set; } // 神經侵襲
    //    public string LF2_13_2 { get; set; } // 淋巴管或血管侵犯
    //    public string LF2_14 { get; set; } // 區域淋巴結檢查數目
    //    public string LF2_15 { get; set; } // 區域淋巴結侵犯數目
    //    public string LF3_1 { get; set; } // 診斷性及分期性手術處置日期
    //    public string LF3_2 { get; set; } // 外院診斷性及分期性手術處置
    //    public string LF3_3 { get; set; } // 申報醫院診斷性及分期性手術處置
    //    public string LF3_4 { get; set; } // 臨床T
    //    public string LF3_5 { get; set; } // 臨床N
    //    public string LF3_6 { get; set; } // 臨床M
    //    public string LF3_7 { get; set; } // 臨床期別組合
    //    public string LF3_8 { get; set; } // 臨床分期字根/字首
    //    public string LF3_10 { get; set; } // 病理T
    //    public string LF3_11 { get; set; } // 病理N
    //    public string LF3_12 { get; set; } // 病理M
    //    public string LF3_13 { get; set; } // 病理期別組合
    //    public string LF3_14 { get; set; } // 病理分期字根/字首
    //    public string LF3_16 { get; set; } // AJCC癌症分期版本與章節
    //    public string LF3_17 { get; set; } // 其他分期系統
    //    public string LF3_19 { get; set; } // 其他分期系統期別(臨床分期)
    //    public string LF3_21 { get; set; } // 其他分期系統期別(病理分期)
    //    public string LF4_1 { get; set; } // 首次療程開始日期
    //    public string LF4_1_1 { get; set; } // 首次手術日期
    //    public string LF4_1_2 { get; set; } // 原發部位最確切的手術切除日期
    //    public string LF4_1_3 { get; set; } // 外院原發部位手術方式
    //    public string LF4_1_4 { get; set; } // 申報醫院原發部位手術方式
    //    public string LF4_1_4_1 { get; set; } // 微創手術
    //    public string LF4_1_5 { get; set; } // 原發部位手術邊緣
    //    public string LF4_1_5_1 { get; set; } // 原發部位手術切緣距離
    //    public string LF4_1_6 { get; set; } // 外院區域淋巴結手術範圍
    //    public string LF4_1_7 { get; set; } // 申報醫院區域淋巴結手術範圍
    //    public string LF4_1_8 { get; set; } // 外院其他部位手術方式
    //    public string LF4_1_9 { get; set; } // 申報醫院其他部位手術方式
    //    public string LF4_1_10 { get; set; } // 原發部位未手術原因
    //    public string LF4_2_1_1 { get; set; } // 放射治療臨床標靶體積摘要
    //    public string LF4_2_1_2 { get; set; } // 放射治療儀器
    //    public string LF4_2_1_3 { get; set; } // 放射治療開始日期
    //    public string LF4_2_1_4 { get; set; } // 放射治療結束日期
    //    public string LF4_2_1_5 { get; set; } // 放射治療與手術順序
    //    public string LF4_2_1_6 { get; set; } // 區域治療與全身性治療順序
    //    public string LF4_2_1_8 { get; set; } // 放射治療執行狀態
    //    public string LF4_2_2_1 { get; set; } // 體外放射治療技術
    //    public string LF4_2_2_2_1 { get; set; } // 最高放射劑量臨床標靶體積
    //    public string LF4_2_2_2_2 { get; set; } // 最高放射劑量臨床標靶體積劑量
    //    public string LF4_2_2_2_3 { get; set; } // 最高放射劑量臨床標靶體積治療次數
    //    public string LF4_2_2_3_1 { get; set; } // 較低放射劑量臨床標靶體積
    //    public string LF4_2_2_3_2 { get; set; } // 最低放射劑量臨床標靶體積劑量
    //    public string LF4_2_2_3_3 { get; set; } // 最低放射劑量臨床標靶體積治療次數
    //    public string LF4_2_3_1 { get; set; } // 其他放射治療儀器
    //    public string LF4_2_3_2 { get; set; } // 其他放射治療技術
    //    public string LF4_2_3_3_1 { get; set; } // 其他放射治療臨床標靶體積
    //    public string LF4_2_3_3_2 { get; set; } // 其他放射劑量臨床標靶體積劑量
    //    public string LF4_2_3_3_3 { get; set; } // 其他放射劑量臨床標靶體積治療次數
    //    public string LF4_3_1 { get; set; } // 全身性治療開始日期
    //    public string LF4_3_2 { get; set; } // 外院化學治療
    //    public string LF4_3_3 { get; set; } // 申報醫院化學治療
    //    public string LF4_3_4 { get; set; } // 申報醫院化學治療開始日期
    //    public string LF4_3_5 { get; set; } // 外院賀爾蒙/類固醇治療
    //    public string LF4_3_6 { get; set; } // 申報醫院賀爾蒙/類固醇治療
    //    public string LF4_3_7 { get; set; } // 申報醫院賀爾蒙/類固醇治療開始日期
    //    public string LF4_3_8 { get; set; } // 外院免疫治療
    //    public string LF4_3_9 { get; set; } // 申報醫院免疫治療
    //    public string LF4_3_10 { get; set; } // 申報醫院免疫治療開始日期
    //    public string LF4_3_11 { get; set; } // 骨髓/幹細胞移植或內分泌處置
    //    public string LF4_3_12 { get; set; } // 申報醫院骨髓/幹細胞移植或內分泌處置開始日期
    //    public string LF4_3_13 { get; set; } // 外院標靶治療
    //    public string LF4_3_14 { get; set; } // 申報醫院標靶治療
    //    public string LF4_3_15 { get; set; } // 申報醫院標靶治療開始日期
    //    public string LF4_4 { get; set; } // 申報醫院緩和照護
    //    public string LF4_5_1 { get; set; } // 其他治療
    //    public string LF4_5_2 { get; set; } // 其他治療開始日期
    //    public string LF5_1 { get; set; } // 首次復發日期
    //    public string LF5_2 { get; set; } // 首次復發型式
    //    public string LF5_3 { get; set; } // 最後聯絡或死亡日期
    //    public string LF5_4 { get; set; } // 生存狀態
    //    public string LF6_1 { get; set; } // 摘錄者(此欄位沒有被寫入)
    //    public string LF7_1 { get; set; } // 身高
    //    public string LF7_2 { get; set; } // 體重
    //    public string LF7_3 { get; set; } // 吸菸行為
    //    public string LF7_4 { get; set; } // 嚼檳榔行為
    //    public string LF7_5 { get; set; } // 喝酒行為
    //    public string LF7_6 { get; set; } // 首次治療前生活功能狀態評估
    //    public string LF8_1 { get; set; } // 癌症部位特定因子1
    //    public string LF8_2 { get; set; } // 癌症部位特定因子2
    //    public string LF8_3 { get; set; } // 癌症部位特定因子3
    //    public string LF8_4 { get; set; } // 癌症部位特定因子4
    //    public string LF8_5 { get; set; } // 癌症部位特定因子5
    //    public string LF8_6 { get; set; } // 癌症部位特定因子6
    //    public string LF8_7 { get; set; } // 癌症部位特定因子7
    //    public string LF8_8 { get; set; } // 癌症部位特定因子8
    //    public string LF8_9 { get; set; } // 癌症部位特定因子9
    //    public string LF8_10 { get; set; } // 癌症部位特定因子10
    //}
    //public class CRSF_Mdata
    //{
    //    public string SF1_1 { get; set; } // 申報醫院代碼
    //    public string SF1_2 { get; set; } // 病歷號碼
    //    public string SF1_3 { get; set; } // 姓名(此欄位不需要)
    //    public string SF1_4 { get; set; } // 身分證統一編號
    //    public string SF1_5 { get; set; } // 性別
    //    public string SF1_6 { get; set; } // 出生日期
    //    public string SF1_7 { get; set; } // 戶籍地代碼
    //    public string SF2_1 { get; set; } // 診斷年齡
    //    public string SF2_2 { get; set; } // 癌症發生順序號碼
    //    public string SF2_3 { get; set; } // 個案分類
    //    public string SF2_3_1 { get; set; } // 診斷狀態分類
    //    public string SF2_3_2 { get; set; } // 治療狀態分類
    //    public string SF2_4 { get; set; } // 首次就診日期
    //    public string SF2_5 { get; set; } // 最初診斷日期
    //    public string SF2_6 { get; set; } // 原發部位
    //    public string SF2_7 { get; set; } // 側性
    //    public string SF2_8 { get; set; } // 組織類型
    //    public string SF2_9 { get; set; } // 性態碼
    //    public string SF2_10_1 { get; set; } // 臨床分級/分化
    //    public string SF2_10_2 { get; set; } // 病理分級/分化
    //    public string SF2_11 { get; set; } // 癌症確診方式
    //    public string SF2_12 { get; set; } // 首次顯微鏡檢證實日期
    //    public string SF4_1_1 { get; set; } // 首次手術日期
    //    public string SF4_1_4 { get; set; } // 申報醫院原發部位手術方式
    //    public string SF4_2_1_3 { get; set; } // 放射治療開始日期
    //    public string SF4_2_1_7 { get; set; } // 放射治療機構
    //    public string SF4_3_3 { get; set; } // 申報醫院化學治療
    //    public string SF4_3_4 { get; set; } // 申報醫院化學治療開始日期
    //    public string SF4_3_6 { get; set; } // 申報醫院賀爾蒙/類固醇治療
    //    public string SF4_3_7 { get; set; } // 申報醫院賀爾蒙/類固醇治療開始日期
    //    public string SF4_3_9 { get; set; } // 申報醫院免疫治療
    //    public string SF4_3_10 { get; set; } // 申報醫院免疫治療開始日期
    //    public string SF4_3_11 { get; set; } // 骨髓/幹細胞移植或內分泌處置
    //    public string SF4_3_12 { get; set; } // 申報醫院骨髓/幹細胞移植或內分泌處置開始日期
    //    public string SF4_3_14 { get; set; } // 申報醫院標靶治療
    //    public string SF4_3_15 { get; set; } // 申報醫院標靶治療開始日期
    //    public string SF4_4 { get; set; } // 申報醫院緩和照護
    //    public string SF4_5_1 { get; set; } // 其他治療
    //    public string SF4_5_2 { get; set; } // 其他治療開始日期
    //    public string SF6_1 { get; set; } // 摘錄者(此欄位沒有被寫入)
    //    public string SF7_1 { get; set; } // 身高
    //    public string SF7_2 { get; set; } // 體重
    //    public string SF7_3 { get; set; } // 吸菸行為
    //    public string SF7_4 { get; set; } // 嚼檳榔行為
    //    public string SF7_5 { get; set; } // 喝酒行為
    //}
}
