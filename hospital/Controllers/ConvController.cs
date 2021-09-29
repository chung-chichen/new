using hospital.Models;
using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text.RegularExpressions;
using System.Web;
using System.Web.Mvc;
using System.Xml;
using System.Data;
using NPOI.SS.UserModel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System.Xml.XPath;
using hospital.Models.Biobank;
using System.Diagnostics;
using System.Globalization;

namespace hospital.Controllers
{
    public class ConvController : Controller
    {
        private BiobankDataDbContext db = new BiobankDataDbContext();

        public ActionResult OPD(string form, string alert, string result, string fileZip, string fileError)
        {
            ViewBag.form = form;
            ViewBag.alert = alert;
            ViewBag.result = result;
            ViewBag.fileZip = fileZip;
            ViewBag.fileError = fileError;
            return View();
        }

        /// <summary>
        /// 門診健保資料A框上傳(XML格式)
        /// A框轉B框
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult OPD(HttpPostedFileBase file, HttpPostedFileBase file_excel)
        {
            Update(file, file_excel, "OPD");
            return RedirectToAction("OPD", new { form = "OPD", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }

        public ActionResult IPD(string form, string alert, string result, string fileZip, string fileError)
        {

            ViewBag.form = form;
            ViewBag.alert = alert;
            ViewBag.result = result;
            ViewBag.fileZip = fileZip;
            ViewBag.fileError = fileError;
            return View();
        }

        /// <summary>
        /// 住院健保資料A框上傳(XML格式)
        /// A框轉B框
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult IPD(HttpPostedFileBase file, HttpPostedFileBase file_excel)
        {
            Update(file, file_excel, "IPD");
            return RedirectToAction("IPD", new { form = "IPD", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }

        public ActionResult LAB(string form, string alert, string result, string fileZip, string fileError)
        {
            ViewBag.form = form;
            ViewBag.alert = alert;
            ViewBag.result = result;
            ViewBag.fileZip = fileZip;
            ViewBag.fileError = fileError;
            return View();
        }

        /// <summary>
        /// 每日檢驗檢查資料A框上傳(XML格式)
        /// A框轉B框
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult LAB_D(HttpPostedFileBase file, HttpPostedFileBase file_excel)
        {
            Update(file, file_excel, "LAB_D");
            return RedirectToAction("LAB", new { form = "LAB_D", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }

        /// <summary>
        /// 每月檢驗檢查資料A框上傳(XML格式)
        /// A框轉B框
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        [HttpPost]
        public ActionResult LAB_M(HttpPostedFileBase file, HttpPostedFileBase file_excel)
        {
            Update(file, file_excel, "LAB_M");
            return RedirectToAction("LAB", new { form = "LAB_M", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }

        /// <summary>
        /// CRLF(癌登長表)資料B框格式檢查(xlxs, csv)
        /// </summary>
        /// <param name="file"></param>
        /// <returns></returns>
        public ActionResult CRLF(string form, string alert, string result, string fileZip, string fileError)
        {
            ViewBag.form = form;
            ViewBag.alert = alert;
            ViewBag.result = result;
            ViewBag.fileZip = fileZip;
            ViewBag.fileError = fileError;
            return View();
        }

        [HttpPost]
        public ActionResult CRLF_TXT(HttpPostedFileBase file, HttpPostedFileBase file_excel)
        {
            Update_excel(file, file_excel, "CRLF_TXT");
            return RedirectToAction("CRLF", new { form = "CRLF_TXT", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }

        public ActionResult CRSF(string form, string alert, string result, string fileZip, string fileError)
        {
            ViewBag.form = form;
            ViewBag.alert = alert;
            ViewBag.result = result;
            ViewBag.fileZip = fileZip;
            ViewBag.fileError = fileError;
            return View();
        }

        [HttpPost]
        public ActionResult CRSF_TXT(HttpPostedFileBase file, HttpPostedFileBase file_excel)
        {
            Update_excel(file, file_excel, "CRSF_TXT");
            return RedirectToAction("CRLF", new { form = "CRSF_TXT", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }


        // ---------------------------------------------------------------------------------------------------------

        // 解析XML檔案
        // 儲存內容陣列
        TOTFA_Tdata TOTFA_tdata = new TOTFA_Tdata();
        TOTFA_Dhead TOTFA_dhead = new TOTFA_Dhead();
        TOTFA_Dbody TOTFA_dbody = new TOTFA_Dbody();
        TOTFA_Pdata TOTFA_pdata = new TOTFA_Pdata();

        TOTFB_Tdata TOTFB_tdata = new TOTFB_Tdata();
        TOTFB_Dhead TOTFB_dhead = new TOTFB_Dhead();
        TOTFB_Dbody TOTFB_dbody = new TOTFB_Dbody();
        TOTFB_Pdata TOTFB_pdata = new TOTFB_Pdata();

        LABM_Hdata LABM_hdata = new LABM_Hdata();
        LABM_Rdata LABM_rdata = new LABM_Rdata();

        LABD_Hdata LABD_hdata = new LABD_Hdata();
        LABD_Rdata LABD_rdata = new LABD_Rdata();


        // 儲存所有資料
        List<TOTFA_Tdata> TOTFAtdata_list = new List<TOTFA_Tdata>();
        List<TOTFA_Dhead> TOTFAdhead_list = new List<TOTFA_Dhead>();
        List<TOTFA_Dbody> TOTFAdbody_list = new List<TOTFA_Dbody>();

        List<TOTFA_Tdata> TOTFAOtdata_list = new List<TOTFA_Tdata>();
        List<TOTFA_Dhead> TOTFAOdhead_list = new List<TOTFA_Dhead>();
        List<TOTFA_Pdata> TOTFAOpdata_list = new List<TOTFA_Pdata>();

        List<TOTFB_Tdata> TOTFBtdata_list = new List<TOTFB_Tdata>();
        List<TOTFB_Dhead> TOTFBdhead_list = new List<TOTFB_Dhead>();
        List<TOTFB_Dbody> TOTFBdbody_list = new List<TOTFB_Dbody>();

        List<TOTFB_Tdata> TOTFBOtdata_list = new List<TOTFB_Tdata>();
        List<TOTFB_Dhead> TOTFBOdhead_list = new List<TOTFB_Dhead>();
        List<TOTFB_Pdata> TOTFBOpdata_list = new List<TOTFB_Pdata>();

        List<LABD_Hdata> LABDhdata_list = new List<LABD_Hdata>();
        List<LABD_Rdata> LABDrdata_list = new List<LABD_Rdata>();

        List<LABM_Hdata> LABMhdata_list = new List<LABM_Hdata>();
        List<LABM_Rdata> LABMrdata_list = new List<LABM_Rdata>();

        //篩選清單 身分證
        Patient_Hash_Guid patient_Hash = new Patient_Hash_Guid();
        List<Patient_Hash_Guid> patient_Hash_Guids = new List<Patient_Hash_Guid>();


        // 變數
        string alert = "0"; // 提醒 0為錯誤 1為成功
        string result = ""; // 上傳結果訊息
        string data = ""; // 轉檔結果訊息
        string fileZip = ""; // ZIP檔名
        string fileError = ""; // Error檔名

        // ---------------------------------------------------------------------------------------------------------

        // 上傳檔案
        public void Update(HttpPostedFileBase file, HttpPostedFileBase file_excel, string type)
        {
            alert = "0";
            result = "";
            data = "";
            fileZip = "";
            fileError = "";
            var type_name = type;
            if (file != null && file_excel != null)
            {
                if (file.ContentLength > 0 && file_excel.ContentLength > 0)
                {
                    // 取得副檔名
                    string extension = Path.GetExtension(file.FileName);
                    string extension_excel = Path.GetExtension(file_excel.FileName);
                    if (extension.Equals(".xml", StringComparison.OrdinalIgnoreCase) &&
                        (extension_excel.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension_excel.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
                    {
                        // 取得檔案名稱
                        var fileName = Path.GetFileName(file.FileName);
                        var fileName_excel = Path.GetFileName(file_excel.FileName);

                        // server路徑
                        var path = Server.MapPath("~/data_updata");
                        // 若資料夾不存在則建立
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                        // 檔案重新命名並儲存至指定路徑
                        DateTime myDate = DateTime.Now;
                        fileName = myDate.ToString("yyyyMMddHHmmss") + "_" + fileName;
                        path = Path.Combine(path, fileName);
                        file.SaveAs(path);

                        // 取得XML之節點
                        XmlDocument doc = new XmlDocument();
                        doc.Load(path);
                        XmlNodeList nodes = doc.DocumentElement.SelectNodes("/outpatient"); ;
                        if (type_name == "OPD")
                        {
                            nodes = doc.DocumentElement.SelectNodes("/outpatient");
                        }
                        else if (type_name == "IPD")
                        {
                            nodes = doc.DocumentElement.SelectNodes("/inpatient");
                        }
                        else if (type_name == "LAB_D")
                        {
                            nodes = doc.DocumentElement.SelectNodes("/patient");
                        }
                        else if (type_name == "LAB_M")
                        {
                            nodes = doc.DocumentElement.SelectNodes("/patient");
                        }

                        var path_excel = Server.MapPath("~/ID_Upload/");
                        if (!Directory.Exists(path_excel))
                        {
                            Directory.CreateDirectory(path_excel);
                        }
                        fileName_excel = myDate.ToString("yyyyMMddHHmmss") + "_" + fileName_excel;
                        var path_upload_excel = Path.Combine(path_excel, fileName_excel);
                        file_excel.SaveAs(path_upload_excel);

                        IWorkbook workbook;
                        string filepath_excel = Server.MapPath("~/ID_Upload/" + fileName_excel);

                        using (FileStream fileStream = new FileStream(filepath_excel, FileMode.Open, FileAccess.Read))
                        {
                            if (extension_excel.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook = new HSSFWorkbook(fileStream);
                            }
                            else if (extension_excel.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook = new XSSFWorkbook(fileStream);
                            }
                            else
                            {
                                workbook = null;
                            }
                        }

                        if (workbook != null)
                        {
                            fileZip = "1";
                            data = GUID_Converter(workbook, fileName_excel);

                            if (data == "轉檔成功")
                            {
                                // 判斷XML是否有資料
                                if (nodes.Count > 0)
                                {
                                    if (type_name == "OPD")
                                    {
                                        data = TOTFA_EXCELtransfer(path, fileName);
                                    }
                                    else if (type_name == "IPD")
                                    {
                                        data = TOTFB_EXCELtransfer(path, fileName);
                                    }
                                    else if (type_name == "LAB_D")
                                    {
                                        data = LABD_EXCELtransfer(path, fileName);
                                    }
                                    else if (type_name == "LAB_M")
                                    {
                                        data = LABM_EXCELtransfer(path, fileName);
                                    }

                                    if (data == "轉檔成功")
                                    {
                                        alert = "1";
                                        result = "XML檔案已轉檔成功！";
                                    }
                                    else
                                    {
                                        alert = "0";
                                        result = "上傳XML檔案內容錯誤！";
                                    }
                                }
                                else
                                {
                                    alert = "0";
                                    result = "上傳XML檔案內容錯誤！";
                                }
                            }
                            else
                            {
                                alert = "2";
                                result = "上傳\"篩選清單\"檔案內容錯誤！";
                            }
                        }
                        else
                        {
                            alert = "0";
                            result = "上傳\"篩選清單\"檔案內容錯誤！";
                        }
                    }
                    else if (!extension.Equals(".xml", StringComparison.OrdinalIgnoreCase))
                    {
                        alert = "0";
                        result = "只能接受XML檔案！";
                    }
                    else if (!(extension_excel.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension_excel.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
                    {
                        alert = "0";
                        result = "\"篩選清單\"只能接受EXCEL檔案！";
                    }
                }
                else
                {
                    alert = "0";
                    result = "請上傳正確之檔案！";
                }
            }
            else
            {
                alert = "0";
                result = "請上傳檔案！";
            }
        }

        // 下載ZIP檔案
        public ActionResult DownloadZIP(string fileZip)
        {
            // 下載的檔案位置
            string filepath = Server.MapPath("~/data_excel_zip/" + fileZip);
            // 回傳檔案名稱
            return File(filepath, "application/zip", fileZip);
        }

        // 下載錯誤報告
        public ActionResult DownloadError(string fileError)
        {
            // 下載的檔案位置
            string filepath = Server.MapPath("~/data_error/" + fileError);
            // 回傳檔案名稱
            return File(filepath, "text/html", fileError);
        }

        //篩選資料
        public string GUID_Converter(IWorkbook workbook, string fileName)
        {
            //DataTable table = new DataTable();

            ISheet sheet = workbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;
            string error = "";
            var code_check_error = "";
            patient_Hash_Guids = new List<Patient_Hash_Guid>();

            for (int i = (sheet.FirstRowNum); i <= rowCount; i++)
            {
                patient_Hash = new Patient_Hash_Guid();
                IRow row = sheet.GetRow(i);
                //DataRow dataRow = table.NewRow();
                //判斷資料行是否錯誤
                bool Data_correct = true;
                if (row != null && i != 0)
                {
                    //身分證判斷
                    if (row.GetCell(0) != null)
                    {
                        if(row.GetCell(0).ToString() != "" && (row.GetCell(1) != null || row.GetCell(1).ToString() != ""))
                        {
                            patient_Hash.Patient_Id = row.GetCell(0).ToString();
                            code_check_error = Person_IDCode(row.GetCell(0).ToString());
                            if (row.GetCell(0).ToString() == "")
                            {
                                error += "第" + (i + 1) + "行 " + headerRow.GetCell(0) + " 欄位不得為空值\n";
                                Data_correct = false;
                            }
                            else if (row.GetCell(0).ToString() != "" && code_check_error != "OK")
                            {
                                error += "第" + (i + 1) + "行 " + headerRow.GetCell(0) + " 內容錯誤，" + code_check_error + "\n";
                                Data_correct = false;
                            }
                        }
                        else
                        {
                            Data_correct = false;
                        }
                    }
                    //出生年月日 不足位數補零
                    if (row.GetCell(1) != null)
                    {
                        if (row.GetCell(1).ToString() != "" && (row.GetCell(0) != null || row.GetCell(0).ToString() != ""))
                        {
                            patient_Hash.Patient_Birth = row.GetCell(1).ToString();
                            code_check_error = Date(row.GetCell(1).ToString(), 7);
                            if (row.GetCell(1).ToString() == "")
                            {
                                error += "第" + (i + 1) + "行 " + headerRow.GetCell(1) + " 欄位不得為空值\n";
                                Data_correct = false;
                            }
                            else if (row.GetCell(1).ToString() != "" && code_check_error != "OK")
                            {
                                error += "第" + (i + 1) + "行 " + headerRow.GetCell(1) + " 內容錯誤，" + code_check_error + "\n";
                                Data_correct = false;
                            }
                        }
                        else
                        {
                            Data_correct = false;
                        }
                    }
                    if (Data_correct)
                    {
                        patient_Hash_Guids.Add(patient_Hash);
                    }
                }
            }

            var path_error = Server.MapPath("~/data_error/");
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            var date = fileName.Substring(0, 15);
            fileError = fileName.Substring(0, fileName.Length - 4) + "_身分證&生日_Error.txt";

            var return_string = "";

            if (error == "")
            {
                error = "沒有錯誤";
                return_string = "轉檔成功";
            }
            else
            {
                return_string = "身份證與生日錯誤";
            }

            error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

            using (var file = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file.WriteLine(error);
            }

            return return_string;
        }

        // 資料判斷
        // text: 文字內容；len: 文字長度
        public string Person_IDCode(string text)
        {
            //以下開始解法二
            bool flag = Regex.IsMatch(text, @"^[A-Za-z]{1}[1-2]{1}[0-9]{8}$");
            //使用正規運算式判斷是否符合格式
            int[] ID = new int[11];//英文字會轉成2個數字,所以多一個空間存放變11個
            int count = 0;
            text = text.ToUpper();//把英文字轉成大寫
            if (flag == true)//如果符合格式就進入運算
            {//先把A~Z的對應值存到陣列裡，分別存進第一個跟第二個位置
                switch (text.Substring(0, 1))//取出輸入的第一個字--英文字母作為判斷
                {
                    //需要安裝System.ValueTuple
                    case "A": (ID[0], ID[1]) = (1, 0); break;//如果是A,ID[0]就放入1,ID[1]就放入0
                    case "B": (ID[0], ID[1]) = (1, 1); break;//以下以此類推
                    case "C": (ID[0], ID[1]) = (1, 2); break;
                    case "D": (ID[0], ID[1]) = (1, 3); break;
                    case "E": (ID[0], ID[1]) = (1, 4); break;
                    case "F": (ID[0], ID[1]) = (1, 5); break;
                    case "G": (ID[0], ID[1]) = (1, 6); break;
                    case "H": (ID[0], ID[1]) = (1, 7); break;
                    case "I": (ID[0], ID[1]) = (3, 4); break;
                    case "J": (ID[0], ID[1]) = (1, 8); break;
                    case "K": (ID[0], ID[1]) = (1, 9); break;
                    case "L": (ID[0], ID[1]) = (2, 0); break;
                    case "M": (ID[0], ID[1]) = (2, 1); break;
                    case "N": (ID[0], ID[1]) = (2, 2); break;
                    case "O": (ID[0], ID[1]) = (3, 5); break;
                    case "P": (ID[0], ID[1]) = (2, 3); break;
                    case "Q": (ID[0], ID[1]) = (2, 4); break;
                    case "R": (ID[0], ID[1]) = (2, 5); break;
                    case "S": (ID[0], ID[1]) = (2, 6); break;
                    case "T": (ID[0], ID[1]) = (2, 7); break;
                    case "U": (ID[0], ID[1]) = (2, 8); break;
                    case "V": (ID[0], ID[1]) = (2, 9); break;
                    case "W": (ID[0], ID[1]) = (3, 2); break;
                    case "X": (ID[0], ID[1]) = (3, 0); break;
                    case "Y": (ID[0], ID[1]) = (3, 1); break;
                    case "Z": (ID[0], ID[1]) = (3, 3); break;
                }
                for (int i = 2; i < ID.Length; i++)//把英文字後方的數字丟進ID[]裡
                {
                    ID[i] = Convert.ToInt32(text.Substring(i - 1, 1));
                }
                for (int j = 1; j < ID.Length - 1; j++)
                {
                    count += ID[j] * (10 - j);//根據公式,ID[1]*9+ID[2]*8......
                }
                count += ID[0] + ID[10];//把沒加到的第一個數加回來
                if (count % 10 == 0)//餘數是0代表正確
                {
                    return "OK";
                }
                else
                {
                    return "身份證不存在";
                }
            }
            else
            {
                return "身份證格式不正確";
            }
        }


        // TOTFA資料
        public void TOTFA_xml(string path)
        {
            XmlDocument doc = new XmlDocument();

            //如果 XML 文檔中存在註釋，讀取時會報錯
            //可以通過 XmlReaderSettings ，XmlReader 組合設置即可
            //文檔讀取完畢後，要關閉 XmlReader
            XmlReaderSettings settings = new XmlReaderSettings();

            //設置忽略註釋
            settings.IgnoreComments = true;

            //通過 XmlReader 設置規則
            XmlReader reader = XmlReader.Create(path, settings);

            //讀取 XmlReader 中緩存的 XML
            doc.Load(reader);
            //doc.Load(path);

            TOTFAtdata_list = new List<TOTFA_Tdata>();
            TOTFAdhead_list = new List<TOTFA_Dhead>();
            TOTFAdbody_list = new List<TOTFA_Dbody>();

            TOTFAOtdata_list = new List<TOTFA_Tdata>();
            TOTFAOdhead_list = new List<TOTFA_Dhead>();
            TOTFAOpdata_list = new List<TOTFA_Pdata>();

            //XmlNodeList nodes_pdatas = doc.SelectNodes("/outpatient/ddata/dbody/pdata");

            XmlNodeList nodes = doc.SelectSingleNode("outpatient").ChildNodes;
            for (int i = 0; i < nodes.Count; i++)
            {
                if (nodes[i].SelectSingleNode("t2") != null || nodes[i].SelectSingleNode("t3") != null)
                {
                    TOTFA_tdata = new TOTFA_Tdata();
                    if (nodes[i].SelectSingleNode("t2") != null)
                        TOTFA_tdata.t2 = nodes[i].SelectSingleNode("t2").InnerText.Trim();
                    if (nodes[i].SelectSingleNode("t3") != null)
                        TOTFA_tdata.t3 = nodes[i].SelectSingleNode("t3").InnerText.Trim();
                    if (nodes[i].SelectSingleNode("t5") != null)
                        TOTFA_tdata.t5 = nodes[i].SelectSingleNode("t5").InnerText.Trim();
                    if (nodes[i].SelectSingleNode("t6") != null)
                        TOTFA_tdata.t6 = nodes[i].SelectSingleNode("t6").InnerText.Trim();
                }

                XmlNodeList nodes_ddata = nodes[i].SelectNodes("dhead");

                if (nodes_ddata.Count > 0)
                {
                    TOTFA_dhead = new TOTFA_Dhead();
                    TOTFA_dbody = new TOTFA_Dbody();

                    if (nodes_ddata[0].SelectSingleNode("d1") != null)
                        TOTFA_dhead.d1 = nodes_ddata[0].SelectSingleNode("d1").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d2") != null)
                        TOTFA_dhead.d2 = nodes_ddata[0].SelectSingleNode("d2").InnerText.Trim();

                    nodes_ddata = nodes[i].SelectNodes("dbody");

                    if (nodes_ddata[0].SelectSingleNode("d4") != null)
                        TOTFA_dbody.d4 = nodes_ddata[0].SelectSingleNode("d4").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d5") != null)
                        TOTFA_dbody.d5 = nodes_ddata[0].SelectSingleNode("d5").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d6") != null)
                        TOTFA_dbody.d6 = nodes_ddata[0].SelectSingleNode("d6").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d7") != null)
                        TOTFA_dbody.d7 = nodes_ddata[0].SelectSingleNode("d7").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d8") != null)
                        TOTFA_dbody.d8 = nodes_ddata[0].SelectSingleNode("d8").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d9") != null)
                        TOTFA_dbody.d9 = nodes_ddata[0].SelectSingleNode("d9").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d10") != null)
                        TOTFA_dbody.d10 = nodes_ddata[0].SelectSingleNode("d10").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d11") != null)
                        TOTFA_dbody.d11 = nodes_ddata[0].SelectSingleNode("d11").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d3") != null)
                        TOTFA_dbody.d3 = nodes_ddata[0].SelectSingleNode("d3").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d19") != null)
                        TOTFA_dbody.d19 = nodes_ddata[0].SelectSingleNode("d19").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d20") != null)
                        TOTFA_dbody.d20 = nodes_ddata[0].SelectSingleNode("d20").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d21") != null)
                        TOTFA_dbody.d21 = nodes_ddata[0].SelectSingleNode("d21").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d22") != null)
                        TOTFA_dbody.d22 = nodes_ddata[0].SelectSingleNode("d22").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d23") != null)
                        TOTFA_dbody.d23 = nodes_ddata[0].SelectSingleNode("d23").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d24") != null)
                        TOTFA_dbody.d24 = nodes_ddata[0].SelectSingleNode("d24").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d25") != null)
                        TOTFA_dbody.d25 = nodes_ddata[0].SelectSingleNode("d25").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d26") != null)
                        TOTFA_dbody.d26 = nodes_ddata[0].SelectSingleNode("d26").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d27") != null)
                        TOTFA_dbody.d27 = nodes_ddata[0].SelectSingleNode("d27").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d28") != null)
                        TOTFA_dbody.d28 = nodes_ddata[0].SelectSingleNode("d28").InnerText.Trim();

                    TOTFAtdata_list.Add(TOTFA_tdata);
                    TOTFAdhead_list.Add(TOTFA_dhead);
                    TOTFAdbody_list.Add(TOTFA_dbody);

                    XmlNodeList nodes_pdata = nodes_ddata[0].SelectNodes("pdata");

                    for (int j = 0; j < nodes_pdata.Count; j++)
                    {
                        TOTFA_pdata = new TOTFA_Pdata();

                        if (nodes_pdata[j].SelectSingleNode("p1") != null)
                            TOTFA_pdata.p1 = nodes_pdata[j].SelectSingleNode("p1").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p2") != null)
                            TOTFA_pdata.p2 = nodes_pdata[j].SelectSingleNode("p2").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p3") != null)
                            TOTFA_pdata.p3 = nodes_pdata[j].SelectSingleNode("p3").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p4") != null)
                            TOTFA_pdata.p4 = nodes_pdata[j].SelectSingleNode("p4").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p5") != null)
                            TOTFA_pdata.p5 = nodes_pdata[j].SelectSingleNode("p5").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p6") != null)
                            TOTFA_pdata.p6 = nodes_pdata[j].SelectSingleNode("p6").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p7") != null)
                            TOTFA_pdata.p7 = nodes_pdata[j].SelectSingleNode("p7").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p9") != null)
                            TOTFA_pdata.p9 = nodes_pdata[j].SelectSingleNode("p9").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p10") != null)
                            TOTFA_pdata.p10 = nodes_pdata[j].SelectSingleNode("p10").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p13") != null)
                            TOTFA_pdata.p13 = nodes_pdata[j].SelectSingleNode("p13").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p14") != null)
                            TOTFA_pdata.p14 = nodes_pdata[j].SelectSingleNode("p14").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p15") != null)
                            TOTFA_pdata.p15 = nodes_pdata[j].SelectSingleNode("p15").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p17") != null)
                            TOTFA_pdata.p17 = nodes_pdata[j].SelectSingleNode("p17").InnerText.Trim();

                        TOTFAOtdata_list.Add(TOTFA_tdata);
                        TOTFAOdhead_list.Add(TOTFA_dhead);
                        TOTFAOpdata_list.Add(TOTFA_pdata);
                    }
                }
            }
            //關閉 XmlReader
            reader.Close();
        }

        // TOTFA錯誤檢查
        public string TOTFA_error(string path, string XMLfileName)
        {
            string error = "";
            var code_check_error = "";
            int line_count = 0;
            int ddata_count = 0;
            int pdata_count = 0;

            // 檢查內容
            using (var file_read = new StreamReader(path))
            {
                while (!file_read.EndOfStream)
                {
                    string line = file_read.ReadLine();
                    line_count++;
                    Debug.WriteLine("Line Count:" + line_count);

                    if (line.IndexOf("</ddata>") != -1)
                        ddata_count++;
                    else if (line.IndexOf("</pdata>") != -1)
                        pdata_count++;

                    if (line.IndexOf("<t2>") != -1)
                    {
                        code_check_error = Code(TOTFAtdata_list[ddata_count].t2, 10, 2);
                        if (TOTFAtdata_list[ddata_count].t2 == "")
                            error += "第" + line_count + "行 t2 欄位不得為空值\n";
                        else if (TOTFAtdata_list[ddata_count].t2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<t3>") != -1)
                    {
                        code_check_error = Date(TOTFAtdata_list[ddata_count].t3, 5);
                        if (TOTFAtdata_list[ddata_count].t3 == "")
                            error += "第" + line_count + "行 t3 欄位不得為空值\n";
                        else if (TOTFAtdata_list[ddata_count].t3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<t5>") != -1)
                    {
                        code_check_error = Check(TOTFAtdata_list[ddata_count].t5, 1, 2);
                        if (TOTFAtdata_list[ddata_count].t5 == "")
                            error += "第" + line_count + "行 t5 欄位不得為空值\n";
                        else if (TOTFAtdata_list[ddata_count].t5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<t6>") != -1)
                    {
                        code_check_error = Date(TOTFAtdata_list[ddata_count].t6, 7);
                        if (TOTFAtdata_list[ddata_count].t6 == "")
                            error += "第" + line_count + "行 t6 欄位不得為空值\n";
                        else if (TOTFAtdata_list[ddata_count].t6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d1>") != -1)
                    {
                        code_check_error = Code(TOTFAdhead_list[ddata_count].d1, 2, 2);
                        code_check_error = Rule_ad1(TOTFAdhead_list[ddata_count].d1);
                        if (TOTFAdhead_list[ddata_count].d1 == "")
                            error += "第" + line_count + "行 d1 欄位不得為空值\n";
                        else if (TOTFAdhead_list[ddata_count].d1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d2>") != -1)
                    {
                        code_check_error = Code(TOTFAdhead_list[ddata_count].d2, 6, 5);
                        if (TOTFAdhead_list[ddata_count].d2 == "")
                            error += "第" + line_count + "行 d2 欄位不得為空值\n";
                        else if (TOTFAdhead_list[ddata_count].d2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d4>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d4, 2, 2);
                        code_check_error = Rule_ad4(TOTFAdbody_list[ddata_count].d4);
                        if (TOTFAdbody_list[ddata_count].d4 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d4 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d5>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d5, 2, 2);
                        code_check_error = Rule_ad4(TOTFAdbody_list[ddata_count].d5);
                        if (TOTFAdbody_list[ddata_count].d5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d6>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d6, 2, 2);
                        code_check_error = Rule_ad4(TOTFAdbody_list[ddata_count].d6);
                        if (TOTFAdbody_list[ddata_count].d6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d7>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d7, 2, 2);
                        code_check_error = Rule_ad4(TOTFAdbody_list[ddata_count].d7);
                        if (TOTFAdbody_list[ddata_count].d7 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d7 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d8>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d8, 2, 2);
                        code_check_error = Rule_ad8(TOTFAdbody_list[ddata_count].d8);
                        if (TOTFAdbody_list[ddata_count].d8 == "")
                            error += "第" + line_count + "行 d8 欄位不得為空值\n";
                        else if (TOTFAdbody_list[ddata_count].d8 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d8 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d9>") != -1)
                    {
                        code_check_error = Date(TOTFAdbody_list[ddata_count].d9, 7);
                        if (TOTFAdbody_list[ddata_count].d9 == "")
                            error += "第" + line_count + "行 d9 欄位不得為空值\n";
                        else if (TOTFAdbody_list[ddata_count].d9 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d9 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d10>") != -1)
                    {
                        code_check_error = Date(TOTFAdbody_list[ddata_count].d10, 7);
                        if ((TOTFAdhead_list[ddata_count].d1 == "08" || TOTFAdhead_list[ddata_count].d1 == "28") && TOTFAdbody_list[ddata_count].d10 == "")
                            error += "第" + line_count + "行 d10 欄位不得為空值\n";
                        else if (TOTFAdbody_list[ddata_count].d10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d10 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d11>") != -1)
                    {
                        code_check_error = Date(TOTFAdbody_list[ddata_count].d11, 7);
                        if (code_check_error == "OK")
                            code_check_error = Birth(TOTFAdbody_list[ddata_count].d11, TOTFAdbody_list[ddata_count].d9);
                        if (TOTFAdbody_list[ddata_count].d11 == "")
                            error += "第" + line_count + "行 d11 欄位不得為空值\n";
                        else if (TOTFAdbody_list[ddata_count].d11 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d11 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d3>") != -1)
                    {
                        code_check_error = Idcard(TOTFAdbody_list[ddata_count].d3, 10);
                        if (TOTFAdbody_list[ddata_count].d3 == "")
                            error += "第" + line_count + "行 d3 欄位不得為空值\n";
                        else if (TOTFAdbody_list[ddata_count].d3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d19>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d19, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d19 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d19 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d20>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d20, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d20 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d20 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d21>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d21, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d21 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d21 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d22>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d22, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d22 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d22 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d23>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d23, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d23 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d23 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d24>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d24, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d24 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d24 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d25>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d25, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d25 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d25 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d26>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d26, 9, 3);
                        if (TOTFAdbody_list[ddata_count].d26 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d26 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d27>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d27, 2, 5);
                        if (TOTFAdbody_list[ddata_count].d27 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d27 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d28>") != -1)
                    {
                        code_check_error = Code(TOTFAdbody_list[ddata_count].d28, 1, 2);
                        if (TOTFAdbody_list[ddata_count].d28 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d28 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p1>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p1, 2, 5);
                        if ((TOTFAOpdata_list[pdata_count].p3 == "1" || TOTFAOpdata_list[pdata_count].p3 == "4") && TOTFAOpdata_list[pdata_count].p1 == "")
                            error += "第" + line_count + "行 p1 欄位不得為空值\n";
                        else if (TOTFAOpdata_list[pdata_count].p1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p2>") != -1)
                    {
                        //如果p2有值才驗證
                        if (TOTFAOpdata_list[pdata_count].p2 != "")
                            code_check_error = Check(TOTFAOpdata_list[pdata_count].p2, 0, 5);
                        if ((TOTFAOpdata_list[pdata_count].p3 == "1" || TOTFAOpdata_list[pdata_count].p3 == "2") && TOTFAOpdata_list[pdata_count].p2 == "")
                            error += "第" + line_count + "行 p2 欄位不得為空值，「d1 案件分類」為1用藥明細、2診療明細(檢驗(查)或物理治療者)，本欄為必填欄位\n";
                        //根據"義大醫院0429核對後系統待修正之項目"
                        //else if (TOTFAOpdata_list[pdata_count].p3 == "3" && TOTFAOpdata_list[pdata_count].p2 != "0")
                        //    error += "第" + line_count + "行 p2 內容錯誤，「p3 醫令類別」為3 特殊材料(屬特殊材料且未交付調劑)，本欄請填0\n";
                        else if (TOTFAOpdata_list[pdata_count].p3 == "4" && TOTFAOpdata_list[pdata_count].p2 != "1")
                            error += "第" + line_count + "行 p2 內容錯誤，「p3 醫令類別」為4 不計價(特殊材料且交付調劑者)，本欄請填1\n";
                        else if (TOTFAOpdata_list[pdata_count].p2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p3>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p3, 1, 2);
                        code_check_error = Rule_ap3(TOTFAOpdata_list[pdata_count].p3);
                        if (TOTFAOpdata_list[pdata_count].p3 == "")
                            error += "第" + line_count + "行 p3 欄位不得為空值\n";
                        else if ((TOTFAOpdata_list[pdata_count].p4 == "R001" || TOTFAOpdata_list[pdata_count].p4 == "R002" || TOTFAOpdata_list[pdata_count].p4 == "R003" || TOTFAOpdata_list[pdata_count].p4 == "R004") && TOTFAOpdata_list[pdata_count].p3 != "G")
                            error += "第" + line_count + "行 p3 內容錯誤，「p4 藥品(項目)代號」為R001、R002、R003、R004，本欄請填G\n";
                        else if (TOTFAOpdata_list[pdata_count].p3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p4>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p4, 12, 3);
                        if (TOTFAOpdata_list[pdata_count].p4 == "")
                            error += "第" + line_count + "行 p4 欄位不得為空值\n";
                        else if (TOTFAOpdata_list[pdata_count].p4 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p4 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p5>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p5, 7, 4);
                        if (TOTFAOpdata_list[pdata_count].p5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p6>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p6, 18, 3);
                        if (TOTFAOpdata_list[pdata_count].p6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p7>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p7, 18, 7);
                        if (TOTFAOpdata_list[pdata_count].p7 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p7 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p9>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p9, 4, 3);
                        if (TOTFAOpdata_list[pdata_count].p9 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p9 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p10>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p10, 7, 4);
                        if (TOTFAOpdata_list[pdata_count].p10 == "")
                            error += "第" + line_count + "行 p10 欄位不得為空值\n";
                        else if ((TOTFAOpdata_list[pdata_count].p4 == "R001" || TOTFAOpdata_list[pdata_count].p4 == "R002" || TOTFAOpdata_list[pdata_count].p4 == "R003" || TOTFAOpdata_list[pdata_count].p4 == "R004") && TOTFAOpdata_list[pdata_count].p10 != "0")
                            error += "第" + line_count + "行 p10 內容錯誤，「p4 藥品(項目)代號」為R001、R002、R003、R004，本欄請填0\n";
                        else if (TOTFAOpdata_list[pdata_count].p10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p10 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p13>") != -1)
                    {
                        code_check_error = Code(TOTFAOpdata_list[pdata_count].p13, 3, 5);
                        if (TOTFAOpdata_list[pdata_count].p13 == "")
                            error += "第" + line_count + "行 p13 欄位不得為空值\n";
                        else if (TOTFAOpdata_list[pdata_count].p13 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p13 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p14>") != -1)
                    {
                        code_check_error = Date(TOTFAOpdata_list[pdata_count].p14, 11);
                        if (TOTFAOpdata_list[pdata_count].p14 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p14 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p15>") != -1)
                    {
                        code_check_error = Date(TOTFAOpdata_list[pdata_count].p15, 11);
                        if (TOTFAOpdata_list[pdata_count].p15 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p15 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p17>") != -1)
                    {
                        if (TOTFAOpdata_list[pdata_count].p17 != "")
                            code_check_error = Check(TOTFAOpdata_list[pdata_count].p17, 1, 3);
                        if ((TOTFAdhead_list[ddata_count].d1 == "08" || TOTFAdhead_list[ddata_count].d1 == "28") && TOTFAOpdata_list[pdata_count].p17 == "")
                            error += "第" + line_count + "行 d10 欄位不得為空值\n";
                        if (TOTFAOpdata_list[pdata_count].p17 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p17 內容錯誤，" + code_check_error + "\n";
                    }

                }
                file_read.Close();

                if (error == "")
                    error = "沒有錯誤";

                error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

                return error;
            }
        }

        public string TOTFA_EXCELtransfer(string path, string XMLfileName)
        {
            TOTFA_xml(path);
            string report = TOTFA_error(path, XMLfileName);
            //將身分證換乘Hash
            //Hash_Id_Birth("TOTFA");

            var path_csv = Server.MapPath("~/data_excel/");
            if (!Directory.Exists(path_csv))
            {
                Directory.CreateDirectory(path_csv);
            }

            var path_zip = Server.MapPath("~/data_excel_zip/");
            if (!Directory.Exists(path_zip))
            {
                Directory.CreateDirectory(path_zip);
            }

            var path_error = Server.MapPath("~/data_error/");
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            //string[] fileName = new string[3];
            string[] fileName = new string[2];
            var date = XMLfileName.Substring(0, 15);
            fileName[0] = date + "TOTFAE.xlsx";
            fileName[1] = date + "TOTFAO1.xlsx";
            //fileName[2] = date + "TOTFAO2.xlsx";
            fileZip = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_ZipFile.zip";
            fileError = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_xml_Error.txt";

            List<TOTFAE> TOTFAEs = new List<TOTFAE>();
            TOTFAE TOTFAE = new TOTFAE();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("TOTFAE");

                sheet.Cells[1, 1].Value = "t2";
                sheet.Cells[1, 2].Value = "t3";
                sheet.Cells[1, 3].Value = "t5";
                sheet.Cells[1, 4].Value = "t6";
                sheet.Cells[1, 5].Value = "d1";
                sheet.Cells[1, 6].Value = "d2";
                sheet.Cells[1, 7].Value = "d4";
                sheet.Cells[1, 8].Value = "d5";
                sheet.Cells[1, 9].Value = "d6";
                sheet.Cells[1, 10].Value = "d7";
                sheet.Cells[1, 11].Value = "d8";
                sheet.Cells[1, 12].Value = "d9";
                sheet.Cells[1, 13].Value = "d10";
                sheet.Cells[1, 14].Value = "d11";
                sheet.Cells[1, 15].Value = "d3";
                sheet.Cells[1, 16].Value = "d19";
                sheet.Cells[1, 17].Value = "d20";
                sheet.Cells[1, 18].Value = "d21";
                sheet.Cells[1, 19].Value = "d22";
                sheet.Cells[1, 20].Value = "d23";
                sheet.Cells[1, 21].Value = "d24";
                sheet.Cells[1, 22].Value = "d25";
                sheet.Cells[1, 23].Value = "d26";
                sheet.Cells[1, 24].Value = "d27";
                sheet.Cells[1, 25].Value = "d28";

                var cell_number = 0;

                for (int i = 0; i < TOTFAtdata_list.Count; i++)
                {
                    patient_Hash = new Patient_Hash_Guid()
                    {
                        Patient_Id = TOTFAdbody_list[i].d3,
                        Patient_Birth = TOTFAdbody_list[i].d11
                    };
                    var id_exist = patient_Hash_Guids.Find(a => a.Patient_Id == patient_Hash.Patient_Id && a.Patient_Birth == patient_Hash.Patient_Birth);
                    if (id_exist != null)
                    {
                        TOTFAE = new TOTFAE()
                        {
                            t2 = TOTFAtdata_list[i].t2,
                            t3 = TOTFAtdata_list[i].t3,
                            t5 = TOTFAtdata_list[i].t5,
                            t6 = TOTFAtdata_list[i].t6,
                            d1 = TOTFAdhead_list[i].d1,
                            d2 = TOTFAdhead_list[i].d2,
                        };
                        sheet.Cells[(i + 2 + cell_number), 1].Value = TOTFAtdata_list[i].t2;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = TOTFAtdata_list[i].t3;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = TOTFAtdata_list[i].t5;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = TOTFAtdata_list[i].t6;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = TOTFAdhead_list[i].d1;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = TOTFAdhead_list[i].d2;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = TOTFAdbody_list[i].d4;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = TOTFAdbody_list[i].d5;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = TOTFAdbody_list[i].d6;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = TOTFAdbody_list[i].d7;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = TOTFAdbody_list[i].d8;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = TOTFAdbody_list[i].d9;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = TOTFAdbody_list[i].d10;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = TOTFAdbody_list[i].d11;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = TOTFAdbody_list[i].d3;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = TOTFAdbody_list[i].d19;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = TOTFAdbody_list[i].d20;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = TOTFAdbody_list[i].d21;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = TOTFAdbody_list[i].d22;
                        sheet.Cells[(i + 2 + cell_number), 20].Value = TOTFAdbody_list[i].d23;
                        sheet.Cells[(i + 2 + cell_number), 21].Value = TOTFAdbody_list[i].d24;
                        sheet.Cells[(i + 2 + cell_number), 22].Value = TOTFAdbody_list[i].d25;
                        sheet.Cells[(i + 2 + cell_number), 23].Value = TOTFAdbody_list[i].d26;
                        sheet.Cells[(i + 2 + cell_number), 24].Value = TOTFAdbody_list[i].d27;
                        sheet.Cells[(i + 2 + cell_number), 25].Value = TOTFAdbody_list[i].d28;
                        TOTFAEs.Add(TOTFAE);
                    }
                    else
                    {
                        cell_number--;
                    }
                }
                p.SaveAs(new FileInfo(path_csv + fileName[0]));
            }

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("TOTFAO1");

                sheet.Cells[1, 1].Value = "t2";
                sheet.Cells[1, 2].Value = "t3";
                sheet.Cells[1, 3].Value = "t5";
                sheet.Cells[1, 4].Value = "t6";
                sheet.Cells[1, 5].Value = "d1";
                sheet.Cells[1, 6].Value = "d2";
                sheet.Cells[1, 7].Value = "p1";
                sheet.Cells[1, 8].Value = "p2";
                sheet.Cells[1, 9].Value = "p3";
                sheet.Cells[1, 10].Value = "p4";
                sheet.Cells[1, 11].Value = "p5";
                sheet.Cells[1, 12].Value = "p6";
                sheet.Cells[1, 13].Value = "p7";
                sheet.Cells[1, 14].Value = "p9";
                sheet.Cells[1, 15].Value = "p10";
                sheet.Cells[1, 16].Value = "p13";
                sheet.Cells[1, 17].Value = "p14";
                sheet.Cells[1, 18].Value = "p15";
                sheet.Cells[1, 19].Value = "p17";

                var cell_number = 0;

                for (int i = 0; i < TOTFAOtdata_list.Count; i++)
                {

                    TOTFAE = new TOTFAE()
                    {
                        t2 = TOTFAOtdata_list[i].t2,
                        t3 = TOTFAOtdata_list[i].t3,
                        t5 = TOTFAOtdata_list[i].t5,
                        t6 = TOTFAOtdata_list[i].t6,
                        d1 = TOTFAOdhead_list[i].d1,
                        d2 = TOTFAOdhead_list[i].d2,
                    };
                    var to_exits = TOTFAEs.Find(f => f.t2 == TOTFAE.t2 && f.t3 == TOTFAE.t3 && f.t5 == TOTFAE.t5 && f.t6 == TOTFAE.t6
                                         && f.d1 == TOTFAE.d1 && f.d2 == TOTFAE.d2);
                    if (to_exits != null)
                    {
                        sheet.Cells[(i + 2 + cell_number), 1].Value = TOTFAOtdata_list[i].t2;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = TOTFAOtdata_list[i].t3;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = TOTFAOtdata_list[i].t5;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = TOTFAOtdata_list[i].t6;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = TOTFAOdhead_list[i].d1;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = TOTFAOdhead_list[i].d2;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = TOTFAOpdata_list[i].p1;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = TOTFAOpdata_list[i].p2;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = TOTFAOpdata_list[i].p3;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = TOTFAOpdata_list[i].p4;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = TOTFAOpdata_list[i].p5;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = TOTFAOpdata_list[i].p6;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = TOTFAOpdata_list[i].p7;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = TOTFAOpdata_list[i].p9;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = TOTFAOpdata_list[i].p10;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = TOTFAOpdata_list[i].p13;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = TOTFAOpdata_list[i].p14;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = TOTFAOpdata_list[i].p15;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = TOTFAOpdata_list[i].p17;
                    }
                    else
                    {
                        cell_number--;
                    }
                }

                p.SaveAs(new FileInfo(path_csv + fileName[1]));
            }

            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    for (int i = 0; i < fileName.Length; i++)
                        readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + fileName[i], fileName[i]);
                }
            }

            using (var file_error = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file_error.WriteLine(report);
            }

            return "轉檔成功";
        }

        // TOTFB資料
        public void TOTFB_xml(string path)
        {
            XmlDocument doc = new XmlDocument();

            //如果 XML 文檔中存在註釋，讀取時會報錯
            //可以通過 XmlReaderSettings ，XmlReader 組合設置即可
            //文檔讀取完畢後，要關閉 XmlReader
            XmlReaderSettings settings = new XmlReaderSettings();

            //設置忽略註釋
            settings.IgnoreComments = true;

            //通過 XmlReader 設置規則
            XmlReader reader = XmlReader.Create(path, settings);

            //讀取 XmlReader 中緩存的 XML
            doc.Load(reader);
            //doc.Load(path);

            TOTFBtdata_list = new List<TOTFB_Tdata>();
            TOTFBdhead_list = new List<TOTFB_Dhead>();
            TOTFBdbody_list = new List<TOTFB_Dbody>();

            TOTFBOtdata_list = new List<TOTFB_Tdata>();
            TOTFBOdhead_list = new List<TOTFB_Dhead>();
            TOTFBOpdata_list = new List<TOTFB_Pdata>();

            //XmlNodeList nodes = doc.DocumentElement.SelectNodes("/inpatient/tdata");
            XmlNodeList nodes = doc.SelectSingleNode("inpatient").ChildNodes;

            for (int i = 0; i < nodes.Count; i++)
            {
                if (nodes[i].SelectSingleNode("t2") != null || nodes[i].SelectSingleNode("t3") != null)
                {
                    TOTFB_tdata = new TOTFB_Tdata();
                    if (nodes[i].SelectSingleNode("t2") != null)
                        TOTFB_tdata.t2 = nodes[i].SelectSingleNode("t2").InnerText.Trim();
                    if (nodes[i].SelectSingleNode("t3") != null)
                        TOTFB_tdata.t3 = nodes[i].SelectSingleNode("t3").InnerText.Trim();
                    if (nodes[i].SelectSingleNode("t5") != null)
                        TOTFB_tdata.t5 = nodes[i].SelectSingleNode("t5").InnerText.Trim();
                    if (nodes[i].SelectSingleNode("t6") != null)
                        TOTFB_tdata.t6 = nodes[i].SelectSingleNode("t6").InnerText.Trim();
                }

                XmlNodeList nodes_ddata = nodes[i].SelectNodes("dhead");

                if (nodes_ddata.Count > 0)
                {
                    TOTFB_dhead = new TOTFB_Dhead();
                    TOTFB_dbody = new TOTFB_Dbody();

                    if (nodes_ddata[0].SelectSingleNode("d1") != null)
                        TOTFB_dhead.d1 = nodes_ddata[0].SelectSingleNode("d1").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d2") != null)
                        TOTFB_dhead.d2 = nodes_ddata[0].SelectSingleNode("d2").InnerText.Trim();

                    nodes_ddata = nodes[i].SelectNodes("dbody");

                    if (nodes_ddata[0].SelectSingleNode("d3") != null)
                        TOTFB_dbody.d3 = nodes_ddata[0].SelectSingleNode("d3").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d6") != null)
                        TOTFB_dbody.d6 = nodes_ddata[0].SelectSingleNode("d6").InnerText.Trim();
                    //2021.4.12修改 根據北榮檔案格式做修改
                    //2021.4.13修改 依據國衛院格式 d1設計在<dhead>
                    //2021.4.15修改 依據國衛院格式 d1設計在<dhead> 因北榮資料 先暫時通融
                    if (nodes_ddata[0].SelectSingleNode("d1") != null)
                        TOTFB_dhead.d1 = nodes_ddata[0].SelectSingleNode("d1").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d9") != null)
                        TOTFB_dbody.d9 = nodes_ddata[0].SelectSingleNode("d9").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d10") != null)
                        TOTFB_dbody.d10 = nodes_ddata[0].SelectSingleNode("d10").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d11") != null)
                        TOTFB_dbody.d11 = nodes_ddata[0].SelectSingleNode("d11").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d14") != null)
                        TOTFB_dbody.d14 = nodes_ddata[0].SelectSingleNode("d14").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d15") != null)
                        TOTFB_dbody.d15 = nodes_ddata[0].SelectSingleNode("d15").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d18") != null)
                        TOTFB_dbody.d18 = nodes_ddata[0].SelectSingleNode("d18").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d21") != null)
                        TOTFB_dbody.d21 = nodes_ddata[0].SelectSingleNode("d21").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d24") != null)
                        TOTFB_dbody.d24 = nodes_ddata[0].SelectSingleNode("d24").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d25") != null)
                        TOTFB_dbody.d25 = nodes_ddata[0].SelectSingleNode("d25").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d26") != null)
                        TOTFB_dbody.d26 = nodes_ddata[0].SelectSingleNode("d26").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d27") != null)
                        TOTFB_dbody.d27 = nodes_ddata[0].SelectSingleNode("d27").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d28") != null)
                        TOTFB_dbody.d28 = nodes_ddata[0].SelectSingleNode("d28").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d29") != null)
                        TOTFB_dbody.d29 = nodes_ddata[0].SelectSingleNode("d28").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d45") != null)
                        TOTFB_dbody.d45 = nodes_ddata[0].SelectSingleNode("d45").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d46") != null)
                        TOTFB_dbody.d46 = nodes_ddata[0].SelectSingleNode("d46").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d47") != null)
                        TOTFB_dbody.d47 = nodes_ddata[0].SelectSingleNode("d47").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d48") != null)
                        TOTFB_dbody.d48 = nodes_ddata[0].SelectSingleNode("d48").InnerText.Trim();
                    if (nodes_ddata[0].SelectSingleNode("d49") != null)
                        TOTFB_dbody.d49 = nodes_ddata[0].SelectSingleNode("d49").InnerText.Trim();

                    TOTFBtdata_list.Add(TOTFB_tdata);
                    TOTFBdhead_list.Add(TOTFB_dhead);
                    TOTFBdbody_list.Add(TOTFB_dbody);

                    XmlNodeList nodes_pdata = nodes_ddata[0].SelectNodes("pdata");

                    for (int j = 0; j < nodes_pdata.Count; j++)
                    {
                        TOTFB_pdata = new TOTFB_Pdata();

                        if (nodes_pdata[j].SelectSingleNode("p1") != null)
                            TOTFB_pdata.p1 = nodes_pdata[j].SelectSingleNode("p1").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p2") != null)
                            TOTFB_pdata.p2 = nodes_pdata[j].SelectSingleNode("p2").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p3") != null)
                            TOTFB_pdata.p3 = nodes_pdata[j].SelectSingleNode("p3").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p5") != null)
                            TOTFB_pdata.p5 = nodes_pdata[j].SelectSingleNode("p5").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p6") != null)
                            TOTFB_pdata.p6 = nodes_pdata[j].SelectSingleNode("p6").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p7") != null)
                            TOTFB_pdata.p7 = nodes_pdata[j].SelectSingleNode("p7").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p8") != null)
                            TOTFB_pdata.p8 = nodes_pdata[j].SelectSingleNode("p8").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p14") != null)
                            TOTFB_pdata.p14 = nodes_pdata[j].SelectSingleNode("p14").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p15") != null)
                            TOTFB_pdata.p15 = nodes_pdata[j].SelectSingleNode("p15").InnerText.Trim();
                        if (nodes_pdata[j].SelectSingleNode("p16") != null)
                            TOTFB_pdata.p16 = nodes_pdata[j].SelectSingleNode("p16").InnerText.Trim();

                        TOTFBOtdata_list.Add(TOTFB_tdata);
                        TOTFBOdhead_list.Add(TOTFB_dhead);
                        TOTFBOpdata_list.Add(TOTFB_pdata);
                    }
                }
            }
            //關閉 XmlReader
            reader.Close();
        }

        // TOTFB錯誤檢查
        public string TOTFB_error(string path, string XMLfileName)
        {
            string error = "";
            var code_check_error = "";
            int line_count = 0;
            int ddata_count = 0;
            int pdata_count = 0;

            // 檢查內容
            using (var file_read = new StreamReader(path))
            {
                while (!file_read.EndOfStream)
                {
                    string line = file_read.ReadLine();
                    line_count++;

                    if (line.IndexOf("</ddata>") != -1)
                        ddata_count++;
                    else if (line.IndexOf("</pdata>") != -1)
                        pdata_count++;

                    if (line.IndexOf("<t2>") != -1)
                    {
                        code_check_error = Code(TOTFBtdata_list[ddata_count].t2, 10, 2);
                        if (TOTFBtdata_list[ddata_count].t2 == "")
                            error += "第" + line_count + "行 t2 欄位不得為空值\n";
                        else if (TOTFBtdata_list[ddata_count].t2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<t3>") != -1)
                    {
                        code_check_error = Date(TOTFBtdata_list[ddata_count].t3, 5);
                        if (TOTFBtdata_list[ddata_count].t3 == "")
                            error += "第" + line_count + "行 t3 欄位不得為空值\n";
                        else if (TOTFBtdata_list[ddata_count].t3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<t5>") != -1)
                    {
                        code_check_error = Check(TOTFBtdata_list[ddata_count].t5, 1, 2);
                        if (TOTFBtdata_list[ddata_count].t5 == "")
                            error += "第" + line_count + "行 t5 欄位不得為空值\n";
                        else if (TOTFBtdata_list[ddata_count].t5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<t6>") != -1)
                    {
                        code_check_error = Date(TOTFBtdata_list[ddata_count].t6, 7);
                        if (TOTFBtdata_list[ddata_count].t6 == "")
                            error += "第" + line_count + "行 t6 欄位不得為空值\n";
                        else if (TOTFBtdata_list[ddata_count].t6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 t6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d1>") != -1)
                    {
                        code_check_error = Code(TOTFBdhead_list[ddata_count].d1, 2, 3);
                        code_check_error = Rule_bd1(TOTFBdhead_list[ddata_count].d1);
                        if (TOTFBdhead_list[ddata_count].d1 == "")
                            error += "第" + line_count + "行 d1 欄位不得為空值\n";
                        else if (TOTFBdhead_list[ddata_count].d1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d2>") != -1)
                    {
                        code_check_error = Code(TOTFBdhead_list[ddata_count].d2, 6, 5);
                        if (TOTFBdhead_list[ddata_count].d2 == "")
                            error += "第" + line_count + "行 d2 欄位不得為空值\n";
                        else if (TOTFBdhead_list[ddata_count].d2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d3>") != -1)
                    {
                        code_check_error = Idcard(TOTFBdbody_list[ddata_count].d3, 10);
                        if (TOTFBdbody_list[ddata_count].d3 == "")
                            error += "第" + line_count + "行 d3 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d6>") != -1)
                    {
                        code_check_error = Date(TOTFBdbody_list[ddata_count].d6, 7);
                        if (code_check_error == "OK")
                            code_check_error = Birth(TOTFBdbody_list[ddata_count].d6, TOTFBdbody_list[ddata_count].d10);
                        if (TOTFBdbody_list[ddata_count].d6 == "")
                            error += "第" + line_count + "行 d6 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d9>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d9, 2, 2);
                        code_check_error = Rule_bd9(TOTFBdbody_list[ddata_count].d9);
                        if (TOTFBdbody_list[ddata_count].d9 == "")
                            error += "第" + line_count + "行 d9 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d9 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d9 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d10>") != -1)
                    {
                        code_check_error = Date(TOTFBdbody_list[ddata_count].d10, 7);
                        if (code_check_error == "OK")
                            code_check_error = Limit(TOTFBdbody_list[ddata_count].d10, TOTFBtdata_list[ddata_count].t3, "0");
                        if (TOTFBdbody_list[ddata_count].d10 == "")
                            error += "第" + line_count + "行 d10 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d10 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d11>") != -1)
                    {
                        code_check_error = Date(TOTFBdbody_list[ddata_count].d11, 7);
                        if (code_check_error == "OK")
                            code_check_error = Limit(TOTFBdbody_list[ddata_count].d11, TOTFBtdata_list[ddata_count].t3, TOTFBdbody_list[ddata_count].d10);
                        if (TOTFBdbody_list[ddata_count].d11 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d11內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d14>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d14, 3, 5);
                        if (TOTFBdbody_list[ddata_count].d14 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d14 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d15>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d15, 3, 5);
                        if (TOTFBdbody_list[ddata_count].d15 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d15 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d18>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d18, 5, 3);
                        if (TOTFBdhead_list[ddata_count].d1 == "5" && TOTFBdbody_list[ddata_count].d18 == "")
                            error += "第" + line_count + "行 d18 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d18 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d18 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d21>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d21, 5, 2);
                        if ((TOTFBdhead_list[ddata_count].d1 == "2" || TOTFBdhead_list[ddata_count].d1 == "A2") && TOTFBdbody_list[ddata_count].d21 == "")
                            error += "第" + line_count + "行 d21 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d21 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d21 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d24>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d24, 1, 2);
                        code_check_error = Rule_bd24(TOTFBdbody_list[ddata_count].d24);
                        if (TOTFBdbody_list[ddata_count].d24 == "")
                            error += "第" + line_count + "行 d24 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d24 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d24 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d25>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d25, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d25 == "")
                            error += "第" + line_count + "行 d25 欄位不得為空值\n";
                        else if (TOTFBdbody_list[ddata_count].d25 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d25 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d26>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d26, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d26 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d26 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d27>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d27, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d27 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d27 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d28>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d28, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d28 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d28 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d29>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d29, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d29 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d29 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d45>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d45, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d45 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d45 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d46>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d46, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d46 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d46 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d47>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d47, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d47 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d47 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d48>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d48, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d48 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d48 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<d49>") != -1)
                    {
                        code_check_error = Code(TOTFBdbody_list[ddata_count].d49, 9, 3);
                        if (TOTFBdbody_list[ddata_count].d49 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 d49 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p1>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p1, 5, 5);
                        if (TOTFBOpdata_list[pdata_count].p1 == "")
                            error += "第" + line_count + "行 p1 欄位不得為空值\n";
                        else if (TOTFBOpdata_list[pdata_count].p1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p2>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p2, 1, 2);
                        code_check_error = Rule_bp2(TOTFBOpdata_list[pdata_count].p2);
                        if (TOTFBOpdata_list[pdata_count].p2 == "")
                            error += "第" + line_count + "行 p2 欄位不得為空值\n";
                        else if (TOTFBOpdata_list[pdata_count].p2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p3>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p3, 12, 3);
                        if (TOTFBOpdata_list[pdata_count].p3 == "")
                            error += "第" + line_count + "行 p3 欄位不得為空值\n";
                        else if (TOTFBOpdata_list[pdata_count].p3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p5>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p5, 7, 4);
                        if (TOTFBOpdata_list[pdata_count].p5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p6>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p6, 18, 7);
                        if (TOTFBOpdata_list[pdata_count].p6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p7>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p7, 4, 3);
                        if (TOTFBOpdata_list[pdata_count].p7 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p7 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p8>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p8, 2, 2);
                        if (TOTFBOpdata_list[pdata_count].p8 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p8 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p14>") != -1)
                    {
                        code_check_error = Date(TOTFBOpdata_list[pdata_count].p14, 11);
                        if (TOTFBOpdata_list[pdata_count].p14 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p14 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p15>") != -1)
                    {
                        code_check_error = Date(TOTFBOpdata_list[pdata_count].p15, 11);
                        if (TOTFBOpdata_list[pdata_count].p15 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p15 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<p16>") != -1)
                    {
                        code_check_error = Code(TOTFBOpdata_list[pdata_count].p16, 7, 4);
                        if (TOTFBOpdata_list[pdata_count].p16 == "")
                            error += "第" + line_count + "行 p16 欄位不得為空值\n";
                        else if (TOTFBOpdata_list[pdata_count].p16 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 p16 內容錯誤，" + code_check_error + "\n";
                    }
                }
                file_read.Close();

                if (error == "")
                    error = "沒有錯誤";

                error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

                return error;
            }
        }

        public string TOTFB_EXCELtransfer(string path, string XMLfileName)
        {
            TOTFB_xml(path);
            string report = TOTFB_error(path, XMLfileName);
            //將身分證換乘Hash
            //Hash_Id_Birth("TOTFB");

            var path_csv = Server.MapPath("~/data_excel/");
            if (!Directory.Exists(path_csv))
            {
                Directory.CreateDirectory(path_csv);
            }

            var path_zip = Server.MapPath("~/data_excel_zip/");
            if (!Directory.Exists(path_zip))
            {
                Directory.CreateDirectory(path_zip);
            }

            var path_error = Server.MapPath("~/data_error/");
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            //string[] fileName = new string[3];
            string[] fileName = new string[2];
            var date = XMLfileName.Substring(0, 15);
            fileName[0] = date + "TOTFBE.xlsx";
            fileName[1] = date + "TOTFBO1.xlsx";
            //fileName[2] = date + "TOTFBO2.xlsx";
            fileZip = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_ZipFile.zip";
            fileError = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_xml_Error.txt";

            List<TOTFAE> TOTFAEs = new List<TOTFAE>();
            TOTFAE TOTFAE = new TOTFAE();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("TOTFBE");

                sheet.Cells[1, 1].Value = "t2";
                sheet.Cells[1, 2].Value = "t3";
                sheet.Cells[1, 3].Value = "t5";
                sheet.Cells[1, 4].Value = "t6";
                sheet.Cells[1, 5].Value = "d1";
                sheet.Cells[1, 6].Value = "d2";
                sheet.Cells[1, 7].Value = "d3";
                sheet.Cells[1, 8].Value = "d6";
                sheet.Cells[1, 9].Value = "d9";
                sheet.Cells[1, 10].Value = "d10";
                sheet.Cells[1, 11].Value = "d11";
                sheet.Cells[1, 12].Value = "d14";
                sheet.Cells[1, 13].Value = "d15";
                sheet.Cells[1, 14].Value = "d18";
                sheet.Cells[1, 15].Value = "d21";
                sheet.Cells[1, 16].Value = "d24";
                sheet.Cells[1, 17].Value = "d25";
                sheet.Cells[1, 18].Value = "d26";
                sheet.Cells[1, 19].Value = "d27";
                sheet.Cells[1, 20].Value = "d28";
                sheet.Cells[1, 21].Value = "d29";
                sheet.Cells[1, 22].Value = "d45";
                sheet.Cells[1, 23].Value = "d46";
                sheet.Cells[1, 24].Value = "d47";
                sheet.Cells[1, 25].Value = "d48";
                sheet.Cells[1, 26].Value = "d49";

                var cell_number = 0;

                for (int i = 0; i < TOTFBtdata_list.Count; i++)
                {

                    patient_Hash = new Patient_Hash_Guid()
                    {
                        Patient_Id = TOTFBdbody_list[i].d3,
                        Patient_Birth = TOTFBdbody_list[i].d6
                    };

                    var id_exist = patient_Hash_Guids.Find(a => a.Patient_Id == patient_Hash.Patient_Id && a.Patient_Birth == patient_Hash.Patient_Birth);
                    if (id_exist != null)
                    {
                        TOTFAE = new TOTFAE()
                        {
                            t2 = TOTFBtdata_list[i].t2,
                            t3 = TOTFBtdata_list[i].t3,
                            t5 = TOTFBtdata_list[i].t5,
                            t6 = TOTFBtdata_list[i].t6,
                            d1 = TOTFBdhead_list[i].d1,
                            d2 = TOTFBdhead_list[i].d2,
                        };
                        sheet.Cells[(i + 2 + cell_number), 1].Value = TOTFBtdata_list[i].t2;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = TOTFBtdata_list[i].t3;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = TOTFBtdata_list[i].t5;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = TOTFBtdata_list[i].t6;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = TOTFBdhead_list[i].d1;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = TOTFBdhead_list[i].d2;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = TOTFBdbody_list[i].d3;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = TOTFBdbody_list[i].d6;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = TOTFBdbody_list[i].d9;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = TOTFBdbody_list[i].d10;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = TOTFBdbody_list[i].d11;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = TOTFBdbody_list[i].d14;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = TOTFBdbody_list[i].d15;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = TOTFBdbody_list[i].d18;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = TOTFBdbody_list[i].d21;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = TOTFBdbody_list[i].d24;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = TOTFBdbody_list[i].d25;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = TOTFBdbody_list[i].d26;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = TOTFBdbody_list[i].d27;
                        sheet.Cells[(i + 2 + cell_number), 20].Value = TOTFBdbody_list[i].d28;
                        sheet.Cells[(i + 2 + cell_number), 21].Value = TOTFBdbody_list[i].d29;
                        sheet.Cells[(i + 2 + cell_number), 22].Value = TOTFBdbody_list[i].d45;
                        sheet.Cells[(i + 2 + cell_number), 23].Value = TOTFBdbody_list[i].d46;
                        sheet.Cells[(i + 2 + cell_number), 24].Value = TOTFBdbody_list[i].d47;
                        sheet.Cells[(i + 2 + cell_number), 25].Value = TOTFBdbody_list[i].d48;
                        sheet.Cells[(i + 2 + cell_number), 26].Value = TOTFBdbody_list[i].d49;
                        TOTFAEs.Add(TOTFAE);
                    }
                    else
                    {
                        cell_number--;
                    }

                }

                p.SaveAs(new FileInfo(path_csv + fileName[0]));
            }

            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("TOTFBO1");

                sheet.Cells[1, 1].Value = "t2";
                sheet.Cells[1, 2].Value = "t3";
                sheet.Cells[1, 3].Value = "t5";
                sheet.Cells[1, 4].Value = "t6";
                sheet.Cells[1, 5].Value = "d1";
                sheet.Cells[1, 6].Value = "d2";
                sheet.Cells[1, 7].Value = "p1";
                sheet.Cells[1, 8].Value = "p2";
                sheet.Cells[1, 9].Value = "p3";
                sheet.Cells[1, 10].Value = "p5";
                sheet.Cells[1, 11].Value = "p6";
                sheet.Cells[1, 12].Value = "p7";
                sheet.Cells[1, 13].Value = "p8";
                sheet.Cells[1, 14].Value = "p14";
                sheet.Cells[1, 15].Value = "p15";
                sheet.Cells[1, 16].Value = "p16";

                var cell_number = 0;

                for (int i = 0; i < TOTFBOtdata_list.Count; i++)
                {
                    TOTFAE = new TOTFAE()
                    {
                        t2 = TOTFBOtdata_list[i].t2,
                        t3 = TOTFBOtdata_list[i].t3,
                        t5 = TOTFBOtdata_list[i].t5,
                        t6 = TOTFBOtdata_list[i].t6,
                        d1 = TOTFBOdhead_list[i].d1,
                        d2 = TOTFBOdhead_list[i].d2,
                    };
                    var to_exits = TOTFAEs.Find(f => f.t2 == TOTFAE.t2 && f.t3 == TOTFAE.t3 && f.t5 == TOTFAE.t5 && f.t6 == TOTFAE.t6
                                         && f.d1 == TOTFAE.d1 && f.d2 == TOTFAE.d2);
                    if (to_exits != null)
                    {
                        sheet.Cells[(i + 2 + cell_number), 1].Value = TOTFBOtdata_list[i].t2;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = TOTFBOtdata_list[i].t3;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = TOTFBOtdata_list[i].t5;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = TOTFBOtdata_list[i].t6;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = TOTFBOdhead_list[i].d1;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = TOTFBOdhead_list[i].d2;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = TOTFBOpdata_list[i].p1;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = TOTFBOpdata_list[i].p2;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = TOTFBOpdata_list[i].p3;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = TOTFBOpdata_list[i].p5;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = TOTFBOpdata_list[i].p6;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = TOTFBOpdata_list[i].p7;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = TOTFBOpdata_list[i].p8;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = TOTFBOpdata_list[i].p14;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = TOTFBOpdata_list[i].p15;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = TOTFBOpdata_list[i].p16;
                    }
                    else
                    {
                        cell_number--;
                    }
                }
                p.SaveAs(new FileInfo(path_csv + fileName[1]));
            }

            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    for (int i = 0; i < fileName.Length; i++)
                        readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + fileName[i], fileName[i]);
                }
            }

            using (var file_error = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file_error.WriteLine(report);
            }

            return "轉檔成功";
        }

        // LABM資料
        public void LABD_xml(string path)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(path);

            LABDhdata_list = new List<LABD_Hdata>();
            LABDrdata_list = new List<LABD_Rdata>();

            XmlNodeList nodes = doc.DocumentElement.SelectNodes("/patient/hdata");

            for (int i = 0; i < nodes.Count; i++)
            {
                LABD_hdata = new LABD_Hdata();

                if (nodes[i].SelectSingleNode("h1") != null)
                    LABD_hdata.h1 = nodes[i].SelectSingleNode("h1").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h2") != null)
                    LABD_hdata.h2 = nodes[i].SelectSingleNode("h2").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h3") != null)
                    LABD_hdata.h3 = nodes[i].SelectSingleNode("h3").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h4") != null)
                    LABD_hdata.h4 = nodes[i].SelectSingleNode("h4").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h5") != null)
                    LABD_hdata.h5 = nodes[i].SelectSingleNode("h5").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h6") != null)
                    LABD_hdata.h6 = nodes[i].SelectSingleNode("h6").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h7") != null)
                    LABD_hdata.h7 = nodes[i].SelectSingleNode("h7").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h8") != null)
                    LABD_hdata.h8 = nodes[i].SelectSingleNode("h8").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h9") != null)
                    LABD_hdata.h9 = nodes[i].SelectSingleNode("h9").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h10") != null)
                    LABD_hdata.h10 = nodes[i].SelectSingleNode("h10").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h11") != null)
                    LABD_hdata.h11 = nodes[i].SelectSingleNode("h11").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h12") != null)
                    LABD_hdata.h12 = nodes[i].SelectSingleNode("h12").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h13") != null)
                    LABD_hdata.h13 = nodes[i].SelectSingleNode("h13").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h14") != null)
                    LABD_hdata.h14 = nodes[i].SelectSingleNode("h14").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h15") != null)
                    LABD_hdata.h15 = nodes[i].SelectSingleNode("h15").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h19") != null)
                    LABD_hdata.h19 = nodes[i].SelectSingleNode("h19").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h20") != null)
                    LABD_hdata.h20 = nodes[i].SelectSingleNode("h20").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h22") != null)
                    LABD_hdata.h22 = nodes[i].SelectSingleNode("h22").InnerText.Trim();

                XmlNodeList nodes_rdata = nodes[i].SelectNodes("rdata");
                for (int j = 0; j < nodes_rdata.Count; j++)
                {
                    LABD_rdata = new LABD_Rdata();

                    if (nodes_rdata[j].SelectSingleNode("r1") != null)
                        LABD_rdata.r1 = nodes_rdata[j].SelectSingleNode("r1").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r2") != null)
                        LABD_rdata.r2 = nodes_rdata[j].SelectSingleNode("r2").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r3") != null)
                        LABD_rdata.r3 = nodes_rdata[j].SelectSingleNode("r3").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r4") != null)
                        LABD_rdata.r4 = nodes_rdata[j].SelectSingleNode("r4").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r5") != null)
                        LABD_rdata.r5 = nodes_rdata[j].SelectSingleNode("r5").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r6-1") != null)
                        LABD_rdata.r6_1 = nodes_rdata[j].SelectSingleNode("r6-1").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r6-2") != null)
                        LABD_rdata.r6_2 = nodes_rdata[j].SelectSingleNode("r6-2").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r7") != null)
                        LABD_rdata.r7 = nodes_rdata[j].SelectSingleNode("r7").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r8-1") != null)
                        LABD_rdata.r8_1 = nodes_rdata[j].SelectSingleNode("r8-1").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r10") != null)
                        LABD_rdata.r10 = nodes_rdata[j].SelectSingleNode("r10").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r12") != null)
                        LABD_rdata.r12 = nodes_rdata[j].SelectSingleNode("r12").InnerText.Trim();

                    LABDhdata_list.Add(LABD_hdata);
                    LABDrdata_list.Add(LABD_rdata);
                }
            }
        }

        // 每月LAB錯誤檢查
        public string LABD_error(string path, string XMLfileName)
        {
            string error = "";
            var code_check_error = "";
            int line_count = 0;
            int rdata_count = 0;

            // 檢查內容
            using (var file_read = new StreamReader(path))
            {
                while (!file_read.EndOfStream)
                {
                    string line = file_read.ReadLine();
                    line_count++;

                    if (line.IndexOf("</rdata>") != -1)
                        rdata_count++;

                    if (line.IndexOf("<h1>") != -1)
                    {
                        code_check_error = Check(LABDhdata_list[rdata_count].h1, 1, 4);
                        if (LABDhdata_list[rdata_count].h1 == "")
                            error += "第" + line_count + "行 h1 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h2>") != -1)
                    {
                        code_check_error = Code(LABDhdata_list[rdata_count].h2, 10, 2);
                        if (LABDhdata_list[rdata_count].h2 == "")
                            error += "第" + line_count + "行 h2 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h3>") != -1)
                    {
                        code_check_error = Code(LABDhdata_list[rdata_count].h3, 2, 1);
                        code_check_error = Rule_h3(LABDhdata_list[rdata_count].h3);
                        if (LABDhdata_list[rdata_count].h3 == "")
                            error += "第" + line_count + "行 h3 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h4>") != -1)
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h4, 5);
                        if (LABDhdata_list[rdata_count].h4 == "")
                            error += "第" + line_count + "行 h4 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h4 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h4 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h5>") != -1)
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h5, 13);
                        if (LABDhdata_list[rdata_count].h5 == "")
                            error += "第" + line_count + "行 h5 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h6>") != -1)
                    {
                        code_check_error = Code(LABDhdata_list[rdata_count].h6, 2, 2);
                        if (LABDhdata_list[rdata_count].h6 == "")
                            error += "第" + line_count + "行 h6 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h7>") != -1)
                    {
                        code_check_error = Code(LABDhdata_list[rdata_count].h7, 4, 2);
                        if (LABDhdata_list[rdata_count].h7 == "")
                            error += "第" + line_count + "行 h7 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h7 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h7 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h8>") != -1)
                    {
                        code_check_error = Code(LABDhdata_list[rdata_count].h8, 1, 1);
                        if (LABDhdata_list[rdata_count].h8 == "")
                            error += "第" + line_count + "行 h8 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h8 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h8 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h9>") != -1)
                    {
                        code_check_error = Idcard(LABDhdata_list[rdata_count].h9, 10);
                        if (LABDhdata_list[rdata_count].h9 == "")
                            error += "第" + line_count + "行 h9 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h9 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h9 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h10>") != -1)
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h10, 7);
                        if (code_check_error == "OK")
                        {
                            //LAB需要比對 h11就醫日期 h13入院日期
                            if (LABDhdata_list[rdata_count].h11 != "")
                            {
                                code_check_error = Birth(LABDhdata_list[rdata_count].h10, LABDhdata_list[rdata_count].h11);
                            }
                            else if (LABDhdata_list[rdata_count].h13 != "")
                            {
                                code_check_error = Birth(LABDhdata_list[rdata_count].h10, LABDhdata_list[rdata_count].h13);
                            }
                        }
                        if (LABDhdata_list[rdata_count].h10 == "")
                            error += "第" + line_count + "行 h10 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h10 內容錯誤，" + code_check_error + "\n";
                    }
                    //就醫日期 入院日期 則一
                    else if (line.IndexOf("<h11>") != -1 && (LABDhdata_list[rdata_count].h13 == "" && LABDhdata_list[rdata_count].h14 == ""))
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h11, 7);
                        if (LABDhdata_list[rdata_count].h11 == "")
                            error += "第" + line_count + "行 h11 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h11 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h11 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h12>") != -1 && (LABDhdata_list[rdata_count].h13 == "" && LABDhdata_list[rdata_count].h14 == ""))
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h12, 7);
                        if ((LABDhdata_list[rdata_count].h7 == "08" || LABDhdata_list[rdata_count].h7 == "28") && LABDhdata_list[rdata_count].h12 == "")
                            error += "第" + line_count + "行 h12 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h12 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h12 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h13>") != -1 && (LABDhdata_list[rdata_count].h11 == "" && LABDhdata_list[rdata_count].h12 == ""))
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h13, 7);
                        if (code_check_error == "OK")
                            code_check_error = Limit(LABDhdata_list[rdata_count].h13, LABDhdata_list[rdata_count].h4, "0");
                        if (LABDhdata_list[rdata_count].h13 == "")
                            error += "第" + line_count + "行 h13 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h13 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h13 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h14>") != -1 && (LABDhdata_list[rdata_count].h11 == "" && LABDhdata_list[rdata_count].h12 == ""))
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h14, 7);
                        if (code_check_error == "OK")
                            code_check_error = Limit(LABDhdata_list[rdata_count].h14, LABDhdata_list[rdata_count].h4, LABDhdata_list[rdata_count].h13);
                        if (LABDhdata_list[rdata_count].h14 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h14 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h15>") != -1)
                    {
                        code_check_error = Code(LABDhdata_list[rdata_count].h15, 12, 3);
                        if (LABDhdata_list[rdata_count].h15 == "")
                            error += "第" + line_count + "行 h15 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h15 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h15 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h19>") != -1)
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h19, 11);
                        if (LABDhdata_list[rdata_count].h19 == "")
                            error += "第" + line_count + "行 h19 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h19 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h19 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h20>") != -1)
                    {
                        code_check_error = Date(LABDhdata_list[rdata_count].h20, 11);
                        if (LABDhdata_list[rdata_count].h20 == "")
                            error += "第" + line_count + "行 h20 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h20 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h20 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h22>") != -1)
                    {
                        if ((LABDhdata_list[rdata_count].h1 == "1" || LABDhdata_list[rdata_count].h1 == "3") && LABDhdata_list[rdata_count].h22 == "")
                            error += "第" + line_count + "行 h22 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h22 != "" && LABDhdata_list[rdata_count].h22.Length > 200)
                            error += "第" + line_count + "行 h22 內容錯誤，字串長度超過200位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r1>") != -1)
                    {
                        code_check_error = Code(LABDrdata_list[rdata_count].r1, 6, 4);
                        if (LABDrdata_list[rdata_count].r1 == "")
                            error += "第" + line_count + "行 r1 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 r1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<r2>") != -1)
                    {
                        if ((LABDhdata_list[rdata_count].h1 == "1" || LABDhdata_list[rdata_count].h1 == "3" || LABDhdata_list[rdata_count].h1 == "4") && LABDrdata_list[rdata_count].r2 == "")
                            error += "第" + line_count + "行 r2 欄位不得為空值\n";
                        else if (LABDhdata_list[rdata_count].h1 == "4" && int.Parse(LABDrdata_list[rdata_count].r1) == 1 && LABDrdata_list[rdata_count].r2 != "BH")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為1，本欄請填BH\n";
                        else if (LABDhdata_list[rdata_count].h1 == "4" && int.Parse(LABDrdata_list[rdata_count].r1) == 2 && LABDrdata_list[rdata_count].r2 != "BW")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為2，本欄請填BW\n";
                        else if (LABDhdata_list[rdata_count].h1 == "4" && int.Parse(LABDrdata_list[rdata_count].r1) == 3 && LABDrdata_list[rdata_count].r2 != "ASA")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為3，本欄請填ASA\n";
                        else if (LABDhdata_list[rdata_count].h1 == "4" && int.Parse(LABDrdata_list[rdata_count].r1) == 4 && LABDrdata_list[rdata_count].r2 != "Surgical Approach")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為4，本欄請填BSurgical Approach\n";
                        else if (LABDrdata_list[rdata_count].r2 != "" && LABDrdata_list[rdata_count].r2.Length > 100)
                            error += "第" + line_count + "行 r2 內容錯誤，字串長度超過100位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r3>") != -1)
                    {
                        if ((LABDhdata_list[rdata_count].h1 == "1" || LABDhdata_list[rdata_count].h1 == "3" || LABDhdata_list[rdata_count].h1 == "4") && LABDrdata_list[rdata_count].r3 == "")
                            error += "第" + line_count + "行 r3 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r3 != "" && LABDrdata_list[rdata_count].r3.Length > 100)
                            error += "第" + line_count + "行 r3 內容錯誤，字串長度超過100位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r4>") != -1)
                    {
                        if ((LABDhdata_list[rdata_count].h1 == "1" || LABDhdata_list[rdata_count].h1 == "4") && LABDrdata_list[rdata_count].r4 == "")
                            error += "第" + line_count + "行 r4 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r4 != "" && LABDrdata_list[rdata_count].r4.Length > 4000)
                            error += "第" + line_count + "行 r4 內容錯誤，字串長度超過4000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r5>") != -1)
                    {
                        if ((LABDhdata_list[rdata_count].h1 == "1" || LABDhdata_list[rdata_count].h1 == "4") && LABDrdata_list[rdata_count].r5 == "")
                            error += "第" + line_count + "行 r5 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r5 != "" && LABDrdata_list[rdata_count].r5.Length > 50)
                            error += "第" + line_count + "行 r5 內容錯誤，字串長度超過50位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r6-1>") != -1)
                    {
                        if (LABDhdata_list[rdata_count].h1 == "1" && LABDrdata_list[rdata_count].r6_1 == "" && LABDrdata_list[rdata_count].r6_2 == "")
                            error += "第" + line_count + "行 r6-1 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r6_1 != "" && LABDrdata_list[rdata_count].r6_1.Length > 1000)
                            error += "第" + line_count + "行 r6-1 內容錯誤，字串長度超過1000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r6-2>") != -1)
                    {
                        if (LABDhdata_list[rdata_count].h1 == "1" && LABDrdata_list[rdata_count].r6_1 == "" && LABDrdata_list[rdata_count].r6_2 == "")
                            error += "第" + line_count + "行 r6-2 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r6_2 != null && LABDrdata_list[rdata_count].r6_2 != "" && LABDrdata_list[rdata_count].r6_2.Length > 1000)
                            error += "第" + line_count + "行 r6-2 內容錯誤，字串長度超過1000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r7>") != -1)
                    {
                        if (LABDhdata_list[rdata_count].h1 == "2" && LABDrdata_list[rdata_count].r7 == "")
                            error += "第" + line_count + "行 r7 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r7 != "" && LABDrdata_list[rdata_count].r7.Length > 4000)
                            error += "第" + line_count + "行 r7 內容錯誤，字串長度超過4000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r8-1>") != -1)
                    {
                        if (LABDhdata_list[rdata_count].h1 == "3" && LABDrdata_list[rdata_count].r8_1 == "")
                            error += "第" + line_count + "行 r8-1 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r8_1 != "" && LABDrdata_list[rdata_count].r8_1.Length > 4000)
                            error += "第" + line_count + "行 r8-1 內容錯誤，字串長度超過4000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r10>") != -1)
                    {
                        code_check_error = Date(LABDrdata_list[rdata_count].r10, 11);
                        if ((LABDhdata_list[rdata_count].h1 == "1" || LABDhdata_list[rdata_count].h1 == "2" || LABDhdata_list[rdata_count].h1 == "3") && LABDrdata_list[rdata_count].r10 == "")
                            error += "第" + line_count + "行 r10 欄位不得為空值\n";
                        else if (LABDrdata_list[rdata_count].r10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 r10 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<r12>") != -1)
                    {
                        code_check_error = Check(LABDrdata_list[rdata_count].r12, 0, 1);
                        if (LABDrdata_list[rdata_count].r12 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 r12 內容錯誤，" + code_check_error + "\n";
                    }
                }
                file_read.Close();

                if (error == "")
                    error = "沒有錯誤";

                error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

                return error;
            }
        }

        // LABM資料
        public void LABM_xml(string path)
        {
            XmlDocument doc = new XmlDocument();
            doc.Load(path);

            LABMhdata_list = new List<LABM_Hdata>();
            LABMrdata_list = new List<LABM_Rdata>();

            XmlNodeList nodes = doc.DocumentElement.SelectNodes("/patient/hdata");

            for (int i = 0; i < nodes.Count; i++)
            {
                LABM_hdata = new LABM_Hdata();

                if (nodes[i].SelectSingleNode("h1") != null)
                    LABM_hdata.h1 = nodes[i].SelectSingleNode("h1").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h2") != null)
                    LABM_hdata.h2 = nodes[i].SelectSingleNode("h2").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h3") != null)
                    LABM_hdata.h3 = nodes[i].SelectSingleNode("h3").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h4") != null)
                    LABM_hdata.h4 = nodes[i].SelectSingleNode("h4").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h5") != null)
                    LABM_hdata.h5 = nodes[i].SelectSingleNode("h5").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h6") != null)
                    LABM_hdata.h6 = nodes[i].SelectSingleNode("h6").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h7") != null)
                    LABM_hdata.h7 = nodes[i].SelectSingleNode("h7").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h8") != null)
                    LABM_hdata.h8 = nodes[i].SelectSingleNode("h8").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h9") != null)
                    LABM_hdata.h9 = nodes[i].SelectSingleNode("h9").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h10") != null)
                    LABM_hdata.h10 = nodes[i].SelectSingleNode("h10").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h11") != null)
                    LABM_hdata.h11 = nodes[i].SelectSingleNode("h11").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h12") != null)
                    LABM_hdata.h12 = nodes[i].SelectSingleNode("h12").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h13") != null)
                    LABM_hdata.h13 = nodes[i].SelectSingleNode("h13").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h14") != null)
                    LABM_hdata.h14 = nodes[i].SelectSingleNode("h14").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h17") != null)
                    LABM_hdata.h17 = nodes[i].SelectSingleNode("h17").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h18") != null)
                    LABM_hdata.h18 = nodes[i].SelectSingleNode("h18").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h22") != null)
                    LABM_hdata.h22 = nodes[i].SelectSingleNode("h22").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h23") != null)
                    LABM_hdata.h23 = nodes[i].SelectSingleNode("h23").InnerText.Trim();
                if (nodes[i].SelectSingleNode("h25") != null)
                    LABM_hdata.h25 = nodes[i].SelectSingleNode("h25").InnerText.Trim();

                XmlNodeList nodes_rdata = nodes[i].SelectNodes("rdata");
                for (int j = 0; j < nodes_rdata.Count; j++)
                {
                    LABM_rdata = new LABM_Rdata();

                    if (nodes_rdata[j].SelectSingleNode("r1") != null)
                        LABM_rdata.r1 = nodes_rdata[j].SelectSingleNode("r1").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r2") != null)
                        LABM_rdata.r2 = nodes_rdata[j].SelectSingleNode("r2").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r3") != null)
                        LABM_rdata.r3 = nodes_rdata[j].SelectSingleNode("r3").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r4") != null)
                        LABM_rdata.r4 = nodes_rdata[j].SelectSingleNode("r4").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r5") != null)
                        LABM_rdata.r5 = nodes_rdata[j].SelectSingleNode("r5").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r6-1") != null)
                        LABM_rdata.r6_1 = nodes_rdata[j].SelectSingleNode("r6-1").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r6-2") != null)
                        LABM_rdata.r6_2 = nodes_rdata[j].SelectSingleNode("r6-2").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r7") != null)
                        LABM_rdata.r7 = nodes_rdata[j].SelectSingleNode("r7").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r8-1") != null)
                        LABM_rdata.r8_1 = nodes_rdata[j].SelectSingleNode("r8-1").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r10") != null)
                        LABM_rdata.r10 = nodes_rdata[j].SelectSingleNode("r10").InnerText.Trim();
                    if (nodes_rdata[j].SelectSingleNode("r12") != null)
                        LABM_rdata.r12 = nodes_rdata[j].SelectSingleNode("r12").InnerText.Trim();

                    LABMhdata_list.Add(LABM_hdata);
                    LABMrdata_list.Add(LABM_rdata);
                }
            }
        }

        // 每月LAB錯誤檢查
        public string LABM_error(string path, string XMLfileName)
        {
            string error = "";
            var code_check_error = "";
            int line_count = 0;
            int rdata_count = 0;

            // 檢查內容
            using (var file_read = new StreamReader(path))
            {
                while (!file_read.EndOfStream)
                {
                    string line = file_read.ReadLine();
                    line_count++;

                    if (line.IndexOf("</rdata>") != -1)
                        rdata_count++;

                    if (line.IndexOf("<h1>") != -1)
                    {
                        code_check_error = Check(LABMhdata_list[rdata_count].h1, 1, 4);
                        if (LABMhdata_list[rdata_count].h1 == "")
                            error += "第" + line_count + "行 h1 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h2>") != -1)
                    {
                        code_check_error = Code(LABMhdata_list[rdata_count].h2, 10, 2);
                        if (LABMhdata_list[rdata_count].h2 == "")
                            error += "第" + line_count + "行 h2 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h2 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h2 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h3>") != -1)
                    {
                        code_check_error = Code(LABMhdata_list[rdata_count].h3, 2, 1);
                        code_check_error = Rule_h3(LABMhdata_list[rdata_count].h3);
                        if (LABMhdata_list[rdata_count].h3 == "")
                            error += "第" + line_count + "行 h3 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h3 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h3 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h4>") != -1)
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h4, 5);
                        if (LABMhdata_list[rdata_count].h4 == "")
                            error += "第" + line_count + "行 h4 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h4 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h4 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h5>") != -1)
                    {
                        code_check_error = Check(LABMhdata_list[rdata_count].h5, 1, 2);
                        if (LABMhdata_list[rdata_count].h5 == "")
                            error += "第" + line_count + "行 h5 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h5 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h5 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h6>") != -1)
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h6, 7);
                        if (LABMhdata_list[rdata_count].h6 == "")
                            error += "第" + line_count + "行 h6 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h6 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h6 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h7>") != -1)
                    {
                        code_check_error = Code(LABMhdata_list[rdata_count].h7, 2, 3);
                        if (LABMhdata_list[rdata_count].h7 == "")
                            error += "第" + line_count + "行 h7 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h7 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h7 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h8>") != -1)
                    {
                        code_check_error = Code(LABMhdata_list[rdata_count].h8, 6, 5);
                        if (LABMhdata_list[rdata_count].h8 == "")
                            error += "第" + line_count + "行 h8 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h8 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h8 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h9>") != -1)
                    {
                        code_check_error = Idcard(LABMhdata_list[rdata_count].h9, 10);
                        if (LABMhdata_list[rdata_count].h9 == "")
                            error += "第" + line_count + "行 h9 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h9 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h9 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h10>") != -1)
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h10, 7);
                        if (code_check_error == "OK")
                        {
                            //LAB需要比對 h11就醫日期 h13入院日期
                            if (LABMhdata_list[rdata_count].h11 != "")
                            {
                                code_check_error = Birth(LABMhdata_list[rdata_count].h10, LABMhdata_list[rdata_count].h11);
                            }
                            else if (LABMhdata_list[rdata_count].h13 != "")
                            {
                                code_check_error = Birth(LABMhdata_list[rdata_count].h10, LABMhdata_list[rdata_count].h13);
                            }
                        }
                        if (LABMhdata_list[rdata_count].h10 == "")
                            error += "第" + line_count + "行 h10 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h10 內容錯誤，" + code_check_error + "\n";
                    }
                    //就醫日期 入院日期 則一
                    else if (line.IndexOf("<h11>") != -1 && (LABMhdata_list[rdata_count].h13 == "" && LABMhdata_list[rdata_count].h14 == ""))
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h11, 7);
                        if (LABMhdata_list[rdata_count].h11 == "")
                            error += "第" + line_count + "行 h11 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h11 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h11 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h12>") != -1 && (LABMhdata_list[rdata_count].h13 == "" && LABMhdata_list[rdata_count].h14 == ""))
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h12, 7);
                        if ((LABMhdata_list[rdata_count].h7 == "08" || LABMhdata_list[rdata_count].h7 == "28") && LABMhdata_list[rdata_count].h12 == "")
                            error += "第" + line_count + "行 h12 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h12 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h12 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h13>") != -1 && (LABMhdata_list[rdata_count].h11 == "" && LABMhdata_list[rdata_count].h12 == ""))
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h13, 7);
                        if (code_check_error == "OK")
                            code_check_error = Limit(LABMhdata_list[rdata_count].h13, LABMhdata_list[rdata_count].h4, "0");
                        if (LABMhdata_list[rdata_count].h13 == "")
                            error += "第" + line_count + "行 h13 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h13 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h13 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h14>") != -1 && (LABMhdata_list[rdata_count].h11 == "" && LABMhdata_list[rdata_count].h12 == ""))
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h14, 7);
                        if (code_check_error == "OK")
                            code_check_error = Limit(LABMhdata_list[rdata_count].h14, LABMhdata_list[rdata_count].h4, LABMhdata_list[rdata_count].h13);
                        if (LABMhdata_list[rdata_count].h14 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h14 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h17>") != -1)
                    {
                        code_check_error = Code(LABMhdata_list[rdata_count].h17, 5, 5);
                        if (LABMhdata_list[rdata_count].h17 == "")
                            error += "第" + line_count + "行 h17 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h17 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h17 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h18>") != -1)
                    {
                        code_check_error = Code(LABMhdata_list[rdata_count].h18, 12, 3);
                        if (LABMhdata_list[rdata_count].h18 == "")
                            error += "第" + line_count + "行 h18 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h18 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h18 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h22>") != -1)
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h22, 11);
                        if (LABMhdata_list[rdata_count].h22 == "")
                            error += "第" + line_count + "行 h22 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h22 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h22 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h23>") != -1)
                    {
                        code_check_error = Date(LABMhdata_list[rdata_count].h23, 11);
                        if (LABMhdata_list[rdata_count].h23 == "")
                            error += "第" + line_count + "行 h23 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h23 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 h23 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<h25>") != -1)
                    {
                        if ((LABMhdata_list[rdata_count].h1 == "1" || LABMhdata_list[rdata_count].h1 == "3") && LABMhdata_list[rdata_count].h25 == "")
                            error += "第" + line_count + "行 h25 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h25 != "" && LABMhdata_list[rdata_count].h25.Length > 200)
                            error += "第" + line_count + "行 h25 內容錯誤，字串長度超過200位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r1>") != -1)
                    {
                        code_check_error = Code(LABMrdata_list[rdata_count].r1, 6, 4);
                        if (LABMrdata_list[rdata_count].r1 == "")
                            error += "第" + line_count + "行 r1 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r1 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 r1 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<r2>") != -1)
                    {
                        if ((LABMhdata_list[rdata_count].h1 == "1" || LABMhdata_list[rdata_count].h1 == "3" || LABMhdata_list[rdata_count].h1 == "4") && LABMrdata_list[rdata_count].r2 == "")
                            error += "第" + line_count + "行 r2 欄位不得為空值\n";
                        else if (LABMhdata_list[rdata_count].h1 == "4" && int.Parse(LABMrdata_list[rdata_count].r1) == 1 && LABMrdata_list[rdata_count].r2 != "BH")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為1，本欄請填BH\n";
                        else if (LABMhdata_list[rdata_count].h1 == "4" && int.Parse(LABMrdata_list[rdata_count].r1) == 2 && LABMrdata_list[rdata_count].r2 != "BW")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為2，本欄請填BW\n";
                        else if (LABMhdata_list[rdata_count].h1 == "4" && int.Parse(LABMrdata_list[rdata_count].r1) == 3 && LABMrdata_list[rdata_count].r2 != "ASA")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為3，本欄請填ASA\n";
                        else if (LABMhdata_list[rdata_count].h1 == "4" && int.Parse(LABMrdata_list[rdata_count].r1) == 4 && LABMrdata_list[rdata_count].r2 != "Surgical Approach")
                            error += "第" + line_count + "行 r2 內容錯誤，「h1報告類別」為4且「r1報告序號」為4，本欄請填BSurgical Approach\n";
                        else if (LABMrdata_list[rdata_count].r2 != "" && LABMrdata_list[rdata_count].r2.Length > 100)
                            error += "第" + line_count + "行 r2 內容錯誤，字串長度超過100位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r3>") != -1)
                    {
                        if ((LABMhdata_list[rdata_count].h1 == "1" || LABMhdata_list[rdata_count].h1 == "3" || LABMhdata_list[rdata_count].h1 == "4") && LABMrdata_list[rdata_count].r3 == "")
                            error += "第" + line_count + "行 r3 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r3 != "" && LABMrdata_list[rdata_count].r3.Length > 100)
                            error += "第" + line_count + "行 r3 內容錯誤，字串長度超過100位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r4>") != -1)
                    {
                        if ((LABMhdata_list[rdata_count].h1 == "1" || LABMhdata_list[rdata_count].h1 == "4") && LABMrdata_list[rdata_count].r4 == "")
                            error += "第" + line_count + "行 r4 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r4 != "" && LABMrdata_list[rdata_count].r4.Length > 4000)
                            error += "第" + line_count + "行 r4 內容錯誤，字串長度超過4000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r5>") != -1)
                    {
                        if ((LABMhdata_list[rdata_count].h1 == "1" || LABMhdata_list[rdata_count].h1 == "4") && LABMrdata_list[rdata_count].r5 == "")
                            error += "第" + line_count + "行 r5 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r5 != "" && LABMrdata_list[rdata_count].r5.Length > 50)
                            error += "第" + line_count + "行 r5 內容錯誤，字串長度超過50位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r6-1>") != -1)
                    {
                        if (LABMhdata_list[rdata_count].h1 == "1" && LABMrdata_list[rdata_count].r6_1 == "" && LABMrdata_list[rdata_count].r6_2 == "")
                            error += "第" + line_count + "行 r6-1 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r6_1 != "" && LABMrdata_list[rdata_count].r6_1.Length > 1000)
                            error += "第" + line_count + "行 r6-1 內容錯誤，字串長度超過1000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r6-2>") != -1)
                    {
                        if (LABMhdata_list[rdata_count].h1 == "1" && LABMrdata_list[rdata_count].r6_1 == "" && LABMrdata_list[rdata_count].r6_2 == "")
                            error += "第" + line_count + "行 r6-2 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r6_2 != null && LABMrdata_list[rdata_count].r6_2 != "" && LABMrdata_list[rdata_count].r6_2.Length > 1000)
                            error += "第" + line_count + "行 r6-2 內容錯誤，字串長度超過1000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r7>") != -1)
                    {
                        if (LABMhdata_list[rdata_count].h1 == "2" && LABMrdata_list[rdata_count].r7 == "")
                            error += "第" + line_count + "行 r7 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r7 != "" && LABMrdata_list[rdata_count].r7.Length > 4000)
                            error += "第" + line_count + "行 r7 內容錯誤，字串長度超過4000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r8-1>") != -1)
                    {
                        if (LABMhdata_list[rdata_count].h1 == "3" && LABMrdata_list[rdata_count].r8_1 == "")
                            error += "第" + line_count + "行 r8-1 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r8_1 != "" && LABMrdata_list[rdata_count].r8_1.Length > 4000)
                            error += "第" + line_count + "行 r8-1 內容錯誤，字串長度超過4000位元(符號算1位元)\n";
                    }
                    else if (line.IndexOf("<r10>") != -1)
                    {
                        code_check_error = Date(LABMrdata_list[rdata_count].r10, 11);
                        if ((LABMhdata_list[rdata_count].h1 == "1" || LABMhdata_list[rdata_count].h1 == "2" || LABMhdata_list[rdata_count].h1 == "3") && LABMrdata_list[rdata_count].r10 == "")
                            error += "第" + line_count + "行 r10 欄位不得為空值\n";
                        else if (LABMrdata_list[rdata_count].r10 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 r10 內容錯誤，" + code_check_error + "\n";
                    }
                    else if (line.IndexOf("<r12>") != -1)
                    {
                        code_check_error = Check(LABMrdata_list[rdata_count].r12, 0, 1);
                        if (LABMrdata_list[rdata_count].r12 != "" && code_check_error != "OK")
                            error += "第" + line_count + "行 r12 內容錯誤，" + code_check_error + "\n";
                    }
                }
                file_read.Close();

                if (error == "")
                    error = "沒有錯誤";

                error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

                return error;
            }
        }

        //每日檢驗申報
        public string LABD_EXCELtransfer(string path, string XMLfileName)
        {
            LABD_xml(path);
            string report = LABD_error(path, XMLfileName);
            //將身分證換乘Hash
            //Hash_Id_Birth("LAB_D");

            // CSV路徑
            var path_csv = Server.MapPath("~/data_excel/");
            // 若資料夾不存在則建立
            if (!Directory.Exists(path_csv))
            {
                Directory.CreateDirectory(path_csv);
            }

            // ZIP路徑
            var path_zip = Server.MapPath("~/data_excel_zip/");
            if (!Directory.Exists(path_zip))
            {
                Directory.CreateDirectory(path_zip);
            }

            // Error路徑
            var path_error = Server.MapPath("~/data_error/");
            // 若資料夾不存在則建立
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            // 檔名
            string fileName = "";
            var date = XMLfileName.Substring(0, 15);
            fileName = date + "LABD.xlsx";
            fileZip = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_ZipFile.zip";
            fileError = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_xml_Error.txt";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("LABM");

                sheet.Cells[1, 1].Value = "h1";
                sheet.Cells[1, 2].Value = "h2";
                sheet.Cells[1, 3].Value = "h3";
                sheet.Cells[1, 4].Value = "h4";
                sheet.Cells[1, 5].Value = "h5";
                sheet.Cells[1, 6].Value = "h6";
                sheet.Cells[1, 7].Value = "h7";
                sheet.Cells[1, 8].Value = "h8";
                sheet.Cells[1, 9].Value = "h9";
                sheet.Cells[1, 10].Value = "h10";
                sheet.Cells[1, 11].Value = "h11";
                sheet.Cells[1, 12].Value = "h12";
                sheet.Cells[1, 13].Value = "h13";
                sheet.Cells[1, 14].Value = "h14";
                sheet.Cells[1, 15].Value = "h15";
                sheet.Cells[1, 16].Value = "h19";
                sheet.Cells[1, 17].Value = "h20";
                sheet.Cells[1, 18].Value = "h22";
                sheet.Cells[1, 19].Value = "r1";
                sheet.Cells[1, 20].Value = "r2";
                sheet.Cells[1, 21].Value = "r3";
                sheet.Cells[1, 22].Value = "r4";
                sheet.Cells[1, 23].Value = "r5";
                sheet.Cells[1, 24].Value = "r6_1";
                sheet.Cells[1, 25].Value = "r6_2";
                sheet.Cells[1, 26].Value = "r7";
                sheet.Cells[1, 27].Value = "r8_1";
                sheet.Cells[1, 28].Value = "r10";
                sheet.Cells[1, 29].Value = "r12";

                var cell_number = 0;

                for (int i = 0; i < LABDhdata_list.Count; i++)
                {
                    patient_Hash = new Patient_Hash_Guid()
                    {
                        Patient_Id = LABDhdata_list[i].h9,
                        Patient_Birth = LABDhdata_list[i].h10
                    };

                    var id_exist = patient_Hash_Guids.Find(a => a.Patient_Id == patient_Hash.Patient_Id && a.Patient_Birth == patient_Hash.Patient_Birth);
                    if (id_exist != null)
                    {
                        sheet.Cells[(i + 2 + cell_number), 1].Value = LABDhdata_list[i].h1;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = LABDhdata_list[i].h2;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = LABDhdata_list[i].h3;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = LABDhdata_list[i].h4;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = LABDhdata_list[i].h5;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = LABDhdata_list[i].h6;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = LABDhdata_list[i].h7;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = LABDhdata_list[i].h8;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = LABDhdata_list[i].h9;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = LABDhdata_list[i].h10;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = LABDhdata_list[i].h11;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = LABDhdata_list[i].h12;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = LABDhdata_list[i].h13;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = LABDhdata_list[i].h14;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = LABDhdata_list[i].h15;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = LABDhdata_list[i].h19;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = LABDhdata_list[i].h20;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = LABDhdata_list[i].h22;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = LABDrdata_list[i].r1;
                        sheet.Cells[(i + 2 + cell_number), 20].Value = LABDrdata_list[i].r2;
                        sheet.Cells[(i + 2 + cell_number), 21].Value = LABDrdata_list[i].r3;
                        sheet.Cells[(i + 2 + cell_number), 22].Value = LABDrdata_list[i].r4;
                        sheet.Cells[(i + 2 + cell_number), 23].Value = LABDrdata_list[i].r5;
                        sheet.Cells[(i + 2 + cell_number), 24].Value = LABDrdata_list[i].r6_1;
                        sheet.Cells[(i + 2 + cell_number), 25].Value = LABDrdata_list[i].r6_2;
                        sheet.Cells[(i + 2 + cell_number), 26].Value = LABDrdata_list[i].r7;
                        sheet.Cells[(i + 2 + cell_number), 27].Value = LABDrdata_list[i].r8_1;
                        sheet.Cells[(i + 2 + cell_number), 28].Value = LABDrdata_list[i].r10;
                        sheet.Cells[(i + 2 + cell_number), 29].Value = LABDrdata_list[i].r12;
                    }
                    else
                    {
                        cell_number--;
                    }
                }
                p.SaveAs(new FileInfo(path_csv + fileName));
            }

            // CSV壓縮ZIP
            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + fileName, fileName);
                }
            }

            // 錯誤報告
            using (var file_error = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file_error.WriteLine(report);
            }

            return "轉檔成功";
        }

        //每月檢驗申報
        public string LABM_EXCELtransfer(string path, string XMLfileName)
        {
            LABM_xml(path);
            string report = LABM_error(path, XMLfileName);
            //將身分證換乘Hash
            //Hash_Id_Birth("LAB_M");

            // CSV路徑
            var path_csv = Server.MapPath("~/data_excel/");
            // 若資料夾不存在則建立
            if (!Directory.Exists(path_csv))
            {
                Directory.CreateDirectory(path_csv);
            }

            // ZIP路徑
            var path_zip = Server.MapPath("~/data_excel_zip/");
            if (!Directory.Exists(path_zip))
            {
                Directory.CreateDirectory(path_zip);
            }

            // Error路徑
            var path_error = Server.MapPath("~/data_error/");
            // 若資料夾不存在則建立
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            // 檔名
            string fileName = "";
            var date = XMLfileName.Substring(0, 15);
            fileName = date + "LABM.xlsx";
            fileZip = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_ZipFile.zip";
            fileError = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_xml_Error.txt";

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("LABM");

                sheet.Cells[1, 1].Value = "h1";
                sheet.Cells[1, 2].Value = "h2";
                sheet.Cells[1, 3].Value = "h3";
                sheet.Cells[1, 4].Value = "h4";
                sheet.Cells[1, 5].Value = "h5";
                sheet.Cells[1, 6].Value = "h6";
                sheet.Cells[1, 7].Value = "h7";
                sheet.Cells[1, 8].Value = "h8";
                sheet.Cells[1, 9].Value = "h9";
                sheet.Cells[1, 10].Value = "h10";
                sheet.Cells[1, 11].Value = "h11";
                sheet.Cells[1, 12].Value = "h12";
                sheet.Cells[1, 13].Value = "h13";
                sheet.Cells[1, 14].Value = "h14";
                sheet.Cells[1, 15].Value = "h17";
                sheet.Cells[1, 16].Value = "h18";
                sheet.Cells[1, 17].Value = "h22";
                sheet.Cells[1, 18].Value = "h23";
                sheet.Cells[1, 19].Value = "h25";
                sheet.Cells[1, 20].Value = "r1";
                sheet.Cells[1, 21].Value = "r2";
                sheet.Cells[1, 22].Value = "r3";
                sheet.Cells[1, 23].Value = "r4";
                sheet.Cells[1, 24].Value = "r5";
                sheet.Cells[1, 25].Value = "r6_1";
                sheet.Cells[1, 26].Value = "r6_2";
                sheet.Cells[1, 27].Value = "r7";
                sheet.Cells[1, 28].Value = "r8_1";
                sheet.Cells[1, 29].Value = "r10";
                sheet.Cells[1, 30].Value = "r12";

                var cell_number = 0;

                for (int i = 0; i < LABMhdata_list.Count; i++)
                {
                    patient_Hash = new Patient_Hash_Guid()
                    {
                        Patient_Id = LABMhdata_list[i].h9,
                        Patient_Birth = LABMhdata_list[i].h10
                    };

                    var id_exist = patient_Hash_Guids.Find(a => a.Patient_Id == patient_Hash.Patient_Id && a.Patient_Birth == patient_Hash.Patient_Birth);
                    if (id_exist != null)
                    {
                        sheet.Cells[(i + 2 + cell_number), 1].Value = LABMhdata_list[i].h1;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = LABMhdata_list[i].h2;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = LABMhdata_list[i].h3;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = LABMhdata_list[i].h4;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = LABMhdata_list[i].h5;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = LABMhdata_list[i].h6;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = LABMhdata_list[i].h7;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = LABMhdata_list[i].h8;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = LABMhdata_list[i].h9;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = LABMhdata_list[i].h10;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = LABMhdata_list[i].h11;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = LABMhdata_list[i].h12;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = LABMhdata_list[i].h13;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = LABMhdata_list[i].h14;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = LABMhdata_list[i].h17;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = LABMhdata_list[i].h18;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = LABMhdata_list[i].h22;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = LABMhdata_list[i].h23;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = LABMhdata_list[i].h25;
                        sheet.Cells[(i + 2 + cell_number), 20].Value = LABMrdata_list[i].r1;
                        sheet.Cells[(i + 2 + cell_number), 21].Value = LABMrdata_list[i].r2;
                        sheet.Cells[(i + 2 + cell_number), 22].Value = LABMrdata_list[i].r3;
                        sheet.Cells[(i + 2 + cell_number), 23].Value = LABMrdata_list[i].r4;
                        sheet.Cells[(i + 2 + cell_number), 24].Value = LABMrdata_list[i].r5;
                        sheet.Cells[(i + 2 + cell_number), 25].Value = LABMrdata_list[i].r6_1;
                        sheet.Cells[(i + 2 + cell_number), 26].Value = LABMrdata_list[i].r6_2;
                        sheet.Cells[(i + 2 + cell_number), 27].Value = LABMrdata_list[i].r7;
                        sheet.Cells[(i + 2 + cell_number), 28].Value = LABMrdata_list[i].r8_1;
                        sheet.Cells[(i + 2 + cell_number), 29].Value = LABMrdata_list[i].r10;
                        sheet.Cells[(i + 2 + cell_number), 30].Value = LABMrdata_list[i].r12;

                    }
                    else
                    {
                        cell_number--;
                    }
                }
                p.SaveAs(new FileInfo(path_csv + fileName));
            }

            // CSV壓縮ZIP
            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + fileName, fileName);
                }
            }

            // 錯誤報告
            using (var file_error = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file_error.WriteLine(report);
            }

            return "轉檔成功";
        }

        public void Update_excel(HttpPostedFileBase file, HttpPostedFileBase file_excel, string type)
        {
            alert = "0";
            result = "";
            data = "";
            fileZip = "";
            fileError = "";
            var type_name = type;
            if (file != null && file_excel != null)
            {
                if (file.ContentLength > 0 && file_excel.ContentLength > 0)
                {
                    string extension = Path.GetExtension(file.FileName);
                    string extension_excel = Path.GetExtension(file_excel.FileName);

                    if (extension.Equals(".txt", StringComparison.OrdinalIgnoreCase) &&
                        (extension_excel.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension_excel.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var fileName_excel = Path.GetFileName(file_excel.FileName);

                        var path = Server.MapPath("~/data_updata");
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                        DateTime myDate = DateTime.Now;
                        fileName = myDate.ToString("yyyyMMddHHmmss") + "_" + fileName;
                        path = Path.Combine(path, fileName);
                        file.SaveAs(path);

                        var path_excel = Server.MapPath("~/ID_Upload/");
                        if (!Directory.Exists(path_excel))
                        {
                            Directory.CreateDirectory(path_excel);
                        }
                        fileName_excel = myDate.ToString("yyyyMMddHHmmss") + "_" + fileName_excel;
                        var path_upload_excel = Path.Combine(path_excel, fileName_excel);
                        file_excel.SaveAs(path_upload_excel);

                        IWorkbook workbook;
                        string filepath_excel = Server.MapPath("~/ID_Upload/" + fileName_excel);

                        using (FileStream fileStream = new FileStream(filepath_excel, FileMode.Open, FileAccess.Read))
                        {
                            if (extension_excel.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook = new HSSFWorkbook(fileStream);
                            }
                            else if (extension_excel.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook = new XSSFWorkbook(fileStream);
                            }
                            else
                            {
                                workbook = null;
                            }
                        }


                        if (workbook != null)
                        {
                            fileZip = "1";
                            data = GUID_Converter(workbook, fileName_excel);
                            if (data == "轉檔成功")
                            {
                                if (type_name == "CRLF_TXT")
                                {
                                    fileZip = "1";
                                    data = CRLF_TXT_Read(path, fileName);
                                }
                                else if (type_name == "CRSF_TXT")
                                {
                                    fileZip = "1";
                                    data = CRSF_TXT_Read(path, fileName);
                                }

                                if (data == "轉檔成功")
                                {
                                    alert = "1";
                                    result = "檔案已轉檔成功！";
                                }
                                else
                                {
                                    alert = "0";
                                    result = "上傳檔案內容錯誤！";
                                }
                            }
                            else
                            {
                                alert = "2";
                                result = "上傳\"篩選清單\"檔案內容錯誤！";
                            }
                        }
                        else
                        {
                            alert = "0";
                            result = "上傳\"篩選清單\"檔案內容錯誤！";
                        }
                    }
                    else if (!extension.Equals(".txt", StringComparison.OrdinalIgnoreCase))
                    {
                        alert = "0";
                        result = "只能接受txt檔案！";
                    }
                    else if (!(extension_excel.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension_excel.Equals(".xlsx", StringComparison.OrdinalIgnoreCase)))
                    {
                        alert = "0";
                        result = "\"篩選清單\"只能接受EXCEL檔案！";
                    }
                }
                else
                {
                    alert = "0";
                    result = "請上傳正確之檔案！";
                }
            }
            else
            {
                alert = "0";
                result = "請上傳檔案！";
            }
        }




        // ---------------------------------------------------------------------------------------------------------
        // ---------------------------------------------------------------------------------------------------------
        // ---------------------------------------------------------------------------------------------------------

        public string Error_CRLF(int listindex, CRLF data, string errortxt)
        {
            var code_check_error = "";
            if (data.LF1_1 != null)
            {
                code_check_error = Code(data.LF1_1, 10, 2);
                if (data.LF1_1 == null || data.LF1_1 == "")
                    errortxt += "第" + listindex + "行 1.1申報醫院代碼 欄位不得為空值\n";
                else if (data.LF1_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.1申報醫院代碼 內容錯誤，" + code_check_error + "\n";
            }
            //code_check_error = Code(data.LF1_2, 10, 2);
            //if (data.LF1_2 == "")
            //    errortxt += "第" + listindex + "行 1.2病歷號碼 欄位不得為空值\n";
            //else if (data.LF1_2 != "" && code_check_error != "OK")
            //    errortxt += "第" + listindex + "行 1.2病歷號碼 內容錯誤，" + code_check_error + "\n";
            //if (data.LF1_3 == "")
            //    errortxt += "第" + listindex + "行 1.3姓名(此欄位不需要) 欄位不得為空值\n";
            if (data.LF1_4 != null)
            {
                code_check_error = Idcard(data.LF1_4, 10);
                if (data.LF1_4 == "")
                    errortxt += "第" + listindex + "行 1.4身分證統一編號 欄位不得為空值\n";
                else if (data.LF1_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.4身分證統一編號 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF1_5 != null)
            {
                code_check_error = Check(data.LF1_5, 1, 4);
                if (data.LF1_5 == "")
                    errortxt += "第" + listindex + "行 1.5性別 欄位不得為空值\n";
                else if (data.LF1_5 != "" && code_check_error != "OK" && data.LF1_5 != "9")
                    errortxt += "第" + listindex + "行 1.5性別 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF1_6 != null)
            {
                code_check_error = Date(data.LF1_6, 8); ;
                if (data.LF1_6 == "")
                    errortxt += "第" + listindex + "行 1.6出生日期 欄位不得為空值\n";
                else if (data.LF1_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.6出生日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF1_7 != null)
            {
                code_check_error = Code(data.LF1_7, 4, 5);
                if (data.LF1_7 == "")
                    errortxt += "第" + listindex + "行 1.7戶籍地代碼 欄位不得為空值\n";
                else if (data.LF1_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.7戶籍地代碼 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_1 != null)
            {
                code_check_error = Code(data.LF2_1, 3, 5);
                code_check_error = Check(data.LF2_1, 0, 120);
                if (data.LF2_1 == "")
                    errortxt += "第" + listindex + "行 2.1診斷年齡 欄位不得為空值\n";
                else if (data.LF2_1 != "" && code_check_error != "OK" && data.LF2_1 != "999")
                    errortxt += "第" + listindex + "行 2.1診斷年齡 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_2 != null)
            {
                code_check_error = Code(data.LF2_2, 2, 5);
                code_check_error = Check(data.LF2_2, 1, 99);
                if (data.LF2_2 == "")
                    errortxt += "第" + listindex + "行 2.2癌症發生順序號碼 欄位不得為空值\n";
                else if (data.LF2_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.2癌症發生順序號碼 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_3 != null)
            {
                code_check_error = Check(data.LF2_3, 0, 9);
                code_check_error = Rule_crlf2_3(data.LF2_3);
                if (data.LF2_3 == "")
                    errortxt += "第" + listindex + "行 2.3個案分類 欄位不得為空值\n";
                else if (data.LF2_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.3個案分類 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_3_1 != null)
            {
                code_check_error = Check(data.LF2_3_1, 1, 8);
                code_check_error = Rule_crlf2_3_1(data.LF2_3_1);
                if (data.LF2_3_1 == "")
                    errortxt += "第" + listindex + "行 2.3.1診斷狀態分類 欄位不得為空值\n";
                else if (data.LF2_3_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.3.1診斷狀態分類 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_3_2 != null)
            {
                code_check_error = Check(data.LF2_3_2, 0, 9);
                if (data.LF2_3_2 == "")
                    errortxt += "第" + listindex + "行 2.3.2治療狀態分類 欄位不得為空值\n";
                else if (data.LF2_3_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.3.2治療狀態分類 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_4 != null)
            {
                code_check_error = Date(data.LF2_4, 8);
                if (data.LF2_4 == "")
                    errortxt += "第" + listindex + "行 2.4首次就診日期 欄位不得為空值\n";
                else if (data.LF2_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.4首次就診日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_5 != null)
            {
                code_check_error = Date(data.LF2_5, 8);
                if (data.LF2_5 == "")
                    errortxt += "第" + listindex + "行 2.5最初診斷日期 欄位不得為空值\n";
                else if (data.LF2_5 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.5最初診斷日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_6 != null)
            {
                code_check_error = Code(data.LF2_6, 4, 3);
                code_check_error = Check(data.LF2_6.Substring(1, 3), 0, 809);
                if (data.LF2_6 == "")
                    errortxt += "第" + listindex + "行 2.6原發部位 欄位不得為空值\n";
                else if (data.LF2_6.Substring(0, 1) != "C")
                    errortxt += "第" + listindex + "行 2.6原發部位 內容錯誤，代碼開頭必須為C\n";
                else if (data.LF2_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.6原發部位 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_7 != null)
            {
                code_check_error = Check(data.LF2_7, 0, 5);
                if (data.LF2_7 == "")
                    errortxt += "第" + listindex + "行 2.7側性 欄位不得為空值\n";
                else if (data.LF2_7 != "" && code_check_error != "OK" && data.LF2_7 != "9")
                    errortxt += "第" + listindex + "行 2.7側性 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_8 != null)
            {
                code_check_error = Code(data.LF2_8, 4, 3);
                if (data.LF2_8 == "")
                    errortxt += "第" + listindex + "行 2.8組織類型 欄位不得為空值\n";
                else if (data.LF2_8 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.8組織類型 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_9 != null)
            {
                code_check_error = Check(data.LF2_9, 2, 3);
                if (data.LF2_9 == "")
                    errortxt += "第" + listindex + "行 2.9性態碼 欄位不得為空值\n";
                else if (data.LF2_9 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.9性態碼 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_10_1 != null)
            {
                code_check_error = Code(data.LF2_10_1, 1, 2);
                code_check_error = Rule_crlf2_10_1(data.LF2_10_1);
                if (data.LF2_10_1 == "")
                    errortxt += "第" + listindex + "行 2.10.1臨床分級/分化 欄位不得為空值\n";
                else if (data.LF2_10_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.10.1臨床分級/分化 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_10_2 != null)
            {
                code_check_error = Code(data.LF2_10_2, 1, 2);
                code_check_error = Rule_crlf2_10_1(data.LF2_10_2);
                if (data.LF2_10_2 == "")
                    errortxt += "第" + listindex + "行 2.10.2病理分級/分化 欄位不得為空值\n";
                else if (data.LF2_10_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.10.2病理分級/分化 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_11 != null)
            {
                code_check_error = Check(data.LF2_11, 1, 9);
                if (data.LF2_11 == "")
                    errortxt += "第" + listindex + "行 2.11癌症確診方式 欄位不得為空值\n";
                else if (data.LF2_11 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.11癌症確診方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_12 != null)
            {
                code_check_error = Date(data.LF2_12, 8);
                if (data.LF2_12 == "")
                    errortxt += "第" + listindex + "行 2.12首次顯微鏡檢證實日期 欄位不得為空值\n";
                else if (data.LF2_12 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.12首次顯微鏡檢證實日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_13 != null)
            {
                code_check_error = Code(data.LF2_13, 3, 5);
                code_check_error = Check(data.LF2_13, 0, 999);
                if (data.LF2_13 == "")
                    errortxt += "第" + listindex + "行 2.13腫瘤大小 欄位不得為空值\n";
                else if (data.LF2_13 == "996" || data.LF2_13 == "997")
                    errortxt += "第" + listindex + "行 2.13腫瘤大小 內容錯誤，代碼錯誤\n";
                else if (data.LF2_13 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.13腫瘤大小 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_13_1 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Rule_crlf2_13_1(data.LF2_13_1);
                if (data.LF2_13_1 == "")
                    errortxt += "第" + listindex + "行 2.13.1神經侵襲 欄位不得為空值\n";
                else if (data.LF2_13_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.13.1神經侵襲 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_13_2 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Rule_crlf2_13_1(data.LF2_13_2);
                if (data.LF2_13_2 == "")
                    errortxt += "第" + listindex + "行 2.13.2淋巴管或血管侵犯 欄位不得為空值\n";
                else if (data.LF2_13_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.13.2淋巴管或血管侵犯 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_14 != null)
            {
                code_check_error = Code(data.LF2_14, 2, 5);
                code_check_error = Check(data.LF2_14, 0, 99);
                if (data.LF2_14 == "")
                    errortxt += "第" + listindex + "行 2.14區域淋巴結檢查數目 欄位不得為空值\n";
                else if (int.Parse(data.LF2_14) >= 91 && int.Parse(data.LF2_14) <= 94)
                    errortxt += "第" + listindex + "行 2.14區域淋巴結檢查數目 內容錯誤，代碼錯誤\n";
                else if (data.LF2_14 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.14區域淋巴結檢查數目 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF2_15 != null)
            {
                code_check_error = Code(data.LF2_15, 2, 5);
                if (data.LF2_15 == "")
                    errortxt += "第" + listindex + "行 2.15區域淋巴結侵犯數目 欄位不得為空值\n";
                else if ((int.Parse(data.LF2_14) >= 91 && int.Parse(data.LF2_14) <= 94) || int.Parse(data.LF2_14) == 96)
                    errortxt += "第" + listindex + "行 2.15區域淋巴結侵犯數目 內容錯誤，代碼錯誤\n";
                else if (data.LF2_15 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.15區域淋巴結侵犯數目 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_1 != null)
            {
                code_check_error = Date(data.LF3_1, 8);
                if (data.LF3_1 == "")
                    errortxt += "第" + listindex + "行 3.1診斷性及分期性手術處置日期 欄位不得為空值\n";
                else if (data.LF3_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.1診斷性及分期性手術處置日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_2 != null)
            {
                code_check_error = Code(data.LF3_2, 2, 1);
                code_check_error = Check(data.LF3_2, 0, 14);
                if (data.LF3_2 == "")
                    errortxt += "第" + listindex + "行 3.2外院診斷性及分期性手術處置 欄位不得為空值\n";
                else if (data.LF3_2 == "08")
                    errortxt += "第" + listindex + "行 3.2外院診斷性及分期性手術處置 內容錯誤，代碼錯誤\n";
                else if (data.LF3_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.2外院診斷性及分期性手術處置 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_3 != null)
            {
                code_check_error = Code(data.LF3_3, 2, 1);
                code_check_error = Check(data.LF3_3, 0, 14);
                if (data.LF3_3 == "")
                    errortxt += "第" + listindex + "行 3.3申報醫院診斷性及分期性手術處置 欄位不得為空值\n";
                else if (data.LF3_3 == "08")
                    errortxt += "第" + listindex + "行 3.3申報醫院診斷性及分期性手術處置 內容錯誤，代碼錯誤\n";
                else if (data.LF3_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.3申報醫院診斷性及分期性手術處置 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_4 != null)
            {
                code_check_error = Code(data.LF3_4, 4, 3);
                code_check_error = Rule_crlf3_4(data.LF3_4);
                if (data.LF3_4 == "")
                    errortxt += "第" + listindex + "行 3.4臨床T 欄位不得為空值\n";
                else if (data.LF3_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.4臨床T 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_5 != null)
            {
                code_check_error = Code(data.LF3_5, 3, 3);
                code_check_error = Rule_crlf3_5(data.LF3_5);
                if (data.LF3_5 == "")
                    errortxt += "第" + listindex + "行 3.5臨床N 欄位不得為空值\n";
                else if (data.LF3_5 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.5臨床N 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_6 != null)
            {
                code_check_error = Code(data.LF3_6, 3, 3);
                code_check_error = Rule_crlf3_6(data.LF3_6);
                if (data.LF3_6 == "")
                    errortxt += "第" + listindex + "行 3.6臨床M 欄位不得為空值\n";
                else if (data.LF3_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.6臨床M 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_7 != null)
            {
                code_check_error = Code(data.LF3_7, 3, 3);
                code_check_error = Rule_crlf3_7(data.LF3_7);
                if (data.LF3_7 == "")
                    errortxt += "第" + listindex + "行 3.7臨床期別組合 欄位不得為空值\n";
                else if (data.LF3_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.7臨床期別組合 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_8 != null)
            {
                code_check_error = Check(data.LF3_8, 0, 6);
                code_check_error = Rule_crlf3_8(data.LF3_8);
                if (data.LF3_8 == "")
                    errortxt += "第" + listindex + "行 3.8臨床分期字根/字首 欄位不得為空值\n";
                else if (data.LF3_8 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.8臨床分期字根/字首 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_10 != null)
            {
                code_check_error = Code(data.LF3_10, 4, 3);
                code_check_error = Rule_crlf3_10(data.LF3_10);
                if (data.LF3_10 == "")
                    errortxt += "第" + listindex + "行 3.10病理T 欄位不得為空值\n";
                else if (data.LF3_10 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.10病理T 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_11 != null)
            {
                code_check_error = Code(data.LF3_11, 3, 3);
                code_check_error = Rule_crlf3_11(data.LF3_11);
                if (data.LF3_11 == "")
                    errortxt += "第" + listindex + "行 3.11病理N 欄位不得為空值\n";
                else if (data.LF3_11 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.11病理N 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_12 != null)
            {
                code_check_error = Code(data.LF3_12, 3, 3);
                code_check_error = Rule_crlf3_12(data.LF3_12);
                if (data.LF3_12 == "")
                    errortxt += "第" + listindex + "行 3.12病理M 欄位不得為空值\n";
                else if (data.LF3_12 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.12病理M 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_13 != null)
            {
                code_check_error = Code(data.LF3_13, 3, 3);
                code_check_error = Rule_crlf3_13(data.LF3_13);
                if (data.LF3_13 == "")
                    errortxt += "第" + listindex + "行 3.13病理期別組合 欄位不得為空值\n";
                else if (data.LF3_13 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.13病理期別組合 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_14 != null)
            {
                code_check_error = Check(data.LF3_14, 0, 6);
                code_check_error = Rule_crlf3_14(data.LF3_14);
                if (data.LF3_14 == "")
                    errortxt += "第" + listindex + "行 3.14病理分期字根/字首 欄位不得為空值\n";
                else if (data.LF3_14 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.14病理分期字根/字首 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_16 != null)
            {
                code_check_error = Code(data.LF3_16, 5, 1);
                if (data.LF3_16 == "")
                    errortxt += "第" + listindex + "行 3.16AJCC癌症分期版本與章節 欄位不得為空值\n";
                else if (data.LF3_16 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.16AJCC癌症分期版本與章節 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_17 != null)
            {
                code_check_error = Code(data.LF3_17, 2, 1);
                code_check_error = Rule_crlf3_17(data.LF3_17);
                if (data.LF3_17 == "")
                    errortxt += "第" + listindex + "行 3.17其他分期系統 欄位不得為空值\n";
                else if (data.LF3_17 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.17其他分期系統 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_19 != null)
            {
                code_check_error = Code(data.LF3_19, 4, 3);
                if (data.LF3_19 == "")
                    errortxt += "第" + listindex + "行 3.19其他分期系統期別(臨床分期) 欄位不得為空值\n";
                else if (data.LF3_19 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.19其他分期系統期別(臨床分期) 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF3_21 != null)
            {
                code_check_error = Code(data.LF3_21, 4, 3);
                if (data.LF3_21 == "")
                    errortxt += "第" + listindex + "行 3.21其他分期系統期別(病理分期) 欄位不得為空值\n";
                else if (data.LF3_21 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 3.21其他分期系統期別(病理分期) 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1 != null)
            {
                code_check_error = Date(data.LF4_1, 8);
                if (data.LF4_1 == "")
                    errortxt += "第" + listindex + "行 4.1首次療程開始日期 欄位不得為空值\n";
                else if (data.LF4_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1首次療程開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_1 != null)
            {
                code_check_error = Date(data.LF4_1_1, 8);
                if (data.LF4_1_1 == "")
                    errortxt += "第" + listindex + "行 4.1.1首次手術日期 欄位不得為空值\n";
                else if (data.LF4_1_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.1首次手術日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_2 != null)
            {
                code_check_error = Date(data.LF4_1_2, 8);
                if (data.LF4_1_2 == "")
                    errortxt += "第" + listindex + "行 4.1.2原發部位最確切的手術切除日期 欄位不得為空值\n";
                else if (data.LF4_1_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.2原發部位最確切的手術切除日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_3 != null)
            {
                code_check_error = Code(data.LF4_1_3, 2, 1);
                code_check_error = Check(data.LF4_1_3, 0, 99);
                if (data.LF4_1_3 == "")
                    errortxt += "第" + listindex + "行 4.1.3外院原發部位手術方式 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_1_3) >= 1 && int.Parse(data.LF4_1_3) <= 9) || (int.Parse(data.LF4_1_3) >= 81 && int.Parse(data.LF4_1_3) <= 89) || (int.Parse(data.LF4_1_3) >= 91 && int.Parse(data.LF4_1_3) <= 97))
                    errortxt += "第" + listindex + "行 4.1.3外院原發部位手術方式 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.3外院原發部位手術方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_4 != null)
            {
                code_check_error = Code(data.LF4_1_4, 2, 1);
                code_check_error = Check(data.LF4_1_4, 0, 99);
                if (data.LF4_1_4 == "")
                    errortxt += "第" + listindex + "行 4.1.4申報醫院原發部位手術方式 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_1_4) >= 1 && int.Parse(data.LF4_1_4) <= 9) || (int.Parse(data.LF4_1_4) >= 81 && int.Parse(data.LF4_1_4) <= 89) || (int.Parse(data.LF4_1_4) >= 91 && int.Parse(data.LF4_1_4) <= 97))
                    errortxt += "第" + listindex + "行 4.1.4申報醫院原發部位手術方式 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.4申報醫院原發部位手術方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_4_1 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Check(data.LF4_1_4_1, 0, 9);
                if (data.LF4_1_4_1 == "")
                    errortxt += "第" + listindex + "行 4.1.4.1微創手術 欄位不得為空值\n";
                else if (int.Parse(data.LF4_1_4_1) >= 5 && int.Parse(data.LF4_1_4_1) >= 7)
                    errortxt += "第" + listindex + "行 4.1.4.1微創手術 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_4_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.4.1微創手術 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_5 != null)
            {
                code_check_error = Code(data.LF4_1_5, 1, 2);
                code_check_error = Rule_crlf4_1_5(data.LF4_1_5);
                if (data.LF4_1_5 == "")
                    errortxt += "第" + listindex + "行 4.1.5原發部位手術邊緣 欄位不得為空值\n";
                else if (data.LF4_1_5 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.5原發部位手術邊緣 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_5_1 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Check(data.LF4_1_5_1, 0, 999);
                if (data.LF4_1_5_1 == "")
                    errortxt += "第" + listindex + "行 4.1.5.1原發部位手術切緣距離 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_1_5_1) >= 981 && int.Parse(data.LF4_1_5_1) <= 986) || int.Parse(data.LF4_1_5_1) == 989 || (int.Parse(data.LF4_1_5_1) >= 992 && int.Parse(data.LF4_1_5_1) <= 998))
                    errortxt += "第" + listindex + "行 4.1.5.1原發部位手術切緣距離 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_5_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.5.1原發部位手術切緣距離 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_6 != null)
            {
                code_check_error = Check(data.LF4_1_6, 0, 9);
                if (data.LF4_1_6 == "")
                    errortxt += "第" + listindex + "行 4.1.6外院區域淋巴結手術範圍 欄位不得為空值\n";
                else if (data.LF4_1_6 == "8")
                    errortxt += "第" + listindex + "行 4.1.6外院區域淋巴結手術範圍 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.6外院區域淋巴結手術範圍 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_7 != null)
            {
                code_check_error = Check(data.LF4_1_7, 0, 9);
                if (data.LF4_1_7 == "")
                    errortxt += "第" + listindex + "行 4.1.7申報醫院區域淋巴結手術範圍 欄位不得為空值\n";
                else if (data.LF4_1_7 == "8")
                    errortxt += "第" + listindex + "行 4.1.7申報醫院區域淋巴結手術範圍 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.7申報醫院區域淋巴結手術範圍 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_8 != null)
            {
                code_check_error = Check(data.LF4_1_8, 0, 9);
                if (data.LF4_1_8 == "")
                    errortxt += "第" + listindex + "行 4.1.8外院其他部位手術方式 欄位不得為空值\n";
                else if (int.Parse(data.LF4_1_8) >= 6 && int.Parse(data.LF4_1_8) <= 8)
                    errortxt += "第" + listindex + "行 4.1.8外院其他部位手術方式 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_8 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.8外院其他部位手術方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_9 != null)
            {
                code_check_error = Check(data.LF4_1_9, 0, 9);
                if (data.LF4_1_9 == "")
                    errortxt += "第" + listindex + "行 4.1.9申報醫院其他部位手術方式 欄位不得為空值\n";
                else if (int.Parse(data.LF4_1_9) >= 6 && int.Parse(data.LF4_1_9) <= 8)
                    errortxt += "第" + listindex + "行 4.1.9申報醫院其他部位手術方式 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_9 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.9申報醫院其他部位手術方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_1_10 != null)
            {
                code_check_error = Check(data.LF4_1_10, 0, 9);
                if (data.LF4_1_10 == "")
                    errortxt += "第" + listindex + "行 4.1.10原發部位未手術原因 欄位不得為空值\n";
                else if (data.LF4_1_10 == "4")
                    errortxt += "第" + listindex + "行 4.1.10原發部位未手術原因 內容錯誤，代碼錯誤\n";
                else if (data.LF4_1_10 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.10原發部位未手術原因 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_1 != null)
            {
                code_check_error = Code(data.LF4_2_1_1, 2, 6);
                code_check_error = Check(data.LF4_2_1_1, 0, 63);
                if (data.LF4_2_1_1 == "")
                    errortxt += "第" + listindex + "行 4.2.1.1放射治療臨床標靶體積摘要 欄位不得為空值\n";
                else if (data.LF4_2_1_1 != "" && code_check_error != "OK" && (data.LF4_2_1_1 != "-9" || data.LF4_2_1_1 != "-1"))
                    errortxt += "第" + listindex + "行 4.2.1.1放射治療臨床標靶體積摘要 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_2 != null)
            {
                code_check_error = Code(data.LF4_2_1_2, 3, 6);
                code_check_error = Check(data.LF4_2_1_2, 0, 127);
                if (data.LF4_2_1_2 == "")
                    errortxt += "第" + listindex + "行 4.2.1.2放射治療儀器 欄位不得為空值\n";
                else if (data.LF4_2_1_2 != "" && code_check_error != "OK" && (data.LF4_2_1_1 != "-9" || data.LF4_2_1_1 != "-1"))
                    errortxt += "第" + listindex + "行 4.2.1.2放射治療儀器 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_3 != null)
            {
                code_check_error = Date(data.LF4_2_1_3, 8);
                if (data.LF4_2_1_3 == "")
                    errortxt += "第" + listindex + "行 4.2.1.3放射治療開始日期 欄位不得為空值\n";
                else if (data.LF4_2_1_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.1.3放射治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_4 != null)
            {
                code_check_error = Date(data.LF4_2_1_4, 8);
                if (data.LF4_2_1_4 == "")
                    errortxt += "第" + listindex + "行 4.2.1.4放射治療結束日期 欄位不得為空值\n";
                else if (data.LF4_2_1_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.1.4放射治療結束日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_5 != null)
            {
                code_check_error = Code(data.LF4_2_1_5, 2, 6);
                code_check_error = Rule_crlf4_2_1_5(data.LF4_2_1_5);
                if (data.LF4_2_1_5 == "")
                    errortxt += "第" + listindex + "行 4.2.1.5放射治療與手術順序 欄位不得為空值\n";
                else if (data.LF4_2_1_5 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.1.5放射治療與手術順序 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_6 != null)
            {
                code_check_error = Code(data.LF4_2_1_6, 2, 6);
                code_check_error = Rule_crlf4_2_1_6(data.LF4_2_1_6);
                if (data.LF4_2_1_6 == "")
                    errortxt += "第" + listindex + "行 4.2.1.6區域治療與全身性治療順序 欄位不得為空值\n";
                else if (data.LF4_2_1_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.1.6區域治療與全身性治療順序 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_1_8 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Check(data.LF4_2_1_8, 0, 10);
                if (data.LF4_2_1_8 == "")
                    errortxt += "第" + listindex + "行 4.2.1.8放射治療執行狀態 欄位不得為空值\n";
                else if (data.LF4_2_1_8 != "" && code_check_error != "OK" && data.LF4_2_1_8 != "99")
                    errortxt += "第" + listindex + "行 4.2.1.8放射治療執行狀態 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_1 != null)
            {
                code_check_error = Code(data.LF4_2_2_1, 3, 6);
                code_check_error = Check(data.LF4_2_2_1, 0, 111);
                if (data.LF4_2_2_1 == "")
                    errortxt += "第" + listindex + "行 4.2.2.1體外放射治療技術 欄位不得為空值\n";
                else if (data.LF4_2_2_1 != "" && code_check_error != "OK" && (data.LF4_2_2_1 != "-9" || data.LF4_2_2_1 != "-1"))
                    errortxt += "第" + listindex + "行 4.2.2.1體外放射治療技術 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_2_1 != null)
            {
                code_check_error = Code(data.LF4_2_2_2_1, 2, 6);
                code_check_error = Check(data.LF4_2_2_2_1, 0, 63);
                if (data.LF4_2_2_2_1 == "")
                    errortxt += "第" + listindex + "行 4.2.2.2.1最高放射劑量臨床標靶體積 欄位不得為空值\n";
                else if (data.LF4_2_2_2_1 != "" && code_check_error != "OK" && (data.LF4_2_2_2_1 != "-9" || data.LF4_2_2_2_1 != "-1"))
                    errortxt += "第" + listindex + "行 4.2.2.2.1最高放射劑量臨床標靶體積 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_2_2 != null)
            {
                code_check_error = Code(data.LF4_2_2_2_2, 5, 5);
                code_check_error = Check(data.LF4_2_2_2_2, 0, 99999);
                if (data.LF4_2_2_2_2 == "")
                    errortxt += "第" + listindex + "行 4.2.2.2.2最高放射劑量臨床標靶體積劑量 欄位不得為空值\n";
                else if (data.LF4_2_2_2_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.2.2.2最高放射劑量臨床標靶體積劑量 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_2_3 != null)
            {
                code_check_error = Code(data.LF4_2_2_2_3, 2, 5);
                code_check_error = Check(data.LF4_2_2_2_3, 0, 99);
                if (data.LF4_2_2_2_3 == "")
                    errortxt += "第" + listindex + "行 4.2.2.2.3最高放射劑量臨床標靶體積治療次數 欄位不得為空值\n";
                else if (data.LF4_2_2_2_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.2.2.3最高放射劑量臨床標靶體積治療次數 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_3_1 != null)
            {
                code_check_error = Code(data.LF4_2_2_3_1, 2, 6);
                if (data.LF4_2_2_3_1 == "")
                    errortxt += "第" + listindex + "行 4.2.2.3.1較低放射劑量臨床標靶體積 欄位不得為空值\n";
                else if (data.LF4_2_2_3_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.2.3.1較低放射劑量臨床標靶體積 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_3_2 != null)
            {
                code_check_error = Code(data.LF4_2_2_3_2, 5, 5);
                code_check_error = Check(data.LF4_2_2_3_2, 0, 99999);
                if (data.LF4_2_2_3_2 == "")
                    errortxt += "第" + listindex + "行 4.2.2.3.2最低放射劑量臨床標靶體積劑量 欄位不得為空值\n";
                else if (data.LF4_2_2_3_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.2.3.2最低放射劑量臨床標靶體積劑量 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_2_3_3 != null)
            {
                code_check_error = Code(data.LF4_2_2_3_3, 2, 5);
                code_check_error = Check(data.LF4_2_2_3_3, 0, 99);
                if (data.LF4_2_2_3_3 == "")
                    errortxt += "第" + listindex + "行 4.2.2.3.3最低放射劑量臨床標靶體積治療次數 欄位不得為空值\n";
                else if (data.LF4_2_2_3_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.2.3.3最低放射劑量臨床標靶體積治療次數 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_3_1 != null)
            {
                code_check_error = Code(data.LF4_2_3_1, 2, 6);
                code_check_error = Rule_crlf4_2_3_1(data.LF4_2_3_1);
                if (data.LF4_2_3_1 == "")
                    errortxt += "第" + listindex + "行 4.2.3.1其他放射治療儀器 欄位不得為空值\n";
                else if (data.LF4_2_3_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.3.1其他放射治療儀器 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_3_2 != null)
            {
                code_check_error = Code(data.LF4_2_3_2, 2, 6);
                code_check_error = Rule_crlf4_2_3_2(data.LF4_2_3_2);
                if (data.LF4_2_3_2 == "")
                    errortxt += "第" + listindex + "行 4.2.3.2其他放射治療技術 欄位不得為空值\n";
                else if (data.LF4_2_3_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.3.2其他放射治療技術 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_3_3_1 != null)
            {
                code_check_error = Code(data.LF4_2_3_3_1, 2, 6);
                code_check_error = Check(data.LF4_2_3_3_1, 0, 63);
                if (data.LF4_2_3_3_1 == "")
                    errortxt += "第" + listindex + "行 4.2.3.3.1其他放射治療臨床標靶體積 欄位不得為空值\n";
                else if (data.LF4_2_3_3_1 != "" && code_check_error != "OK" && (data.LF4_2_3_3_1 != "-9" || data.LF4_2_3_3_1 != "-1"))
                    errortxt += "第" + listindex + "行 4.2.3.3.1其他放射治療臨床標靶體積 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_3_3_2 != null)
            {
                code_check_error = Code(data.LF4_2_3_3_2, 5, 5);
                code_check_error = Check(data.LF4_2_3_3_2, 0, 99999);
                if (data.LF4_2_3_3_2 == "")
                    errortxt += "第" + listindex + "行 4.2.3.3.2其他放射劑量臨床標靶體積劑量 欄位不得為空值\n";
                else if (data.LF4_2_3_3_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.3.3.2其他放射劑量臨床標靶體積劑量 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_2_3_3_3 != null)
            {
                code_check_error = Code(data.LF4_2_3_3_3, 2, 5);
                code_check_error = Check(data.LF4_2_3_3_3, 0, 99);
                if (data.LF4_2_3_3_3 == "")
                    errortxt += "第" + listindex + "行 4.2.3.3.3其他放射劑量臨床標靶體積治療次數 欄位不得為空值\n";
                else if (data.LF4_2_3_3_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.3.3.3其他放射劑量臨床標靶體積治療次數 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_1 != null)
            {
                code_check_error = Date(data.LF4_3_1, 8);
                if (data.LF4_3_1 == "")
                    errortxt += "第" + listindex + "行 4.3.1全身性治療開始日期 欄位不得為空值\n";
                else if (data.LF4_3_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.1全身性治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_2 != null)
            {
                code_check_error = Code(data.LF4_3_2, 2, 1);
                code_check_error = Check(data.LF4_3_2, 0, 31);
                if (data.LF4_3_2 == "")
                    errortxt += "第" + listindex + "行 4.3.2外院化學治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_2) >= 14 && int.Parse(data.LF4_3_2) <= 19) || (int.Parse(data.LF4_3_2) >= 22 && int.Parse(data.LF4_3_2) <= 29))
                    errortxt += "第" + listindex + "行 4.3.2外院化學治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_2 != "" && code_check_error != "OK" && data.LF4_3_2 != "99")
                    errortxt += "第" + listindex + "行 4.3.2外院化學治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_3 != null)
            {
                code_check_error = Code(data.LF4_3_3, 2, 1);
                code_check_error = Check(data.LF4_3_3, 0, 99);
                if (data.LF4_3_3 == "")
                    errortxt += "第" + listindex + "行 4.3.3申報醫院化學治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_3) >= 14 && int.Parse(data.LF4_3_3) <= 19) || (int.Parse(data.LF4_3_3) >= 22 && int.Parse(data.LF4_3_3) <= 29) || (int.Parse(data.LF4_3_3) >= 32 && int.Parse(data.LF4_3_3) <= 81) || (int.Parse(data.LF4_3_3) == 84) || (int.Parse(data.LF4_3_3) >= 89 && int.Parse(data.LF4_3_3) <= 98))
                    errortxt += "第" + listindex + "行 4.3.3申報醫院化學治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.3申報醫院化學治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_4 != null)
            {
                code_check_error = Date(data.LF4_3_4, 8);
                if (data.LF4_3_4 == "")
                    errortxt += "第" + listindex + "行 4.3.4申報醫院化學治療開始日期 欄位不得為空值\n";
                else if (data.LF4_3_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.4申報醫院化學治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_5 != null)
            {
                code_check_error = Code(data.LF4_3_5, 2, 1);
                code_check_error = Check(data.LF4_3_5, 0, 31);
                if (data.LF4_3_5 == "")
                    errortxt += "第" + listindex + "行 4.3.5外院賀爾蒙/類固醇治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_5) >= 4 && int.Parse(data.LF4_3_5) <= 19) || (int.Parse(data.LF4_3_5) >= 22 && int.Parse(data.LF4_3_5) <= 29))
                    errortxt += "第" + listindex + "行 4.3.5外院賀爾蒙/類固醇治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_5 != "" && code_check_error != "OK" && data.LF4_3_5 != "99")
                    errortxt += "第" + listindex + "行 4.3.5外院賀爾蒙/類固醇治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_6 != null)
            {
                code_check_error = Code(data.LF4_3_6, 2, 1);
                code_check_error = Check(data.LF4_3_6, 0, 99);
                if (data.LF4_3_6 == "")
                    errortxt += "第" + listindex + "行 4.3.6申報醫院賀爾蒙/類固醇治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_6) >= 4 && int.Parse(data.LF4_3_6) <= 19) || (int.Parse(data.LF4_3_6) >= 22 && int.Parse(data.LF4_3_6) <= 29) || (int.Parse(data.LF4_3_6) >= 32 && int.Parse(data.LF4_3_6) <= 81) || (int.Parse(data.LF4_3_6) == 84) || (int.Parse(data.LF4_3_6) >= 89 && int.Parse(data.LF4_3_6) <= 98))
                    errortxt += "第" + listindex + "行 4.3.6申報醫院賀爾蒙/類固醇治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.6申報醫院賀爾蒙/類固醇治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_7 != null)
            {
                code_check_error = Date(data.LF4_3_7, 8);
                if (data.LF4_3_7 == "")
                    errortxt += "第" + listindex + "行 4.3.7申報醫院賀爾蒙/類固醇治療開始日期 欄位不得為空值\n";
                else if (data.LF4_3_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.7申報醫院賀爾蒙/類固醇治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_8 != null)
            {
                code_check_error = Code(data.LF4_3_8, 2, 1);
                code_check_error = Check(data.LF4_3_8, 0, 31);
                if (data.LF4_3_8 == "")
                    errortxt += "第" + listindex + "行 4.3.8外院免疫治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_8) >= 4 && int.Parse(data.LF4_3_8) <= 19) || (int.Parse(data.LF4_3_8) >= 22 && int.Parse(data.LF4_3_8) <= 29))
                    errortxt += "第" + listindex + "行 4.3.8外院免疫治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_8 != "" && code_check_error != "OK" && data.LF4_3_8 != "99")
                    errortxt += "第" + listindex + "行 4.3.8外院免疫治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_9 != null)
            {
                code_check_error = Code(data.LF4_3_9, 2, 1);
                code_check_error = Check(data.LF4_3_9, 0, 99);
                if (data.LF4_3_9 == "")
                    errortxt += "第" + listindex + "行 4.3.9申報醫院免疫治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_9) >= 4 && int.Parse(data.LF4_3_9) <= 19) || (int.Parse(data.LF4_3_9) >= 22 && int.Parse(data.LF4_3_9) <= 29) || (int.Parse(data.LF4_3_9) >= 32 && int.Parse(data.LF4_3_9) <= 81) || (int.Parse(data.LF4_3_9) == 84) || (int.Parse(data.LF4_3_9) >= 89 && int.Parse(data.LF4_3_9) <= 98))
                    errortxt += "第" + listindex + "行 4.3.9申報醫院免疫治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_9 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.9申報醫院免疫治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_10 != null)
            {
                code_check_error = Date(data.LF4_3_10, 8);
                if (data.LF4_3_10 == "")
                    errortxt += "第" + listindex + "行 4.3.10申報醫院免疫治療開始日期 欄位不得為空值\n";
                else if (data.LF4_3_10 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.10申報醫院免疫治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_11 != null)
            {
                code_check_error = Code(data.LF4_3_11, 2, 1);
                code_check_error = Check(data.LF4_3_11, 0, 99);
                if (data.LF4_3_11 == "")
                    errortxt += "第" + listindex + "行 4.3.11骨髓/幹細胞移植或內分泌處置 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_11) >= 1 && int.Parse(data.LF4_3_11) <= 9) || (int.Parse(data.LF4_3_11) >= 13 && int.Parse(data.LF4_3_11) <= 19) || (int.Parse(data.LF4_3_11) >= 23 && int.Parse(data.LF4_3_11) <= 24) || (int.Parse(data.LF4_3_11) >= 26 && int.Parse(data.LF4_3_11) <= 29) || (int.Parse(data.LF4_3_11) >= 31 && int.Parse(data.LF4_3_11) <= 39) || (int.Parse(data.LF4_3_11) >= 41 && int.Parse(data.LF4_3_11) <= 49) || (int.Parse(data.LF4_3_11) >= 51 && int.Parse(data.LF4_3_11) <= 81) || (int.Parse(data.LF4_3_11) == 84) || (int.Parse(data.LF4_3_11) >= 89 && int.Parse(data.LF4_3_11) <= 98))
                    errortxt += "第" + listindex + "行 4.3.11骨髓/幹細胞移植或內分泌處置 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_11 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.11骨髓/幹細胞移植或內分泌處置 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_12 != null)
            {
                code_check_error = Date(data.LF4_3_12, 8);
                if (data.LF4_3_12 == "")
                    errortxt += "第" + listindex + "行 4.3.12申報醫院骨髓/幹細胞移植或內分泌處置開始日期 欄位不得為空值\n";
                else if (data.LF4_3_12 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.12申報醫院骨髓/幹細胞移植或內分泌處置開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_13 != null)
            {
                code_check_error = Code(data.LF4_3_13, 2, 1);
                code_check_error = Check(data.LF4_3_13, 0, 31);
                if (data.LF4_3_13 == "")
                    errortxt += "第" + listindex + "行 4.3.13外院標靶治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_13) >= 2 && int.Parse(data.LF4_3_13) <= 19) || (int.Parse(data.LF4_3_13) >= 22 && int.Parse(data.LF4_3_13) <= 29))
                    errortxt += "第" + listindex + "行 4.3.13外院標靶治療 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_13 != "" && code_check_error != "OK" && data.LF4_3_13 != "99")
                    errortxt += "第" + listindex + "行 4.3.13外院標靶治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_14 != null)
            {
                code_check_error = Code(data.LF4_3_14, 2, 1);
                code_check_error = Check(data.LF4_3_14, 0, 99);
                if (data.LF4_3_14 == "")
                    errortxt += "第" + listindex + "行 4.3.14申報醫院標靶治療 欄位不得為空值\n";
                else if ((int.Parse(data.LF4_3_14) >= 2 && int.Parse(data.LF4_3_14) <= 19) || (int.Parse(data.LF4_3_14) >= 22 && int.Parse(data.LF4_3_14) <= 29) || (int.Parse(data.LF4_3_14) >= 32 && int.Parse(data.LF4_3_14) <= 81) || (int.Parse(data.LF4_3_14) == 84) || (int.Parse(data.LF4_3_14) >= 89 && int.Parse(data.LF4_3_14) <= 98))
                    errortxt += "第" + listindex + "行 4.3.14申報醫院標靶治療內容錯誤 內容錯誤，代碼錯誤\n";
                else if (data.LF4_3_14 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.14申報醫院標靶治療內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_3_15 != null)
            {
                code_check_error = Date(data.LF4_3_15, 8);
                if (data.LF4_3_15 == "")
                    errortxt += "第" + listindex + "行 4.3.15申報醫院標靶治療開始日期 欄位不得為空值\n";
                else if (data.LF4_3_15 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.15申報醫院標靶治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_4 != null)
            {
                code_check_error = Check(data.LF4_4, 0, 7);
                if (data.LF4_4 == "")
                    errortxt += "第" + listindex + "行 4.4申報醫院緩和照護 欄位不得為空值\n";
                else if (data.LF4_4 != "" && code_check_error != "OK" && data.LF4_4 != "9")
                    errortxt += "第" + listindex + "行 4.4申報醫院緩和照護 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_5_1 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Check(data.LF4_5_1, 0, 3);
                if (data.LF4_5_1 == "")
                    errortxt += "第" + listindex + "行 4.5.1其他治療 欄位不得為空值\n";
                else if (data.LF4_5_1 != "" && code_check_error != "OK" && data.LF4_5_1 != "99")
                    errortxt += "第" + listindex + "行 4.5.1其他治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF4_5_2 != null)
            {
                code_check_error = Date(data.LF4_5_2, 8);
                if (data.LF4_5_2 == "")
                    errortxt += "第" + listindex + "行 4.5.2其他治療開始日期 欄位不得為空值\n";
                else if (data.LF4_5_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.5.2其他治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF5_1 != null)
            {
                code_check_error = Date(data.LF5_1, 8);
                if (data.LF5_1 == "")
                    errortxt += "第" + listindex + "行 5.1首次復發日期 欄位不得為空值\n";
                if (data.LF5_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 5.1首次復發日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF5_2 != null)
            {
                code_check_error = Code(data.LF5_2, 2, 1);
                code_check_error = Rule_crlf5_2(data.LF5_2);
                if (data.LF5_2 == "")
                    errortxt += "第" + listindex + "行 5.2首次復發型式 欄位不得為空值\n";
                else if (data.LF5_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 5.2首次復發型式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF5_3 != null)
            {
                code_check_error = Date(data.LF5_3, 8);
                if (data.LF5_3 == "")
                    errortxt += "第" + listindex + "行 5.3最後聯絡或死亡日期 欄位不得為空值\n";
                else if (data.LF5_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 5.3最後聯絡或死亡日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF5_4 != null)
            {
                code_check_error = Check(data.LF5_4, 0, 1);
                if (data.LF5_4 == "")
                    errortxt += "第" + listindex + "行 5.4生存狀態 欄位不得為空值\n";
                else if (data.LF5_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 5.4生存狀態 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF7_1 != null)
            {
                code_check_error = Code(data.LF7_1, 3, 5);
                code_check_error = Check(data.LF7_1, 0, 999);
                if (data.LF7_1 == "")
                    errortxt += "第" + listindex + "行 7.1身高 欄位不得為空值\n";
                else if (data.LF7_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.1身高 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF7_2 != null)
            {
                code_check_error = Code(data.LF7_2, 3, 5);
                code_check_error = Check(data.LF7_2, 0, 999);
                if (data.LF7_2 == "")
                    errortxt += "第" + listindex + "行 7.2體重 欄位不得為空值\n";
                else if (data.LF7_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.2體重 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF7_3 != null)
            {
                code_check_error = Code(data.LF7_3, 6, 1);
                if (data.LF7_3 == "")
                    errortxt += "第" + listindex + "行 7.3吸菸行為 欄位不得為空值\n";
                else if (data.LF7_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.3吸菸行為 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF7_4 != null)
            {
                code_check_error = Code(data.LF7_4, 6, 1);
                if (data.LF7_4 == "")
                    errortxt += "第" + listindex + "行 7.4嚼檳榔行為 欄位不得為空值\n";
                else if (data.LF7_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.4嚼檳榔行為 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF7_5 != null)
            {
                code_check_error = Code(data.LF7_5, 3, 1);
                code_check_error = Check(data.LF7_5, 0, 4);
                if (data.LF7_5 == "")
                    errortxt += "第" + listindex + "行 7.5喝酒行為 欄位不得為空值\n";
                else if (data.LF7_5 != "" && code_check_error != "OK" && (data.LF7_5 != "009" || data.LF7_5 != "999"))
                    errortxt += "第" + listindex + "行 7.5喝酒行為 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF7_6 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Code(data.LF7_6, 3, 1);
                code_check_error = Rule_crlf7_6(data.LF7_6);
                if (data.LF7_6 == "")
                    errortxt += "第" + listindex + "行 7.6首次治療前生活功能狀態評估 欄位不得為空值\n";
                else if (data.LF7_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.6首次治療前生活功能狀態評估 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_1 != null)
            {
                code_check_error = Code(data.LF8_1, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_1 == "")
                    errortxt += "第" + listindex + "行 8.1癌症部位特定因子1 欄位不得為空值\n";
                else if (data.LF8_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.1癌症部位特定因子1 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_2 != null)
            {
                code_check_error = Code(data.LF8_2, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_2 == "")
                    errortxt += "第" + listindex + "行 8.2癌症部位特定因子2 欄位不得為空值\n";
                else if (data.LF8_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.2癌症部位特定因子2 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_3 != null)
            {
                code_check_error = Code(data.LF8_3, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_3 == "")
                    errortxt += "第" + listindex + "行 8.3癌症部位特定因子3 欄位不得為空值\n";
                else if (data.LF8_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.3癌症部位特定因子3 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_4 != null)
            {
                code_check_error = Code(data.LF8_4, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_4 == "")
                    errortxt += "第" + listindex + "行 8.4癌症部位特定因子4 欄位不得為空值\n";
                else if (data.LF8_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.4癌症部位特定因子4 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_5 != null)
            {
                code_check_error = Code(data.LF8_5, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_5 == "")
                    errortxt += "第" + listindex + "行 8.5癌症部位特定因子5 欄位不得為空值\n";
                else if (data.LF8_5 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.5癌症部位特定因子5 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_6 != null)
            {
                code_check_error = Code(data.LF8_6, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_6 == "")
                    errortxt += "第" + listindex + "行 8.6癌症部位特定因子6 欄位不得為空值\n";
                else if (data.LF8_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.6癌症部位特定因子6 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_7 != null)
            {
                code_check_error = Code(data.LF8_7, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_7 == "")
                    errortxt += "第" + listindex + "行 8.7癌症部位特定因子7 欄位不得為空值\n";
                else if (data.LF8_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.7癌症部位特定因子7 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_8 != null)
            {
                code_check_error = Code(data.LF8_8, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_8 == "")
                    errortxt += "第" + listindex + "行 8.8癌症部位特定因子8 欄位不得為空值\n";
                else if (data.LF8_8 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.8癌症部位特定因子8 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_9 != null)
            {
                code_check_error = Code(data.LF8_9, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_9 == "")
                    errortxt += "第" + listindex + "行 8.9癌症部位特定因子9 欄位不得為空值\n";
                else if (data.LF8_9 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.9癌症部位特定因子9 內容錯誤，" + code_check_error + "\n";
            }
            if (data.LF8_10 != null)
            {
                code_check_error = Code(data.LF8_10, 3, 2);
                //因為如癌有其他欄位，所以暫時不偵測
                //code_check_error = Check(data.LF8_1, 0, 999);
                if (data.LF8_10 == "")
                    errortxt += "第" + listindex + "行 8.10癌症部位特定因子10 欄位不得為空值\n";
                else if (data.LF8_10 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 8.10癌症部位特定因子10 內容錯誤，" + code_check_error + "\n";
            }

            return errortxt;
        }

        public string CRLF_TXT_Read(string path, string fileName)
        {
            string error = "";
            var CRLF_list = new List<CRLF>();

            // 讀取CRLF txt檔案
            using (var sr = new StreamReader(path))
            {
                var line = sr.ReadLine();
                //Continue to read until you reach end of file
                //CRLF 欄位及欄位長度
                var crlf_col = new Dictionary<string, int>
                {
                    {"1.1",10},{"1.2",10},{"1.3",10},{"1.4",10},{"1.5",1},{"1.6",8},{"1.7",4},
                    {"2.1",3},{"2.2",2},{"2.3",1},{"2.3.1",1},{"2.3.2",1},{"2.4",8},{"2.5",8},{"2.6",4},{"2.7",1},{"2.8",4},{"2.9",1},
                    {"2.10.1",1},{"2.10.2",1},{"2.11",1},{"2.12",8},{"2.13",3},{"2.13.1",1},{"2.13.2",1},{"2.14",2},{"2.15",2},
                    {"3.1",8},{"3.2",2},{"3.3",2},{"3.4",4},{"3.5",3},{"3.6",3},{"3.7",3},{"3.8",1},{"3.10",4},
                    {"3.11",3},{"3.12",3},{"3.13",3},{"3.14",1},{"3.16",5},{"3.17",2},{"3.19",4},{"3.21",4},
                    {"4.1",8},{"4.1.1",8},{"4.1.2",8},{"4.1.3",2},{"4.1.4",2},{"4.1.4.1",1},{"4.1.5",1},{"4.1.5.1",3},{"4.1.6",1},{"4.1.7",1},
                    {"4.1.8",1},{"4.1.9",1},{"4.1.10",1},
                    {"4.2.1.1",2},{"4.2.1.2",3},{"4.2.1.3",8},{"4.2.1.4",8},{"4.2.1.5",2},{"4.2.1.6",2},{"4.2.1.8",2},
                    {"4.2.2.1",3},{"4.2.2.2.1",2},{"4.2.2.2.2",5},{"4.2.2.2.3",2},{"4.2.2.3.1",2},{"4.2.2.3.2",5},{"4.2.2.3.3",2},
                    {"4.2.3.1",2},{"4.2.3.2",2},{"4.2.3.3.1",2},{"4.2.3.3.2",5},{"4.2.3.3.3",2},
                    {"4.3.1",8},{"4.3.2",2},{"4.3.3",2},{"4.3.4",8},{"4.3.5",2},{"4.3.6",2},{"4.3.7",8},{"4.3.8",2},{"4.3.9",2},
                    {"4.3.10",8},{"4.3.11",2},{"4.3.12",8},{"4.3.13",2},{"4.3.14",2},{"4.3.15",8},{"4.4",1},{"4.5.1",2},{"4.5.2",8},
                    {"5.1",8},{"5.2",2},{"5.3",8},{"5.4",1},
                    {"6.1",10},
                    {"7.1",3},{"7.2",3},{"7.3",6},{"7.4",6},{"7.5",3},{"7.6",3},
                    {"8.1",3},{"8.2",3},{"8.3",3},{"8.4",3},{"8.5",3},{"8.6",3},{"8.7",3},{"8.8",3},{"8.9",3},{"8.10",3}
                };

                var i = 1;
                while (line != null)
                {
                    //判斷長度是否正確
                    var crlf = new CRLF();
                    if (line.Length == 421)
                    {
                        /*
                        line.Substring(0, crlf_col[""]);
                        line = line.Substring(crlf_col[""], line.Length - crlf_col[""]);
                        */
                        crlf.LF1_1 = line.Substring(0, crlf_col["1.1"]).Trim();
                        line = line.Substring(crlf_col["1.1"], line.Length - crlf_col["1.1"]);
                        crlf.LF1_2 = line.Substring(0, crlf_col["1.2"]).Trim();
                        line = line.Substring(crlf_col["1.2"], line.Length - crlf_col["1.2"]);
                        crlf.LF1_3 = line.Substring(0, crlf_col["1.3"]).Trim();
                        line = line.Substring(crlf_col["1.3"], line.Length - crlf_col["1.3"]);
                        crlf.LF1_4 = line.Substring(0, crlf_col["1.4"]).Trim();
                        line = line.Substring(crlf_col["1.4"], line.Length - crlf_col["1.4"]);
                        crlf.LF1_5 = line.Substring(0, crlf_col["1.5"]).Trim();
                        line = line.Substring(crlf_col["1.5"], line.Length - crlf_col["1.5"]);
                        crlf.LF1_6 = line.Substring(0, crlf_col["1.6"]).Trim();
                        line = line.Substring(crlf_col["1.6"], line.Length - crlf_col["1.6"]);
                        crlf.LF1_7 = line.Substring(0, crlf_col["1.7"]).Trim();
                        line = line.Substring(crlf_col["1.7"], line.Length - crlf_col["1.7"]);
                        crlf.LF2_1 = line.Substring(0, crlf_col["2.1"]).Trim();
                        line = line.Substring(crlf_col["2.1"], line.Length - crlf_col["2.1"]);
                        crlf.LF2_2 = line.Substring(0, crlf_col["2.2"]).Trim();
                        line = line.Substring(crlf_col["2.2"], line.Length - crlf_col["2.2"]);
                        crlf.LF2_3 = line.Substring(0, crlf_col["2.3"]).Trim();
                        line = line.Substring(crlf_col["2.3"], line.Length - crlf_col["2.3"]);
                        crlf.LF2_3_1 = line.Substring(0, crlf_col["2.3.1"]).Trim();
                        line = line.Substring(crlf_col["2.3.1"], line.Length - crlf_col["2.3.1"]);
                        crlf.LF2_3_2 = line.Substring(0, crlf_col["2.3.2"]).Trim();
                        line = line.Substring(crlf_col["2.3.2"], line.Length - crlf_col["2.3.2"]);
                        crlf.LF2_4 = line.Substring(0, crlf_col["2.4"]).Trim();
                        line = line.Substring(crlf_col["2.4"], line.Length - crlf_col["2.4"]);
                        crlf.LF2_5 = line.Substring(0, crlf_col["2.5"]).Trim();
                        line = line.Substring(crlf_col["2.5"], line.Length - crlf_col["2.5"]);
                        crlf.LF2_6 = line.Substring(0, crlf_col["2.6"]).Trim();
                        line = line.Substring(crlf_col["2.6"], line.Length - crlf_col["2.6"]);
                        crlf.LF2_7 = line.Substring(0, crlf_col["2.7"]).Trim();
                        line = line.Substring(crlf_col["2.7"], line.Length - crlf_col["2.7"]);
                        crlf.LF2_8 = line.Substring(0, crlf_col["2.8"]).Trim();
                        line = line.Substring(crlf_col["2.8"], line.Length - crlf_col["2.8"]);
                        crlf.LF2_9 = line.Substring(0, crlf_col["2.9"]).Trim();
                        line = line.Substring(crlf_col["2.9"], line.Length - crlf_col["2.9"]);
                        crlf.LF2_10_1 = line.Substring(0, crlf_col["2.10.1"]).Trim();
                        line = line.Substring(crlf_col["2.10.1"], line.Length - crlf_col["2.10.1"]);
                        crlf.LF2_10_2 = line.Substring(0, crlf_col["2.10.2"]).Trim();
                        line = line.Substring(crlf_col["2.10.2"], line.Length - crlf_col["2.10.2"]);
                        crlf.LF2_11 = line.Substring(0, crlf_col["2.11"]).Trim();
                        line = line.Substring(crlf_col["2.11"], line.Length - crlf_col["2.11"]);
                        crlf.LF2_12 = line.Substring(0, crlf_col["2.12"]).Trim();
                        line = line.Substring(crlf_col["2.12"], line.Length - crlf_col["2.12"]);
                        crlf.LF2_13 = line.Substring(0, crlf_col["2.13"]).Trim();
                        line = line.Substring(crlf_col["2.13"], line.Length - crlf_col["2.13"]);
                        crlf.LF2_13_1 = line.Substring(0, crlf_col["2.13.1"]).Trim();
                        line = line.Substring(crlf_col["2.13.1"], line.Length - crlf_col["2.13.1"]);
                        crlf.LF2_13_2 = line.Substring(0, crlf_col["2.13.2"]).Trim();
                        line = line.Substring(crlf_col["2.13.2"], line.Length - crlf_col["2.13.2"]);
                        crlf.LF2_14 = line.Substring(0, crlf_col["2.14"]).Trim();
                        line = line.Substring(crlf_col["2.14"], line.Length - crlf_col["2.14"]);
                        crlf.LF2_15 = line.Substring(0, crlf_col["2.15"]).Trim();
                        line = line.Substring(crlf_col["2.15"], line.Length - crlf_col["2.15"]);
                        crlf.LF3_1 = line.Substring(0, crlf_col["3.1"]).Trim();
                        line = line.Substring(crlf_col["3.1"], line.Length - crlf_col["3.1"]);
                        crlf.LF3_2 = line.Substring(0, crlf_col["3.2"]).Trim();
                        line = line.Substring(crlf_col["3.2"], line.Length - crlf_col["3.2"]);
                        crlf.LF3_3 = line.Substring(0, crlf_col["3.3"]).Trim();
                        line = line.Substring(crlf_col["3.3"], line.Length - crlf_col["3.3"]);
                        crlf.LF3_4 = line.Substring(0, crlf_col["3.4"]).Trim();
                        line = line.Substring(crlf_col["3.4"], line.Length - crlf_col["3.4"]);
                        crlf.LF3_5 = line.Substring(0, crlf_col["3.5"]).Trim();
                        line = line.Substring(crlf_col["3.5"], line.Length - crlf_col["3.5"]);
                        crlf.LF3_6 = line.Substring(0, crlf_col["3.6"]).Trim();
                        line = line.Substring(crlf_col["3.6"], line.Length - crlf_col["3.6"]);
                        crlf.LF3_7 = line.Substring(0, crlf_col["3.7"]).Trim();
                        line = line.Substring(crlf_col["3.7"], line.Length - crlf_col["3.7"]);
                        crlf.LF3_8 = line.Substring(0, crlf_col["3.8"]).Trim();
                        line = line.Substring(crlf_col["3.8"], line.Length - crlf_col["3.8"]);
                        crlf.LF3_10 = line.Substring(0, crlf_col["3.10"]).Trim();
                        line = line.Substring(crlf_col["3.10"], line.Length - crlf_col["3.10"]);
                        crlf.LF3_11 = line.Substring(0, crlf_col["3.11"]).Trim();
                        line = line.Substring(crlf_col["3.11"], line.Length - crlf_col["3.11"]);
                        crlf.LF3_12 = line.Substring(0, crlf_col["3.12"]).Trim();
                        line = line.Substring(crlf_col["3.12"], line.Length - crlf_col["3.12"]);
                        crlf.LF3_13 = line.Substring(0, crlf_col["3.13"]).Trim();
                        line = line.Substring(crlf_col["3.13"], line.Length - crlf_col["3.13"]);
                        crlf.LF3_14 = line.Substring(0, crlf_col["3.14"]).Trim();
                        line = line.Substring(crlf_col["3.14"], line.Length - crlf_col["3.14"]);
                        crlf.LF3_16 = line.Substring(0, crlf_col["3.16"]).Trim();
                        line = line.Substring(crlf_col["3.16"], line.Length - crlf_col["3.16"]);
                        crlf.LF3_17 = line.Substring(0, crlf_col["3.17"]).Trim();
                        line = line.Substring(crlf_col["3.17"], line.Length - crlf_col["3.17"]);
                        crlf.LF3_19 = line.Substring(0, crlf_col["3.19"]).Trim();
                        line = line.Substring(crlf_col["3.19"], line.Length - crlf_col["3.19"]);
                        crlf.LF3_21 = line.Substring(0, crlf_col["3.21"]).Trim();
                        line = line.Substring(crlf_col["3.21"], line.Length - crlf_col["3.21"]);
                        crlf.LF4_1 = line.Substring(0, crlf_col["4.1"]).Trim();
                        line = line.Substring(crlf_col["4.1"], line.Length - crlf_col["4.1"]);
                        crlf.LF4_1_1 = line.Substring(0, crlf_col["4.1.1"]).Trim();
                        line = line.Substring(crlf_col["4.1.1"], line.Length - crlf_col["4.1.1"]);
                        crlf.LF4_1_2 = line.Substring(0, crlf_col["4.1.2"]).Trim();
                        line = line.Substring(crlf_col["4.1.2"], line.Length - crlf_col["4.1.2"]);
                        crlf.LF4_1_3 = line.Substring(0, crlf_col["4.1.3"]).Trim();
                        line = line.Substring(crlf_col["4.1.3"], line.Length - crlf_col["4.1.3"]);
                        crlf.LF4_1_4 = line.Substring(0, crlf_col["4.1.4"]).Trim();
                        line = line.Substring(crlf_col["4.1.4"], line.Length - crlf_col["4.1.4"]);
                        crlf.LF4_1_4_1 = line.Substring(0, crlf_col["4.1.4.1"]).Trim();
                        line = line.Substring(crlf_col["4.1.4.1"], line.Length - crlf_col["4.1.4.1"]);
                        crlf.LF4_1_5 = line.Substring(0, crlf_col["4.1.5"]).Trim();
                        line = line.Substring(crlf_col["4.1.5"], line.Length - crlf_col["4.1.5"]);
                        crlf.LF4_1_5_1 = line.Substring(0, crlf_col["4.1.5.1"]).Trim();
                        line = line.Substring(crlf_col["4.1.5.1"], line.Length - crlf_col["4.1.5.1"]);
                        crlf.LF4_1_6 = line.Substring(0, crlf_col["4.1.6"]).Trim();
                        line = line.Substring(crlf_col["4.1.6"], line.Length - crlf_col["4.1.6"]);
                        crlf.LF4_1_7 = line.Substring(0, crlf_col["4.1.7"]).Trim();
                        line = line.Substring(crlf_col["4.1.7"], line.Length - crlf_col["4.1.7"]);
                        crlf.LF4_1_8 = line.Substring(0, crlf_col["4.1.8"]).Trim();
                        line = line.Substring(crlf_col["4.1.8"], line.Length - crlf_col["4.1.8"]);
                        crlf.LF4_1_9 = line.Substring(0, crlf_col["4.1.9"]).Trim();
                        line = line.Substring(crlf_col["4.1.9"], line.Length - crlf_col["4.1.9"]);
                        crlf.LF4_1_10 = line.Substring(0, crlf_col["4.1.10"]).Trim();
                        line = line.Substring(crlf_col["4.1.10"], line.Length - crlf_col["4.1.10"]);
                        crlf.LF4_2_1_1 = line.Substring(0, crlf_col["4.2.1.1"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.1"], line.Length - crlf_col["4.2.1.1"]);
                        crlf.LF4_2_1_2 = line.Substring(0, crlf_col["4.2.1.2"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.2"], line.Length - crlf_col["4.2.1.2"]);
                        crlf.LF4_2_1_3 = line.Substring(0, crlf_col["4.2.1.3"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.3"], line.Length - crlf_col["4.2.1.3"]);
                        crlf.LF4_2_1_4 = line.Substring(0, crlf_col["4.2.1.4"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.4"], line.Length - crlf_col["4.2.1.4"]);
                        crlf.LF4_2_1_5 = line.Substring(0, crlf_col["4.2.1.5"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.5"], line.Length - crlf_col["4.2.1.5"]);
                        crlf.LF4_2_1_6 = line.Substring(0, crlf_col["4.2.1.6"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.6"], line.Length - crlf_col["4.2.1.6"]);
                        crlf.LF4_2_1_8 = line.Substring(0, crlf_col["4.2.1.8"]).Trim();
                        line = line.Substring(crlf_col["4.2.1.8"], line.Length - crlf_col["4.2.1.8"]);
                        crlf.LF4_2_2_1 = line.Substring(0, crlf_col["4.2.2.1"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.1"], line.Length - crlf_col["4.2.2.1"]);
                        crlf.LF4_2_2_2_1 = line.Substring(0, crlf_col["4.2.2.2.1"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.2.1"], line.Length - crlf_col["4.2.2.2.1"]);
                        crlf.LF4_2_2_2_2 = line.Substring(0, crlf_col["4.2.2.2.2"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.2.2"], line.Length - crlf_col["4.2.2.2.2"]);
                        crlf.LF4_2_2_2_3 = line.Substring(0, crlf_col["4.2.2.2.3"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.2.3"], line.Length - crlf_col["4.2.2.2.3"]);
                        crlf.LF4_2_2_3_1 = line.Substring(0, crlf_col["4.2.2.3.1"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.3.1"], line.Length - crlf_col["4.2.2.3.1"]);
                        crlf.LF4_2_2_3_2 = line.Substring(0, crlf_col["4.2.2.3.2"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.3.2"], line.Length - crlf_col["4.2.2.3.2"]);
                        crlf.LF4_2_2_3_3 = line.Substring(0, crlf_col["4.2.2.3.3"]).Trim();
                        line = line.Substring(crlf_col["4.2.2.3.3"], line.Length - crlf_col["4.2.2.3.3"]);
                        crlf.LF4_2_3_1 = line.Substring(0, crlf_col["4.2.3.1"]).Trim();
                        line = line.Substring(crlf_col["4.2.3.1"], line.Length - crlf_col["4.2.3.1"]);
                        crlf.LF4_2_3_2 = line.Substring(0, crlf_col["4.2.3.2"]).Trim();
                        line = line.Substring(crlf_col["4.2.3.2"], line.Length - crlf_col["4.2.3.2"]);
                        crlf.LF4_2_3_3_1 = line.Substring(0, crlf_col["4.2.3.3.1"]).Trim();
                        line = line.Substring(crlf_col["4.2.3.3.1"], line.Length - crlf_col["4.2.3.3.1"]);
                        crlf.LF4_2_3_3_2 = line.Substring(0, crlf_col["4.2.3.3.2"]).Trim();
                        line = line.Substring(crlf_col["4.2.3.3.2"], line.Length - crlf_col["4.2.3.3.2"]);
                        crlf.LF4_2_3_3_3 = line.Substring(0, crlf_col["4.2.3.3.3"]).Trim();
                        line = line.Substring(crlf_col["4.2.3.3.3"], line.Length - crlf_col["4.2.3.3.3"]);
                        crlf.LF4_3_1 = line.Substring(0, crlf_col["4.3.1"]).Trim();
                        line = line.Substring(crlf_col["4.3.1"], line.Length - crlf_col["4.3.1"]);
                        crlf.LF4_3_2 = line.Substring(0, crlf_col["4.3.2"]).Trim();
                        line = line.Substring(crlf_col["4.3.2"], line.Length - crlf_col["4.3.2"]);
                        crlf.LF4_3_3 = line.Substring(0, crlf_col["4.3.3"]).Trim();
                        line = line.Substring(crlf_col["4.3.3"], line.Length - crlf_col["4.3.3"]);
                        crlf.LF4_3_4 = line.Substring(0, crlf_col["4.3.4"]).Trim();
                        line = line.Substring(crlf_col["4.3.4"], line.Length - crlf_col["4.3.4"]);
                        crlf.LF4_3_5 = line.Substring(0, crlf_col["4.3.5"]).Trim();
                        line = line.Substring(crlf_col["4.3.5"], line.Length - crlf_col["4.3.5"]);
                        crlf.LF4_3_6 = line.Substring(0, crlf_col["4.3.6"]).Trim();
                        line = line.Substring(crlf_col["4.3.6"], line.Length - crlf_col["4.3.6"]);
                        crlf.LF4_3_7 = line.Substring(0, crlf_col["4.3.7"]).Trim();
                        line = line.Substring(crlf_col["4.3.7"], line.Length - crlf_col["4.3.7"]);
                        crlf.LF4_3_8 = line.Substring(0, crlf_col["4.3.8"]).Trim();
                        line = line.Substring(crlf_col["4.3.8"], line.Length - crlf_col["4.3.8"]);
                        crlf.LF4_3_9 = line.Substring(0, crlf_col["4.3.9"]).Trim();
                        line = line.Substring(crlf_col["4.3.9"], line.Length - crlf_col["4.3.9"]);
                        crlf.LF4_3_10 = line.Substring(0, crlf_col["4.3.10"]).Trim();
                        line = line.Substring(crlf_col["4.3.10"], line.Length - crlf_col["4.3.10"]);
                        crlf.LF4_3_11 = line.Substring(0, crlf_col["4.3.11"]).Trim();
                        line = line.Substring(crlf_col["4.3.11"], line.Length - crlf_col["4.3.11"]);
                        crlf.LF4_3_12 = line.Substring(0, crlf_col["4.3.12"]).Trim();
                        line = line.Substring(crlf_col["4.3.12"], line.Length - crlf_col["4.3.12"]);
                        crlf.LF4_3_13 = line.Substring(0, crlf_col["4.3.13"]).Trim();
                        line = line.Substring(crlf_col["4.3.13"], line.Length - crlf_col["4.3.13"]);
                        crlf.LF4_3_14 = line.Substring(0, crlf_col["4.3.14"]).Trim();
                        line = line.Substring(crlf_col["4.3.14"], line.Length - crlf_col["4.3.14"]);
                        crlf.LF4_3_15 = line.Substring(0, crlf_col["4.3.15"]).Trim();
                        line = line.Substring(crlf_col["4.3.15"], line.Length - crlf_col["4.3.15"]);
                        crlf.LF4_4 = line.Substring(0, crlf_col["4.4"]).Trim();
                        line = line.Substring(crlf_col["4.4"], line.Length - crlf_col["4.4"]);
                        crlf.LF4_5_1 = line.Substring(0, crlf_col["4.5.1"]).Trim();
                        line = line.Substring(crlf_col["4.5.1"], line.Length - crlf_col["4.5.1"]);
                        crlf.LF4_5_2 = line.Substring(0, crlf_col["4.5.2"]).Trim();
                        line = line.Substring(crlf_col["4.5.2"], line.Length - crlf_col["4.5.2"]);
                        crlf.LF5_1 = line.Substring(0, crlf_col["5.1"]).Trim();
                        line = line.Substring(crlf_col["5.1"], line.Length - crlf_col["5.1"]);
                        crlf.LF5_2 = line.Substring(0, crlf_col["5.2"]).Trim();
                        line = line.Substring(crlf_col["5.2"], line.Length - crlf_col["5.2"]);
                        crlf.LF5_3 = line.Substring(0, crlf_col["5.3"]).Trim();
                        line = line.Substring(crlf_col["5.3"], line.Length - crlf_col["5.3"]);
                        crlf.LF5_4 = line.Substring(0, crlf_col["5.4"]).Trim();
                        line = line.Substring(crlf_col["5.4"], line.Length - crlf_col["5.4"]);
                        crlf.LF6_1 = line.Substring(0, crlf_col["6.1"]).Trim();
                        line = line.Substring(crlf_col["6.1"], line.Length - crlf_col["6.1"]);
                        crlf.LF7_1 = line.Substring(0, crlf_col["7.1"]).Trim();
                        line = line.Substring(crlf_col["7.1"], line.Length - crlf_col["7.1"]);
                        crlf.LF7_2 = line.Substring(0, crlf_col["7.2"]).Trim();
                        line = line.Substring(crlf_col["7.2"], line.Length - crlf_col["7.2"]);
                        crlf.LF7_3 = line.Substring(0, crlf_col["7.3"]).Trim();
                        line = line.Substring(crlf_col["7.3"], line.Length - crlf_col["7.3"]);
                        crlf.LF7_4 = line.Substring(0, crlf_col["7.4"]).Trim();
                        line = line.Substring(crlf_col["7.4"], line.Length - crlf_col["7.4"]);
                        crlf.LF7_5 = line.Substring(0, crlf_col["7.5"]).Trim();
                        line = line.Substring(crlf_col["7.5"], line.Length - crlf_col["7.5"]);
                        crlf.LF7_6 = line.Substring(0, crlf_col["7.6"]).Trim();
                        line = line.Substring(crlf_col["7.6"], line.Length - crlf_col["7.6"]);
                        crlf.LF8_1 = line.Substring(0, crlf_col["8.1"]).Trim();
                        line = line.Substring(crlf_col["8.1"], line.Length - crlf_col["8.1"]);
                        crlf.LF8_2 = line.Substring(0, crlf_col["8.2"]).Trim();
                        line = line.Substring(crlf_col["8.2"], line.Length - crlf_col["8.2"]);
                        crlf.LF8_3 = line.Substring(0, crlf_col["8.3"]).Trim();
                        line = line.Substring(crlf_col["8.3"], line.Length - crlf_col["8.3"]);
                        crlf.LF8_4 = line.Substring(0, crlf_col["8.4"]).Trim();
                        line = line.Substring(crlf_col["8.4"], line.Length - crlf_col["8.4"]);
                        crlf.LF8_5 = line.Substring(0, crlf_col["8.5"]).Trim();
                        line = line.Substring(crlf_col["8.5"], line.Length - crlf_col["8.5"]);
                        crlf.LF8_6 = line.Substring(0, crlf_col["8.6"]).Trim();
                        line = line.Substring(crlf_col["8.6"], line.Length - crlf_col["8.6"]);
                        crlf.LF8_7 = line.Substring(0, crlf_col["8.7"]).Trim();
                        line = line.Substring(crlf_col["8.7"], line.Length - crlf_col["8.7"]);
                        crlf.LF8_8 = line.Substring(0, crlf_col["8.8"]).Trim();
                        line = line.Substring(crlf_col["8.8"], line.Length - crlf_col["8.8"]);
                        crlf.LF8_9 = line.Substring(0, crlf_col["8.9"]).Trim();
                        line = line.Substring(crlf_col["8.9"], line.Length - crlf_col["8.9"]);
                        crlf.LF8_10 = line.Substring(0, crlf_col["8.10"]).Trim();
                        line = line.Substring(crlf_col["8.10"], line.Length - crlf_col["8.10"]);
                        //Read the next line
                        error = Error_CRLF(i, crlf, error);
                    }
                    else
                    {
                        error = error + "第" + i + "行 資料長度錯誤，請確認是否符合CRLF標準長度422字元\n";
                    }

                    CRLF_list.Add(crlf);
                    i++;

                    line = sr.ReadLine();

                }
            }

            var path_csv = Server.MapPath("~/data_excel/");
            if (!Directory.Exists(path_csv))
            {
                Directory.CreateDirectory(path_csv);
            }
            // ZIP路徑
            var path_zip = Server.MapPath("~/data_excel_zip/");
            if (!Directory.Exists(path_zip))
            {
                Directory.CreateDirectory(path_zip);
            }

            // 檔名
            fileName = fileName.Substring(0, fileName.LastIndexOf(".")) + ".xlsx";
            fileZip = fileName.Substring(0, fileName.LastIndexOf(".")) + "_ZipFile.zip";


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("CRLF");

                sheet.Cells[1, 1].Value = "1.1";
                sheet.Cells[1, 2].Value = "1.4";
                sheet.Cells[1, 3].Value = "1.5";
                sheet.Cells[1, 4].Value = "1.6";
                sheet.Cells[1, 5].Value = "1.7";
                sheet.Cells[1, 6].Value = "2.1";
                sheet.Cells[1, 7].Value = "2.2";
                sheet.Cells[1, 8].Value = "2.3";
                sheet.Cells[1, 9].Value = "2.3.1";
                sheet.Cells[1, 10].Value = "2.3.2";
                sheet.Cells[1, 11].Value = "2.4";
                sheet.Cells[1, 12].Value = "2.5";
                sheet.Cells[1, 13].Value = "2.6";
                sheet.Cells[1, 14].Value = "2.7";
                sheet.Cells[1, 15].Value = "2.8";
                sheet.Cells[1, 16].Value = "2.9";
                sheet.Cells[1, 17].Value = "2.10.1";
                sheet.Cells[1, 18].Value = "2.10.2";
                sheet.Cells[1, 19].Value = "2.11";
                sheet.Cells[1, 20].Value = "2.12";
                sheet.Cells[1, 21].Value = "2.13";
                sheet.Cells[1, 22].Value = "2.13.1";
                sheet.Cells[1, 23].Value = "2.13.2";
                sheet.Cells[1, 24].Value = "2.14";
                sheet.Cells[1, 25].Value = "2.15";
                sheet.Cells[1, 26].Value = "3.1";
                sheet.Cells[1, 27].Value = "3.2";
                sheet.Cells[1, 28].Value = "3.3";
                sheet.Cells[1, 29].Value = "3.4";
                sheet.Cells[1, 30].Value = "3.5";
                sheet.Cells[1, 31].Value = "3.6";
                sheet.Cells[1, 32].Value = "3.7";
                sheet.Cells[1, 33].Value = "3.8";
                sheet.Cells[1, 34].Value = "3.10";
                sheet.Cells[1, 35].Value = "3.11";
                sheet.Cells[1, 36].Value = "3.12";
                sheet.Cells[1, 37].Value = "3.13";
                sheet.Cells[1, 38].Value = "3.14";
                sheet.Cells[1, 39].Value = "3.16";
                sheet.Cells[1, 40].Value = "3.17";
                sheet.Cells[1, 41].Value = "3.19";
                sheet.Cells[1, 42].Value = "3.21";
                sheet.Cells[1, 43].Value = "4.1";
                sheet.Cells[1, 44].Value = "4.1.1";
                sheet.Cells[1, 45].Value = "4.1.2";
                sheet.Cells[1, 46].Value = "4.1.3";
                sheet.Cells[1, 47].Value = "4.1.4";
                sheet.Cells[1, 48].Value = "4.1.4.1";
                sheet.Cells[1, 49].Value = "4.1.5";
                sheet.Cells[1, 50].Value = "4.1.5.1";
                sheet.Cells[1, 51].Value = "4.1.6";
                sheet.Cells[1, 52].Value = "4.1.7";
                sheet.Cells[1, 53].Value = "4.1.8";
                sheet.Cells[1, 54].Value = "4.1.9";
                sheet.Cells[1, 55].Value = "4.1.10";
                sheet.Cells[1, 56].Value = "4.2.1.1";
                sheet.Cells[1, 57].Value = "4.2.1.2";
                sheet.Cells[1, 58].Value = "4.2.1.3";
                sheet.Cells[1, 59].Value = "4.2.1.4";
                sheet.Cells[1, 60].Value = "4.2.1.5";
                sheet.Cells[1, 61].Value = "4.2.1.6";
                sheet.Cells[1, 62].Value = "4.2.1.8";
                sheet.Cells[1, 63].Value = "4.2.2.1";
                sheet.Cells[1, 64].Value = "4.2.2.2.1";
                sheet.Cells[1, 65].Value = "4.2.2.2.2";
                sheet.Cells[1, 66].Value = "4.2.2.2.3";
                sheet.Cells[1, 67].Value = "4.2.2.3.1";
                sheet.Cells[1, 68].Value = "4.2.2.3.2";
                sheet.Cells[1, 69].Value = "4.2.2.3.3";
                sheet.Cells[1, 70].Value = "4.2.3.1";
                sheet.Cells[1, 71].Value = "4.2.3.2";
                sheet.Cells[1, 72].Value = "4.2.3.3.1";
                sheet.Cells[1, 73].Value = "4.2.3.3.2";
                sheet.Cells[1, 74].Value = "4.2.3.3.3";
                sheet.Cells[1, 75].Value = "4.3.1";
                sheet.Cells[1, 76].Value = "4.3.2";
                sheet.Cells[1, 77].Value = "4.3.3";
                sheet.Cells[1, 78].Value = "4.3.4";
                sheet.Cells[1, 79].Value = "4.3.5";
                sheet.Cells[1, 80].Value = "4.3.6";
                sheet.Cells[1, 81].Value = "4.3.7";
                sheet.Cells[1, 82].Value = "4.3.8";
                sheet.Cells[1, 83].Value = "4.3.9";
                sheet.Cells[1, 84].Value = "4.3.10";
                sheet.Cells[1, 85].Value = "4.3.11";
                sheet.Cells[1, 86].Value = "4.3.12";
                sheet.Cells[1, 87].Value = "4.3.13";
                sheet.Cells[1, 88].Value = "4.3.14";
                sheet.Cells[1, 89].Value = "4.3.15";
                sheet.Cells[1, 90].Value = "4.4";
                sheet.Cells[1, 91].Value = "4.5.1";
                sheet.Cells[1, 92].Value = "4.5.2";
                sheet.Cells[1, 93].Value = "5.1";
                sheet.Cells[1, 94].Value = "5.2";
                sheet.Cells[1, 95].Value = "5.3";
                sheet.Cells[1, 96].Value = "5.4";
                sheet.Cells[1, 97].Value = "7.1";
                sheet.Cells[1, 98].Value = "7.2";
                sheet.Cells[1, 99].Value = "7.3";
                sheet.Cells[1, 100].Value = "7.4";
                sheet.Cells[1, 101].Value = "7.5";
                sheet.Cells[1, 102].Value = "7.6";
                sheet.Cells[1, 103].Value = "8.1";
                sheet.Cells[1, 104].Value = "8.2";
                sheet.Cells[1, 105].Value = "8.3";
                sheet.Cells[1, 106].Value = "8.4";
                sheet.Cells[1, 107].Value = "8.5";
                sheet.Cells[1, 108].Value = "8.6";
                sheet.Cells[1, 109].Value = "8.7";
                sheet.Cells[1, 110].Value = "8.8";
                sheet.Cells[1, 111].Value = "8.9";
                sheet.Cells[1, 112].Value = "8.10";

                var cell_number = 0;

                for (int i = 0; i < CRLF_list.Count; i++)
                {

                    //產生民國生日
                    string x = CRLF_list[i].LF1_6;
                    DateTime dt = DateTime.ParseExact(x, "yyyyMMdd", CultureInfo.InvariantCulture).AddYears(-1911); ;
                    string Patient_Birth = dt.ToString("yyyMMdd");
                    if (Patient_Birth.Length != 7 && Patient_Birth.Length < 7)
                    {
                        Patient_Birth = "0" + Patient_Birth;
                    }
                    patient_Hash = new Patient_Hash_Guid()
                    {
                        Patient_Id = CRLF_list[i].LF1_4,
                        Patient_Birth = Patient_Birth
                    };

                    var id_exist = patient_Hash_Guids.Find(a => a.Patient_Id == patient_Hash.Patient_Id && a.Patient_Birth == patient_Hash.Patient_Birth);
                    if (id_exist != null)
                    {
                        sheet.Cells[(i + 2 + cell_number), 1].Value = CRLF_list[i].LF1_1;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = CRLF_list[i].LF1_4;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = CRLF_list[i].LF1_5;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = CRLF_list[i].LF1_6;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = CRLF_list[i].LF1_7;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = CRLF_list[i].LF2_1;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = CRLF_list[i].LF2_2;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = CRLF_list[i].LF2_3;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = CRLF_list[i].LF2_3_1;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = CRLF_list[i].LF2_3_2;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = CRLF_list[i].LF2_4;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = CRLF_list[i].LF2_5;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = CRLF_list[i].LF2_6;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = CRLF_list[i].LF2_7;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = CRLF_list[i].LF2_8;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = CRLF_list[i].LF2_9;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = CRLF_list[i].LF2_10_1;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = CRLF_list[i].LF2_10_2;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = CRLF_list[i].LF2_11;
                        sheet.Cells[(i + 2 + cell_number), 20].Value = CRLF_list[i].LF2_12;
                        sheet.Cells[(i + 2 + cell_number), 21].Value = CRLF_list[i].LF2_13;
                        sheet.Cells[(i + 2 + cell_number), 22].Value = CRLF_list[i].LF2_13_1;
                        sheet.Cells[(i + 2 + cell_number), 23].Value = CRLF_list[i].LF2_13_2;
                        sheet.Cells[(i + 2 + cell_number), 24].Value = CRLF_list[i].LF2_14;
                        sheet.Cells[(i + 2 + cell_number), 25].Value = CRLF_list[i].LF2_15;
                        sheet.Cells[(i + 2 + cell_number), 26].Value = CRLF_list[i].LF3_1;
                        sheet.Cells[(i + 2 + cell_number), 27].Value = CRLF_list[i].LF3_2;
                        sheet.Cells[(i + 2 + cell_number), 28].Value = CRLF_list[i].LF3_3;
                        sheet.Cells[(i + 2 + cell_number), 29].Value = CRLF_list[i].LF3_4;
                        sheet.Cells[(i + 2 + cell_number), 30].Value = CRLF_list[i].LF3_5;
                        sheet.Cells[(i + 2 + cell_number), 31].Value = CRLF_list[i].LF3_6;
                        sheet.Cells[(i + 2 + cell_number), 32].Value = CRLF_list[i].LF3_7;
                        sheet.Cells[(i + 2 + cell_number), 33].Value = CRLF_list[i].LF3_8;
                        sheet.Cells[(i + 2 + cell_number), 34].Value = CRLF_list[i].LF3_10;
                        sheet.Cells[(i + 2 + cell_number), 35].Value = CRLF_list[i].LF3_11;
                        sheet.Cells[(i + 2 + cell_number), 36].Value = CRLF_list[i].LF3_12;
                        sheet.Cells[(i + 2 + cell_number), 37].Value = CRLF_list[i].LF3_13;
                        sheet.Cells[(i + 2 + cell_number), 38].Value = CRLF_list[i].LF3_14;
                        sheet.Cells[(i + 2 + cell_number), 39].Value = CRLF_list[i].LF3_16;
                        sheet.Cells[(i + 2 + cell_number), 40].Value = CRLF_list[i].LF3_17;
                        sheet.Cells[(i + 2 + cell_number), 41].Value = CRLF_list[i].LF3_19;
                        sheet.Cells[(i + 2 + cell_number), 42].Value = CRLF_list[i].LF3_21;
                        sheet.Cells[(i + 2 + cell_number), 43].Value = CRLF_list[i].LF4_1;
                        sheet.Cells[(i + 2 + cell_number), 44].Value = CRLF_list[i].LF4_1_1;
                        sheet.Cells[(i + 2 + cell_number), 45].Value = CRLF_list[i].LF4_1_2;
                        sheet.Cells[(i + 2 + cell_number), 46].Value = CRLF_list[i].LF4_1_3;
                        sheet.Cells[(i + 2 + cell_number), 47].Value = CRLF_list[i].LF4_1_4;
                        sheet.Cells[(i + 2 + cell_number), 48].Value = CRLF_list[i].LF4_1_4_1;
                        sheet.Cells[(i + 2 + cell_number), 49].Value = CRLF_list[i].LF4_1_5;
                        sheet.Cells[(i + 2 + cell_number), 50].Value = CRLF_list[i].LF4_1_5_1;
                        sheet.Cells[(i + 2 + cell_number), 51].Value = CRLF_list[i].LF4_1_6;
                        sheet.Cells[(i + 2 + cell_number), 52].Value = CRLF_list[i].LF4_1_7;
                        sheet.Cells[(i + 2 + cell_number), 53].Value = CRLF_list[i].LF4_1_8;
                        sheet.Cells[(i + 2 + cell_number), 54].Value = CRLF_list[i].LF4_1_9;
                        sheet.Cells[(i + 2 + cell_number), 55].Value = CRLF_list[i].LF4_1_10;
                        sheet.Cells[(i + 2 + cell_number), 56].Value = CRLF_list[i].LF4_2_1_1;
                        sheet.Cells[(i + 2 + cell_number), 57].Value = CRLF_list[i].LF4_2_1_2;
                        sheet.Cells[(i + 2 + cell_number), 58].Value = CRLF_list[i].LF4_2_1_3;
                        sheet.Cells[(i + 2 + cell_number), 59].Value = CRLF_list[i].LF4_2_1_4;
                        sheet.Cells[(i + 2 + cell_number), 60].Value = CRLF_list[i].LF4_2_1_5;
                        sheet.Cells[(i + 2 + cell_number), 61].Value = CRLF_list[i].LF4_2_1_6;
                        sheet.Cells[(i + 2 + cell_number), 62].Value = CRLF_list[i].LF4_2_1_8;
                        sheet.Cells[(i + 2 + cell_number), 63].Value = CRLF_list[i].LF4_2_2_1;
                        sheet.Cells[(i + 2 + cell_number), 64].Value = CRLF_list[i].LF4_2_2_2_1;
                        sheet.Cells[(i + 2 + cell_number), 65].Value = CRLF_list[i].LF4_2_2_2_2;
                        sheet.Cells[(i + 2 + cell_number), 66].Value = CRLF_list[i].LF4_2_2_2_3;
                        sheet.Cells[(i + 2 + cell_number), 67].Value = CRLF_list[i].LF4_2_2_3_1;
                        sheet.Cells[(i + 2 + cell_number), 68].Value = CRLF_list[i].LF4_2_2_3_2;
                        sheet.Cells[(i + 2 + cell_number), 69].Value = CRLF_list[i].LF4_2_2_3_3;
                        sheet.Cells[(i + 2 + cell_number), 70].Value = CRLF_list[i].LF4_2_3_1;
                        sheet.Cells[(i + 2 + cell_number), 71].Value = CRLF_list[i].LF4_2_3_2;
                        sheet.Cells[(i + 2 + cell_number), 72].Value = CRLF_list[i].LF4_2_3_3_1;
                        sheet.Cells[(i + 2 + cell_number), 73].Value = CRLF_list[i].LF4_2_3_3_2;
                        sheet.Cells[(i + 2 + cell_number), 74].Value = CRLF_list[i].LF4_2_3_3_3;
                        sheet.Cells[(i + 2 + cell_number), 75].Value = CRLF_list[i].LF4_3_1;
                        sheet.Cells[(i + 2 + cell_number), 76].Value = CRLF_list[i].LF4_3_2;
                        sheet.Cells[(i + 2 + cell_number), 77].Value = CRLF_list[i].LF4_3_3;
                        sheet.Cells[(i + 2 + cell_number), 78].Value = CRLF_list[i].LF4_3_4;
                        sheet.Cells[(i + 2 + cell_number), 79].Value = CRLF_list[i].LF4_3_5;
                        sheet.Cells[(i + 2 + cell_number), 80].Value = CRLF_list[i].LF4_3_6;
                        sheet.Cells[(i + 2 + cell_number), 81].Value = CRLF_list[i].LF4_3_7;
                        sheet.Cells[(i + 2 + cell_number), 82].Value = CRLF_list[i].LF4_3_8;
                        sheet.Cells[(i + 2 + cell_number), 83].Value = CRLF_list[i].LF4_3_9;
                        sheet.Cells[(i + 2 + cell_number), 84].Value = CRLF_list[i].LF4_3_10;
                        sheet.Cells[(i + 2 + cell_number), 85].Value = CRLF_list[i].LF4_3_11;
                        sheet.Cells[(i + 2 + cell_number), 86].Value = CRLF_list[i].LF4_3_12;
                        sheet.Cells[(i + 2 + cell_number), 87].Value = CRLF_list[i].LF4_3_13;
                        sheet.Cells[(i + 2 + cell_number), 88].Value = CRLF_list[i].LF4_3_14;
                        sheet.Cells[(i + 2 + cell_number), 89].Value = CRLF_list[i].LF4_3_15;
                        sheet.Cells[(i + 2 + cell_number), 90].Value = CRLF_list[i].LF4_4;
                        sheet.Cells[(i + 2 + cell_number), 91].Value = CRLF_list[i].LF4_5_1;
                        sheet.Cells[(i + 2 + cell_number), 92].Value = CRLF_list[i].LF4_5_2;
                        sheet.Cells[(i + 2 + cell_number), 93].Value = CRLF_list[i].LF5_1;
                        sheet.Cells[(i + 2 + cell_number), 94].Value = CRLF_list[i].LF5_2;
                        sheet.Cells[(i + 2 + cell_number), 95].Value = CRLF_list[i].LF5_3;
                        sheet.Cells[(i + 2 + cell_number), 96].Value = CRLF_list[i].LF5_4;
                        sheet.Cells[(i + 2 + cell_number), 97].Value = CRLF_list[i].LF7_1;
                        sheet.Cells[(i + 2 + cell_number), 98].Value = CRLF_list[i].LF7_2;
                        sheet.Cells[(i + 2 + cell_number), 99].Value = CRLF_list[i].LF7_3;
                        sheet.Cells[(i + 2 + cell_number), 100].Value = CRLF_list[i].LF7_4;
                        sheet.Cells[(i + 2 + cell_number), 101].Value = CRLF_list[i].LF7_5;
                        sheet.Cells[(i + 2 + cell_number), 102].Value = CRLF_list[i].LF7_6;
                        sheet.Cells[(i + 2 + cell_number), 103].Value = CRLF_list[i].LF8_1;
                        sheet.Cells[(i + 2 + cell_number), 104].Value = CRLF_list[i].LF8_2;
                        sheet.Cells[(i + 2 + cell_number), 105].Value = CRLF_list[i].LF8_3;
                        sheet.Cells[(i + 2 + cell_number), 106].Value = CRLF_list[i].LF8_4;
                        sheet.Cells[(i + 2 + cell_number), 107].Value = CRLF_list[i].LF8_5;
                        sheet.Cells[(i + 2 + cell_number), 108].Value = CRLF_list[i].LF8_6;
                        sheet.Cells[(i + 2 + cell_number), 109].Value = CRLF_list[i].LF8_7;
                        sheet.Cells[(i + 2 + cell_number), 110].Value = CRLF_list[i].LF8_8;
                        sheet.Cells[(i + 2 + cell_number), 111].Value = CRLF_list[i].LF8_9;
                        sheet.Cells[(i + 2 + cell_number), 112].Value = CRLF_list[i].LF8_10;
                    }
                    else
                    {
                        cell_number--;
                    }
                }
                p.SaveAs(new FileInfo(path_csv + fileName));
            }

            //if(error == "")
            //    db.SaveChanges();

            // xlsx壓縮ZIP
            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + fileName, fileName);
                }
            }


            ErrorReport(fileName, error);
            return "轉檔成功";
        }

        public string Error_CRSF(int listindex, CRSF data, string errortxt)
        {
            var code_check_error = "";

            if (data.SF1_1 != null)
            {
                code_check_error = Code(data.SF1_1, 10, 2);
                if (data.SF1_1 == null || data.SF1_1 == "")
                    errortxt += "第" + listindex + "行 1.1申報醫院代碼 欄位不得為空值\n";
                else if (data.SF1_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.1申報醫院代碼 內容錯誤，" + code_check_error + "\n";
            }
            //code_check_error = Code(data.SF1_2, 10, 2);
            //if (data.SF1_2 == "")
            //    errortxt += "第" + listindex + "行 1.2病歷號碼 欄位不得為空值\n";
            //else if (data.SF1_2 != "" && code_check_error != "OK")
            //    errortxt += "第" + listindex + "行 1.2病歷號碼 內容錯誤，" + code_check_error + "\n";}
            //if (data.SF1_3 == "")
            //    errortxt += "第" + listindex + "行 1.3姓名(此欄位不需要) 欄位不得為空值\n";
            if (data.SF1_4 != null)
            {
                code_check_error = Idcard(data.SF1_4, 10);
                if (data.SF1_4 == "")
                    errortxt += "第" + listindex + "行 1.4身分證統一編號 欄位不得為空值\n";
                else if (data.SF1_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.4身分證統一編號 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF1_5 != null)
            {
                code_check_error = Check(data.SF1_5, 1, 4);
                if (data.SF1_5 == "")
                    errortxt += "第" + listindex + "行 1.5性別 欄位不得為空值\n";
                else if (data.SF1_5 != "" && code_check_error != "OK" && data.SF1_5 != "9")
                    errortxt += "第" + listindex + "行 1.5性別 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF1_6 != null)
            {
                code_check_error = Date(data.SF1_6, 8); ;
                if (data.SF1_6 == "")
                    errortxt += "第" + listindex + "行 1.6出生日期 欄位不得為空值\n";
                else if (data.SF1_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.6出生日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF1_7 != null)
            {
                code_check_error = Code(data.SF1_7, 4, 5);
                if (data.SF1_7 == "")
                    errortxt += "第" + listindex + "行 1.7戶籍地代碼 欄位不得為空值\n";
                else if (data.SF1_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 1.7戶籍地代碼 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_1 != null)
            {
                code_check_error = Code(data.SF2_1, 3, 5);
                code_check_error = Check(data.SF2_1, 0, 120);
                if (data.SF2_1 == "")
                    errortxt += "第" + listindex + "行 2.1診斷年齡 欄位不得為空值\n";
                else if (data.SF2_1 != "" && code_check_error != "OK" && data.SF2_1 != "999")
                    errortxt += "第" + listindex + "行 2.1診斷年齡 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_2 != null)
            {
                code_check_error = Code(data.SF2_2, 2, 5);
                code_check_error = Check(data.SF2_2, 1, 99);
                if (data.SF2_2 == "")
                    errortxt += "第" + listindex + "行 2.2癌症發生順序號碼 欄位不得為空值\n";
                else if (data.SF2_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.2癌症發生順序號碼 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_3 != null)
            {
                code_check_error = Check(data.SF2_3, 0, 9);
                code_check_error = Rule_crlf2_3(data.SF2_3);
                if (data.SF2_3 == "")
                    errortxt += "第" + listindex + "行 2.3個案分類 欄位不得為空值\n";
                else if (data.SF2_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.3個案分類 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_3_1 != null)
            {
                code_check_error = Check(data.SF2_3_1, 1, 8);
                code_check_error = Rule_crlf2_3_1(data.SF2_3_1);
                if (data.SF2_3_1 == "")
                    errortxt += "第" + listindex + "行 2.3.1診斷狀態分類 欄位不得為空值\n";
                else if (data.SF2_3_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.3.1診斷狀態分類 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_3_2 != null)
            {
                code_check_error = Check(data.SF2_3_2, 0, 9);
                if (data.SF2_3_2 == "")
                    errortxt += "第" + listindex + "行 2.3.2治療狀態分類 欄位不得為空值\n";
                else if (data.SF2_3_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.3.2治療狀態分類 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_4 != null)
            {
                code_check_error = Date(data.SF2_4, 8);
                if (data.SF2_4 == "")
                    errortxt += "第" + listindex + "行 2.4首次就診日期 欄位不得為空值\n";
                else if (data.SF2_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.4首次就診日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_5 != null)
            {
                code_check_error = Date(data.SF2_5, 8);
                if (data.SF2_5 == "")
                    errortxt += "第" + listindex + "行 2.5最初診斷日期 欄位不得為空值\n";
                else if (data.SF2_5 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.5最初診斷日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_6 != null)
            {
                code_check_error = Code(data.SF2_6, 4, 3);
                code_check_error = Check(data.SF2_6.Substring(1, 3), 0, 809);
                if (data.SF2_6 == "")
                    errortxt += "第" + listindex + "行 2.6原發部位 欄位不得為空值\n";
                else if (data.SF2_6.Substring(0, 1) != "C")
                    errortxt += "第" + listindex + "行 2.6原發部位 內容錯誤，代碼開頭必須為C\n";
                else if (data.SF2_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.6原發部位 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_7 != null)
            {
                code_check_error = Check(data.SF2_7, 0, 5);
                if (data.SF2_7 == "")
                    errortxt += "第" + listindex + "行 2.7側性 欄位不得為空值\n";
                else if (data.SF2_7 != "" && code_check_error != "OK" && data.SF2_7 != "9")
                    errortxt += "第" + listindex + "行 2.7側性 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_8 != null)
            {
                code_check_error = Code(data.SF2_8, 4, 3);
                if (data.SF2_8 == "")
                    errortxt += "第" + listindex + "行 2.8組織類型 欄位不得為空值\n";
                else if (data.SF2_8 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.8組織類型 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_9 != null)
            {
                code_check_error = Check(data.SF2_9, 2, 3);
                if (data.SF2_9 == "")
                    errortxt += "第" + listindex + "行 2.9性態碼 欄位不得為空值\n";
                else if (data.SF2_9 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.9性態碼 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_10_1 != null)
            {
                code_check_error = Code(data.SF2_10_1, 1, 2);
                code_check_error = Rule_crlf2_10_1(data.SF2_10_1);
                if (data.SF2_10_1 == "")
                    errortxt += "第" + listindex + "行 2.10.1臨床分級/分化 欄位不得為空值\n";
                else if (data.SF2_10_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.10.1臨床分級/分化 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_10_2 != null)
            {
                code_check_error = Code(data.SF2_10_2, 1, 2);
                code_check_error = Rule_crlf2_10_1(data.SF2_10_2);
                if (data.SF2_10_2 == "")
                    errortxt += "第" + listindex + "行 2.10.2病理分級/分化 欄位不得為空值\n";
                else if (data.SF2_10_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.10.2病理分級/分化 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_11 != null)
            {
                code_check_error = Check(data.SF2_11, 1, 9);
                if (data.SF2_11 == "")
                    errortxt += "第" + listindex + "行 2.11癌症確診方式 欄位不得為空值\n";
                else if (data.SF2_11 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.11癌症確診方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF2_12 != null)
            {
                code_check_error = Date(data.SF2_12, 8);
                if (data.SF2_12 == "")
                    errortxt += "第" + listindex + "行 2.12首次顯微鏡檢證實日期 欄位不得為空值\n";
                else if (data.SF2_12 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 2.12首次顯微鏡檢證實日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_1_1 != null)
            {
                code_check_error = Date(data.SF4_1_1, 8);
                if (data.SF4_1_1 == "")
                    errortxt += "第" + listindex + "行 4.1.1首次手術日期 欄位不得為空值\n";
                else if (data.SF4_1_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.1首次手術日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_1_4 != null)
            {
                code_check_error = Code(data.SF4_1_4, 2, 1);
                code_check_error = Check(data.SF4_1_4, 0, 99);
                if (data.SF4_1_4 == "")
                    errortxt += "第" + listindex + "行 4.1.4申報醫院原發部位手術方式 欄位不得為空值\n";
                else if ((int.Parse(data.SF4_1_4) >= 1 && int.Parse(data.SF4_1_4) <= 9) || (int.Parse(data.SF4_1_4) >= 81 && int.Parse(data.SF4_1_4) <= 89) || (int.Parse(data.SF4_1_4) >= 91 && int.Parse(data.SF4_1_4) <= 97))
                    errortxt += "第" + listindex + "行 4.1.4申報醫院原發部位手術方式 內容錯誤，代碼錯誤\n";
                else if (data.SF4_1_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.1.4申報醫院原發部位手術方式 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_2_1_3 != null)
            {
                code_check_error = Date(data.SF4_2_1_3, 8);
                if (data.SF4_2_1_3 == "")
                    errortxt += "第" + listindex + "行 4.2.1.3放射治療開始日期 欄位不得為空值\n";
                else if (data.SF4_2_1_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.2.1.3放射治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_3 != null)
            {
                code_check_error = Code(data.SF4_3_3, 2, 1);
                code_check_error = Check(data.SF4_3_3, 0, 99);
                if (data.SF4_3_3 == "")
                    errortxt += "第" + listindex + "行 4.3.3申報醫院化學治療 欄位不得為空值\n";
                else if ((int.Parse(data.SF4_3_3) >= 14 && int.Parse(data.SF4_3_3) <= 19) || (int.Parse(data.SF4_3_3) >= 22 && int.Parse(data.SF4_3_3) <= 29) || (int.Parse(data.SF4_3_3) >= 32 && int.Parse(data.SF4_3_3) <= 81) || (int.Parse(data.SF4_3_3) == 84) || (int.Parse(data.SF4_3_3) >= 89 && int.Parse(data.SF4_3_3) <= 98))
                    errortxt += "第" + listindex + "行 4.3.3申報醫院化學治療 內容錯誤，代碼錯誤\n";
                else if (data.SF4_3_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.3申報醫院化學治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_4 != null)
            {
                code_check_error = Date(data.SF4_3_4, 8);
                if (data.SF4_3_4 == "")
                    errortxt += "第" + listindex + "行 4.3.4申報醫院化學治療開始日期 欄位不得為空值\n";
                else if (data.SF4_3_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.4申報醫院化學治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_6 != null)
            {
                code_check_error = Code(data.SF4_3_6, 2, 1);
                code_check_error = Check(data.SF4_3_6, 0, 99);
                if (data.SF4_3_6 == "")
                    errortxt += "第" + listindex + "行 4.3.6申報醫院賀爾蒙/類固醇治療 欄位不得為空值\n";
                else if ((int.Parse(data.SF4_3_6) >= 4 && int.Parse(data.SF4_3_6) <= 19) || (int.Parse(data.SF4_3_6) >= 22 && int.Parse(data.SF4_3_6) <= 29) || (int.Parse(data.SF4_3_6) >= 32 && int.Parse(data.SF4_3_6) <= 81) || (int.Parse(data.SF4_3_6) == 84) || (int.Parse(data.SF4_3_6) >= 89 && int.Parse(data.SF4_3_6) <= 98))
                    errortxt += "第" + listindex + "行 4.3.6申報醫院賀爾蒙/類固醇治療 內容錯誤，代碼錯誤\n";
                else if (data.SF4_3_6 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.6申報醫院賀爾蒙/類固醇治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_7 != null)
            {
                code_check_error = Date(data.SF4_3_7, 8);
                if (data.SF4_3_7 == "")
                    errortxt += "第" + listindex + "行 4.3.7申報醫院賀爾蒙/類固醇治療開始日期 欄位不得為空值\n";
                else if (data.SF4_3_7 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.7申報醫院賀爾蒙/類固醇治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_9 != null)
            {
                code_check_error = Code(data.SF4_3_9, 2, 1);
                code_check_error = Check(data.SF4_3_9, 0, 99);
                if (data.SF4_3_9 == "")
                    errortxt += "第" + listindex + "行 4.3.9申報醫院免疫治療 欄位不得為空值\n";
                else if ((int.Parse(data.SF4_3_9) >= 4 && int.Parse(data.SF4_3_9) <= 19) || (int.Parse(data.SF4_3_9) >= 22 && int.Parse(data.SF4_3_9) <= 29) || (int.Parse(data.SF4_3_9) >= 32 && int.Parse(data.SF4_3_9) <= 81) || (int.Parse(data.SF4_3_9) == 84) || (int.Parse(data.SF4_3_9) >= 89 && int.Parse(data.SF4_3_9) <= 98))
                    errortxt += "第" + listindex + "行 4.3.9申報醫院免疫治療 內容錯誤，代碼錯誤\n";
                else if (data.SF4_3_9 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.9申報醫院免疫治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_10 != null)
            {
                code_check_error = Date(data.SF4_3_10, 8);
                if (data.SF4_3_10 == "")
                    errortxt += "第" + listindex + "行 4.3.10申報醫院免疫治療開始日期 欄位不得為空值\n";
                else if (data.SF4_3_10 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.10申報醫院免疫治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_11 != null)
            {
                code_check_error = Code(data.SF4_3_11, 2, 1);
                code_check_error = Check(data.SF4_3_11, 0, 99);
                if (data.SF4_3_11 == "")
                    errortxt += "第" + listindex + "行 4.3.11骨髓/幹細胞移植或內分泌處置 欄位不得為空值\n";
                else if ((int.Parse(data.SF4_3_11) >= 1 && int.Parse(data.SF4_3_11) <= 9) || (int.Parse(data.SF4_3_11) >= 13 && int.Parse(data.SF4_3_11) <= 19) || (int.Parse(data.SF4_3_11) >= 23 && int.Parse(data.SF4_3_11) <= 24) || (int.Parse(data.SF4_3_11) >= 26 && int.Parse(data.SF4_3_11) <= 29) || (int.Parse(data.SF4_3_11) >= 31 && int.Parse(data.SF4_3_11) <= 39) || (int.Parse(data.SF4_3_11) >= 41 && int.Parse(data.SF4_3_11) <= 49) || (int.Parse(data.SF4_3_11) >= 51 && int.Parse(data.SF4_3_11) <= 81) || (int.Parse(data.SF4_3_11) == 84) || (int.Parse(data.SF4_3_11) >= 89 && int.Parse(data.SF4_3_11) <= 98))
                    errortxt += "第" + listindex + "行 4.3.11骨髓/幹細胞移植或內分泌處置 內容錯誤，代碼錯誤\n";
                else if (data.SF4_3_11 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.11骨髓/幹細胞移植或內分泌處置 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_12 != null)
            {
                code_check_error = Date(data.SF4_3_12, 8);
                if (data.SF4_3_12 == "")
                    errortxt += "第" + listindex + "行 4.3.12申報醫院骨髓/幹細胞移植或內分泌處置開始日期 欄位不得為空值\n";
                else if (data.SF4_3_12 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.12申報醫院骨髓/幹細胞移植或內分泌處置開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_14 != null)
            {
                code_check_error = Code(data.SF4_3_14, 2, 1);
                code_check_error = Check(data.SF4_3_14, 0, 99);
                if (data.SF4_3_14 == "")
                    errortxt += "第" + listindex + "行 4.3.14申報醫院標靶治療 欄位不得為空值\n";
                else if ((int.Parse(data.SF4_3_14) >= 2 && int.Parse(data.SF4_3_14) <= 19) || (int.Parse(data.SF4_3_14) >= 22 && int.Parse(data.SF4_3_14) <= 29) || (int.Parse(data.SF4_3_14) >= 32 && int.Parse(data.SF4_3_14) <= 81) || (int.Parse(data.SF4_3_14) == 84) || (int.Parse(data.SF4_3_14) >= 89 && int.Parse(data.SF4_3_14) <= 98))
                    errortxt += "第" + listindex + "行 4.3.14申報醫院標靶治療內容錯誤 內容錯誤，代碼錯誤\n";
                else if (data.SF4_3_14 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.14申報醫院標靶治療內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_3_15 != null)
            {
                code_check_error = Date(data.SF4_3_15, 8);
                if (data.SF4_3_15 == "")
                    errortxt += "第" + listindex + "行 4.3.15申報醫院標靶治療開始日期 欄位不得為空值\n";
                else if (data.SF4_3_15 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.3.15申報醫院標靶治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_4 != null)
            {
                code_check_error = Check(data.SF4_4, 0, 7);
                if (data.SF4_4 == "")
                    errortxt += "第" + listindex + "行 4.4申報醫院緩和照護 欄位不得為空值\n";
                else if (data.SF4_4 != "" && code_check_error != "OK" && data.SF4_4 != "9")
                    errortxt += "第" + listindex + "行 4.4申報醫院緩和照護 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_5_1 != null)
            {
                code_check_error = "資料格式未定義";
                code_check_error = Check(data.SF4_5_1, 0, 3);
                if (data.SF4_5_1 == "")
                    errortxt += "第" + listindex + "行 4.5.1其他治療 欄位不得為空值\n";
                else if (data.SF4_5_1 != "" && code_check_error != "OK" && data.SF4_5_1 != "99")
                    errortxt += "第" + listindex + "行 4.5.1其他治療 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF4_5_2 != null)
            {
                code_check_error = Date(data.SF4_5_2, 8);
                if (data.SF4_5_2 == "")
                    errortxt += "第" + listindex + "行 4.5.2其他治療開始日期 欄位不得為空值\n";
                else if (data.SF4_5_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 4.5.2其他治療開始日期 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF7_1 != null)
            {
                code_check_error = Code(data.SF7_1, 3, 5);
                code_check_error = Check(data.SF7_1, 0, 999);
                if (data.SF7_1 == "")
                    errortxt += "第" + listindex + "行 7.1身高 欄位不得為空值\n";
                else if (data.SF7_1 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.1身高 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF7_2 != null)
            {
                code_check_error = Code(data.SF7_2, 3, 5);
                code_check_error = Check(data.SF7_2, 0, 999);
                if (data.SF7_2 == "")
                    errortxt += "第" + listindex + "行 7.2體重 欄位不得為空值\n";
                else if (data.SF7_2 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.2體重 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF7_3 != null)
            {
                code_check_error = Code(data.SF7_3, 6, 1);
                if (data.SF7_3 == "")
                    errortxt += "第" + listindex + "行 7.3吸菸行為 欄位不得為空值\n";
                else if (data.SF7_3 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.3吸菸行為 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF7_4 != null)
            {
                code_check_error = Code(data.SF7_4, 6, 1);
                if (data.SF7_4 == "")
                    errortxt += "第" + listindex + "行 7.4嚼檳榔行為 欄位不得為空值\n";
                else if (data.SF7_4 != "" && code_check_error != "OK")
                    errortxt += "第" + listindex + "行 7.4嚼檳榔行為 內容錯誤，" + code_check_error + "\n";
            }
            if (data.SF7_5 != null)
            {
                code_check_error = Code(data.SF7_5, 3, 1);
                code_check_error = Check(data.SF7_5, 0, 4);
                if (data.SF7_5 == "")
                    errortxt += "第" + listindex + "行 7.5喝酒行為 欄位不得為空值\n";
                else if (data.SF7_5 != "" && code_check_error != "OK" && (data.SF7_5 != "009" || data.SF7_5 != "999"))
                    errortxt += "第" + listindex + "行 7.5喝酒行為 內容錯誤，" + code_check_error + "\n";
            }

            return errortxt;
        }

        public string CRSF_TXT_Read(string path, string fileName)
        {
            string error = "";
            var CRSF_list = new List<CRSF>();

            // 讀取CRLF txt檔案
            using (var sr = new StreamReader(path))
            {
                var line = sr.ReadLine();
                //Continue to read until you reach end of file
                //CRLF 欄位及欄位長度
                var crsf_col = new Dictionary<string, int>
                {
                    {"1.1",10},{"1.2",10},{"1.3",10},{"1.4",10},{"1.5",1},{"1.6",8},{"1.7",4},
                    {"2.1",3},{"2.2",2},{"2.3",1},{"2.3.1",1},{"2.3.2",1},{"2.4",8},{"2.5",8},{"2.6",4},{"2.7",1},{"2.8",4},{"2.9",1},
                    {"2.10.1",1},{"2.10.2",1},{"2.11",1},{"2.12",8},
                    {"4.1.1",8},{"4.1.4",2},
                    {"4.2.1.3",8},{"4.2.1.7",1},
                    {"4.3.3",2},{"4.3.4",8},{"4.3.6",2},{"4.3.7",8},{"4.3.9",2},
                    {"4.3.10",8},{"4.3.11",2},{"4.3.12",8},{"4.3.14",2},{"4.3.15",8},{"4.4",1},{"4.5.1",2},{"4.5.2",8},
                    {"6.1",10},
                    { "7.1",3},{"7.2",3},{"7.3",6},{"7.4",6},{"7.5",3}
                };

                var i = 1;
                while (line != null)
                {
                    //判斷長度是否正確
                    var crsf_mdata = new CRSF();
                    if (line.Length == 209)
                    {
                        /*
                        line.Substring(0, crlf_col[""]);
                        line = line.Substring(crlf_col[""], line.Length - crlf_col[""]);
                        */
                        crsf_mdata.SF1_1 = line.Substring(0, crsf_col["1.1"]).Trim();
                        line = line.Substring(crsf_col["1.1"], line.Length - crsf_col["1.1"]);
                        crsf_mdata.SF1_2 = line.Substring(0, crsf_col["1.2"]).Trim();
                        line = line.Substring(crsf_col["1.2"], line.Length - crsf_col["1.2"]);
                        crsf_mdata.SF1_3 = line.Substring(0, crsf_col["1.3"]).Trim();
                        line = line.Substring(crsf_col["1.3"], line.Length - crsf_col["1.3"]);
                        crsf_mdata.SF1_4 = line.Substring(0, crsf_col["1.4"]).Trim();
                        line = line.Substring(crsf_col["1.4"], line.Length - crsf_col["1.4"]);
                        crsf_mdata.SF1_5 = line.Substring(0, crsf_col["1.5"]).Trim();
                        line = line.Substring(crsf_col["1.5"], line.Length - crsf_col["1.5"]);
                        crsf_mdata.SF1_6 = line.Substring(0, crsf_col["1.6"]).Trim();
                        line = line.Substring(crsf_col["1.6"], line.Length - crsf_col["1.6"]);
                        crsf_mdata.SF1_7 = line.Substring(0, crsf_col["1.7"]).Trim();
                        line = line.Substring(crsf_col["1.7"], line.Length - crsf_col["1.7"]);
                        crsf_mdata.SF2_1 = line.Substring(0, crsf_col["2.1"]).Trim();
                        line = line.Substring(crsf_col["2.1"], line.Length - crsf_col["2.1"]);
                        crsf_mdata.SF2_2 = line.Substring(0, crsf_col["2.2"]).Trim();
                        line = line.Substring(crsf_col["2.2"], line.Length - crsf_col["2.2"]);
                        crsf_mdata.SF2_3 = line.Substring(0, crsf_col["2.3"]).Trim();
                        line = line.Substring(crsf_col["2.3"], line.Length - crsf_col["2.3"]);
                        crsf_mdata.SF2_3_1 = line.Substring(0, crsf_col["2.3.1"]).Trim();
                        line = line.Substring(crsf_col["2.3.1"], line.Length - crsf_col["2.3.1"]);
                        crsf_mdata.SF2_3_2 = line.Substring(0, crsf_col["2.3.2"]).Trim();
                        line = line.Substring(crsf_col["2.3.2"], line.Length - crsf_col["2.3.2"]);
                        crsf_mdata.SF2_4 = line.Substring(0, crsf_col["2.4"]).Trim();
                        line = line.Substring(crsf_col["2.4"], line.Length - crsf_col["2.4"]);
                        crsf_mdata.SF2_5 = line.Substring(0, crsf_col["2.5"]).Trim();
                        line = line.Substring(crsf_col["2.5"], line.Length - crsf_col["2.5"]);
                        crsf_mdata.SF2_6 = line.Substring(0, crsf_col["2.6"]).Trim();
                        line = line.Substring(crsf_col["2.6"], line.Length - crsf_col["2.6"]);
                        crsf_mdata.SF2_7 = line.Substring(0, crsf_col["2.7"]).Trim();
                        line = line.Substring(crsf_col["2.7"], line.Length - crsf_col["2.7"]);
                        crsf_mdata.SF2_8 = line.Substring(0, crsf_col["2.8"]).Trim();
                        line = line.Substring(crsf_col["2.8"], line.Length - crsf_col["2.8"]);
                        crsf_mdata.SF2_9 = line.Substring(0, crsf_col["2.9"]).Trim();
                        line = line.Substring(crsf_col["2.9"], line.Length - crsf_col["2.9"]);
                        crsf_mdata.SF2_10_1 = line.Substring(0, crsf_col["2.10.1"]).Trim();
                        line = line.Substring(crsf_col["2.10.1"], line.Length - crsf_col["2.10.1"]);
                        crsf_mdata.SF2_10_2 = line.Substring(0, crsf_col["2.10.2"]).Trim();
                        line = line.Substring(crsf_col["2.10.2"], line.Length - crsf_col["2.10.2"]);
                        crsf_mdata.SF2_11 = line.Substring(0, crsf_col["2.11"]).Trim();
                        line = line.Substring(crsf_col["2.11"], line.Length - crsf_col["2.11"]);
                        crsf_mdata.SF2_12 = line.Substring(0, crsf_col["2.12"]).Trim();
                        line = line.Substring(crsf_col["2.12"], line.Length - crsf_col["2.12"]);
                        crsf_mdata.SF4_1_1 = line.Substring(0, crsf_col["4.1.1"]).Trim();
                        line = line.Substring(crsf_col["4.1.1"], line.Length - crsf_col["4.1.1"]);
                        crsf_mdata.SF4_1_4 = line.Substring(0, crsf_col["4.1.4"]).Trim();
                        line = line.Substring(crsf_col["4.1.4"], line.Length - crsf_col["4.1.4"]);
                        crsf_mdata.SF4_2_1_3 = line.Substring(0, crsf_col["4.2.1.3"]).Trim();
                        line = line.Substring(crsf_col["4.2.1.3"], line.Length - crsf_col["4.2.1.3"]);
                        crsf_mdata.SF4_2_1_7 = line.Substring(0, crsf_col["4.2.1.7"]).Trim();
                        line = line.Substring(crsf_col["4.2.1.7"], line.Length - crsf_col["4.2.1.7"]);
                        crsf_mdata.SF4_3_3 = line.Substring(0, crsf_col["4.3.3"]).Trim();
                        line = line.Substring(crsf_col["4.3.3"], line.Length - crsf_col["4.3.3"]);
                        crsf_mdata.SF4_3_4 = line.Substring(0, crsf_col["4.3.4"]).Trim();
                        line = line.Substring(crsf_col["4.3.4"], line.Length - crsf_col["4.3.4"]);
                        crsf_mdata.SF4_3_6 = line.Substring(0, crsf_col["4.3.6"]).Trim();
                        line = line.Substring(crsf_col["4.3.6"], line.Length - crsf_col["4.3.6"]);
                        crsf_mdata.SF4_3_7 = line.Substring(0, crsf_col["4.3.7"]).Trim();
                        line = line.Substring(crsf_col["4.3.7"], line.Length - crsf_col["4.3.7"]);
                        crsf_mdata.SF4_3_9 = line.Substring(0, crsf_col["4.3.9"]).Trim();
                        line = line.Substring(crsf_col["4.3.9"], line.Length - crsf_col["4.3.9"]);
                        crsf_mdata.SF4_3_10 = line.Substring(0, crsf_col["4.3.10"]).Trim();
                        line = line.Substring(crsf_col["4.3.10"], line.Length - crsf_col["4.3.10"]);
                        crsf_mdata.SF4_3_11 = line.Substring(0, crsf_col["4.3.11"]).Trim();
                        line = line.Substring(crsf_col["4.3.11"], line.Length - crsf_col["4.3.11"]);
                        crsf_mdata.SF4_3_12 = line.Substring(0, crsf_col["4.3.12"]).Trim();
                        line = line.Substring(crsf_col["4.3.12"], line.Length - crsf_col["4.3.12"]);
                        crsf_mdata.SF4_3_14 = line.Substring(0, crsf_col["4.3.14"]).Trim();
                        line = line.Substring(crsf_col["4.3.14"], line.Length - crsf_col["4.3.14"]);
                        crsf_mdata.SF4_3_15 = line.Substring(0, crsf_col["4.3.15"]).Trim();
                        line = line.Substring(crsf_col["4.3.15"], line.Length - crsf_col["4.3.15"]);
                        crsf_mdata.SF4_4 = line.Substring(0, crsf_col["4.4"]).Trim();
                        line = line.Substring(crsf_col["4.4"], line.Length - crsf_col["4.4"]);
                        crsf_mdata.SF4_5_1 = line.Substring(0, crsf_col["4.5.1"]).Trim();
                        line = line.Substring(crsf_col["4.5.1"], line.Length - crsf_col["4.5.1"]);
                        crsf_mdata.SF4_5_2 = line.Substring(0, crsf_col["4.5.2"]).Trim();
                        line = line.Substring(crsf_col["4.5.2"], line.Length - crsf_col["4.5.2"]);
                        crsf_mdata.SF6_1 = line.Substring(0, crsf_col["6.1"]).Trim();
                        line = line.Substring(crsf_col["6.1"], line.Length - crsf_col["6.1"]);
                        crsf_mdata.SF7_1 = line.Substring(0, crsf_col["7.1"]).Trim();
                        line = line.Substring(crsf_col["7.1"], line.Length - crsf_col["7.1"]);
                        crsf_mdata.SF7_2 = line.Substring(0, crsf_col["7.2"]).Trim();
                        line = line.Substring(crsf_col["7.2"], line.Length - crsf_col["7.2"]);
                        crsf_mdata.SF7_3 = line.Substring(0, crsf_col["7.3"]).Trim();
                        line = line.Substring(crsf_col["7.3"], line.Length - crsf_col["7.3"]);
                        crsf_mdata.SF7_4 = line.Substring(0, crsf_col["7.4"]).Trim();
                        line = line.Substring(crsf_col["7.4"], line.Length - crsf_col["7.4"]);
                        crsf_mdata.SF7_5 = line.Substring(0, crsf_col["7.5"]).Trim();
                        //line = line.Substring(crsf_col["7.5"], line.Length - crsf_col["7.5"]);
                        //Read the next line
                        error = Error_CRSF(i, crsf_mdata, error);
                    }
                    else
                    {
                        error = error + "第" + i + "行 資料長度錯誤，請確認是否符合CRSF標準長度209字元\n";
                    }

                    CRSF_list.Add(crsf_mdata);
                    i++;

                    line = sr.ReadLine();

                }
            }

            var path_csv = Server.MapPath("~/data_excel/");
            if (!Directory.Exists(path_csv))
            {
                Directory.CreateDirectory(path_csv);
            }
            // ZIP路徑
            var path_zip = Server.MapPath("~/data_excel_zip/");
            if (!Directory.Exists(path_zip))
            {
                Directory.CreateDirectory(path_zip);
            }

            // 檔名
            fileName = fileName.Substring(0, fileName.LastIndexOf(".")) + ".xlsx";
            fileZip = fileName.Substring(0, fileName.LastIndexOf(".")) + "_ZipFile.zip";


            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet = p.Workbook.Worksheets.Add("CRSF");

                sheet.Cells[1, 1].Value = "1.1";
                sheet.Cells[1, 2].Value = "1.4";
                sheet.Cells[1, 3].Value = "1.5";
                sheet.Cells[1, 4].Value = "1.6";
                sheet.Cells[1, 5].Value = "1.7";
                sheet.Cells[1, 6].Value = "2.1";
                sheet.Cells[1, 7].Value = "2.2";
                sheet.Cells[1, 8].Value = "2.3";
                sheet.Cells[1, 9].Value = "2.3.1";
                sheet.Cells[1, 10].Value = "2.3.2";
                sheet.Cells[1, 11].Value = "2.4";
                sheet.Cells[1, 12].Value = "2.5";
                sheet.Cells[1, 13].Value = "2.6";
                sheet.Cells[1, 14].Value = "2.7";
                sheet.Cells[1, 15].Value = "2.8";
                sheet.Cells[1, 16].Value = "2.9";
                sheet.Cells[1, 17].Value = "2.10.1";
                sheet.Cells[1, 18].Value = "2.10.2";
                sheet.Cells[1, 19].Value = "2.11";
                sheet.Cells[1, 20].Value = "2.12";
                sheet.Cells[1, 21].Value = "4.1.1";
                sheet.Cells[1, 22].Value = "4.1.4";
                sheet.Cells[1, 23].Value = "4.2.1.3";
                sheet.Cells[1, 24].Value = "4.3.3";
                sheet.Cells[1, 25].Value = "4.3.4";
                sheet.Cells[1, 26].Value = "4.3.6";
                sheet.Cells[1, 27].Value = "4.3.7";
                sheet.Cells[1, 28].Value = "4.3.9";
                sheet.Cells[1, 29].Value = "4.3.10";
                sheet.Cells[1, 30].Value = "4.3.11";
                sheet.Cells[1, 31].Value = "4.3.12";
                sheet.Cells[1, 32].Value = "4.3.14";
                sheet.Cells[1, 33].Value = "4.3.15";
                sheet.Cells[1, 34].Value = "4.4";
                sheet.Cells[1, 35].Value = "4.5.1";
                sheet.Cells[1, 36].Value = "4.5.2";
                sheet.Cells[1, 37].Value = "7.1";
                sheet.Cells[1, 38].Value = "7.2";
                sheet.Cells[1, 39].Value = "7.3";
                sheet.Cells[1, 40].Value = "7.4";
                sheet.Cells[1, 41].Value = "7.5";

                var cell_number = 0;

                for (int i = 0; i < CRSF_list.Count; i++)
                {
                    //產生民國生日
                    string x = CRSF_list[i].SF1_6;
                    DateTime dt = DateTime.ParseExact(x, "yyyyMMdd", CultureInfo.InvariantCulture).AddYears(-1911); ;
                    string Patient_Birth = dt.ToString("yyyMMdd");
                    if (Patient_Birth.Length != 7 && Patient_Birth.Length < 7)
                    {
                        Patient_Birth = "0" + Patient_Birth;
                    }
                    patient_Hash = new Patient_Hash_Guid()
                    {
                        Patient_Id = CRSF_list[i].SF1_4,
                        Patient_Birth = Patient_Birth
                    };

                    var id_exist = patient_Hash_Guids.Find(a => a.Patient_Id == patient_Hash.Patient_Id && a.Patient_Birth == patient_Hash.Patient_Birth);
                    if (id_exist != null)
                    {
                        sheet.Cells[(i + 2 + cell_number), 1].Value = CRSF_list[i].SF1_1;
                        sheet.Cells[(i + 2 + cell_number), 2].Value = CRSF_list[i].SF1_4;
                        sheet.Cells[(i + 2 + cell_number), 3].Value = CRSF_list[i].SF1_5;
                        sheet.Cells[(i + 2 + cell_number), 4].Value = CRSF_list[i].SF1_6;
                        sheet.Cells[(i + 2 + cell_number), 5].Value = CRSF_list[i].SF1_7;
                        sheet.Cells[(i + 2 + cell_number), 6].Value = CRSF_list[i].SF2_1;
                        sheet.Cells[(i + 2 + cell_number), 7].Value = CRSF_list[i].SF2_2;
                        sheet.Cells[(i + 2 + cell_number), 8].Value = CRSF_list[i].SF2_3;
                        sheet.Cells[(i + 2 + cell_number), 9].Value = CRSF_list[i].SF2_3_1;
                        sheet.Cells[(i + 2 + cell_number), 10].Value = CRSF_list[i].SF2_3_2;
                        sheet.Cells[(i + 2 + cell_number), 11].Value = CRSF_list[i].SF2_4;
                        sheet.Cells[(i + 2 + cell_number), 12].Value = CRSF_list[i].SF2_5;
                        sheet.Cells[(i + 2 + cell_number), 13].Value = CRSF_list[i].SF2_6;
                        sheet.Cells[(i + 2 + cell_number), 14].Value = CRSF_list[i].SF2_7;
                        sheet.Cells[(i + 2 + cell_number), 15].Value = CRSF_list[i].SF2_8;
                        sheet.Cells[(i + 2 + cell_number), 16].Value = CRSF_list[i].SF2_9;
                        sheet.Cells[(i + 2 + cell_number), 17].Value = CRSF_list[i].SF2_10_1;
                        sheet.Cells[(i + 2 + cell_number), 18].Value = CRSF_list[i].SF2_10_2;
                        sheet.Cells[(i + 2 + cell_number), 19].Value = CRSF_list[i].SF2_11;
                        sheet.Cells[(i + 2 + cell_number), 20].Value = CRSF_list[i].SF2_12;
                        sheet.Cells[(i + 2 + cell_number), 21].Value = CRSF_list[i].SF4_1_1;
                        sheet.Cells[(i + 2 + cell_number), 22].Value = CRSF_list[i].SF4_1_4;
                        sheet.Cells[(i + 2 + cell_number), 23].Value = CRSF_list[i].SF4_2_1_3;
                        sheet.Cells[(i + 2 + cell_number), 24].Value = CRSF_list[i].SF4_3_3;
                        sheet.Cells[(i + 2 + cell_number), 25].Value = CRSF_list[i].SF4_3_4;
                        sheet.Cells[(i + 2 + cell_number), 26].Value = CRSF_list[i].SF4_3_6;
                        sheet.Cells[(i + 2 + cell_number), 27].Value = CRSF_list[i].SF4_3_7;
                        sheet.Cells[(i + 2 + cell_number), 28].Value = CRSF_list[i].SF4_3_9;
                        sheet.Cells[(i + 2 + cell_number), 29].Value = CRSF_list[i].SF4_3_10;
                        sheet.Cells[(i + 2 + cell_number), 30].Value = CRSF_list[i].SF4_3_11;
                        sheet.Cells[(i + 2 + cell_number), 31].Value = CRSF_list[i].SF4_3_12;
                        sheet.Cells[(i + 2 + cell_number), 32].Value = CRSF_list[i].SF4_3_14;
                        sheet.Cells[(i + 2 + cell_number), 33].Value = CRSF_list[i].SF4_3_15;
                        sheet.Cells[(i + 2 + cell_number), 34].Value = CRSF_list[i].SF4_4;
                        sheet.Cells[(i + 2 + cell_number), 35].Value = CRSF_list[i].SF4_5_1;
                        sheet.Cells[(i + 2 + cell_number), 36].Value = CRSF_list[i].SF4_5_2;
                        sheet.Cells[(i + 2 + cell_number), 37].Value = CRSF_list[i].SF7_1;
                        sheet.Cells[(i + 2 + cell_number), 38].Value = CRSF_list[i].SF7_2;
                        sheet.Cells[(i + 2 + cell_number), 39].Value = CRSF_list[i].SF7_3;
                        sheet.Cells[(i + 2 + cell_number), 40].Value = CRSF_list[i].SF7_4;
                        sheet.Cells[(i + 2 + cell_number), 41].Value = CRSF_list[i].SF7_5;
                    }
                    else
                    {
                        cell_number--;
                    }

                }

                p.SaveAs(new FileInfo(path_csv + fileName));
            }

            //if(error == "")
            //    db.SaveChanges();

            // xlsx壓縮ZIP
            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + fileName, fileName);
                }
            }


            ErrorReport(fileName, error);
            return "轉檔成功";
        }


        public ActionResult DownloadPartError(string fileError)
        {
            string filepath = Server.MapPath("~/data_part_error/" + fileError);
            return File(filepath, "text/html", fileError);
        }

        public void ErrorReport(string fileName, string error)
        {
            var path_error = Server.MapPath("~/data_part_error/");
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            fileError = fileName.Substring(0, fileName.LastIndexOf(".")) + "_Error.txt";

            if (error == "")
                error = "沒有錯誤";

            error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

            using (var file = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file.WriteLine(error);
            }
        }

        // ---------------------------------------------------------------------------------------------------------

        // 類別
        // text: 文字內容；min: 最小值；max: 最大值
        public string Check(string text, int min, int max)
        {
            int len = text.Length;
            if (len != 0)
                if (Regex.IsMatch(text, @"^[0-9]{" + len + "}$") && int.Parse(text) >= min && int.Parse(text) <= max)
                    return "OK";
            return "沒有符合的類別";
        }

        // 日期
        // text: 文字內容；len: 文字長度
        public string Date(string text, int len)
        {
            try
            {
                //驗證內容是否正確
                var Date_error = Code(text, len, 5);
                if (text.Trim().Length == len && Date_error == "OK")
                {
                    var year = int.Parse(text.Substring(0, 3)) + 1911;
                    var month = int.Parse(text.Substring(3, 2));
                    var day = 1;
                    var hour = 0;
                    var min = 0;
                    var sec = 0;
                    if (len == 7)
                        day = int.Parse(text.Substring(5, 2));
                    if (len == 11)
                    {
                        hour = int.Parse(text.Substring(7, 2));
                        min = int.Parse(text.Substring(9, 2));
                    }
                    if (len == 13)
                    {
                        hour = int.Parse(text.Substring(7, 2));
                        min = int.Parse(text.Substring(9, 2));
                        sec = int.Parse(text.Substring(11, 2));
                    }
                    if (len == 8)
                    {
                        year = int.Parse(text.Substring(0, 4));
                        month = int.Parse(text.Substring(4, 2));
                        day = int.Parse(text.Substring(6, 2));
                        if (text == "99999999" || text == "88888888" || text == "00000000")
                        {
                            return "OK";
                        }
                        else if (year == 9999 || month == 99 || day == 99 ||
                            year == 8888 || month == 88 || day == 88 ||
                            year == 0000 || month == 00 || day == 00)
                        {
                            return "OK";
                        }
                    }


                    new DateTime(year, month, day, hour, min, sec);
                    return "OK";
                }
                else if (text.Trim().Length < len && Date_error == "OK")
                {
                    return "日期格式不足位數需補0";
                }
                else if (text.Trim().Length > len && Date_error == "OK")
                {
                    return "日期格式超出長度";
                }
                else
                {
                    return "日期格式錯誤";
                }
            }
            catch
            {
                return "日期格式錯誤";
            }

            //return "日期格式不足位數需補0";
        }

        // 生日
        // text: 文字內容；date: 就醫日期或入院日期
        public string Birth(string text, string date)
        {
            try
            {
                DateTime input_day = new DateTime(int.Parse(text.Substring(0, 3)) + 1911, int.Parse(text.Substring(3, 2)), int.Parse(text.Substring(5, 2)));
                DateTime date1 = new DateTime(int.Parse(date.Substring(0, 3)) + 1911, int.Parse(date.Substring(3, 2)), int.Parse(date.Substring(5, 2)));
                DateTime nowdate = DateTime.Now;
                var age = int.Parse(nowdate.ToString("yyyy")) - (int.Parse(text.Substring(0, 3)) + 1911);
                if (input_day <= date1 && age < 150)
                    return "OK";
            }
            catch
            {
                return "日期格式錯誤";
            }

            return "日期必須小於等於就醫日期或入院日期且年齡必須小於150歲";
        }

        // 日期限制
        // text: 文字內容；date1: 費用年月的最後一日；date2: 入院日期
        public string Limit(string text, string date1, string date2)
        {
            try
            {
                DateTime input_day = new DateTime(int.Parse(text.Substring(0, 3)) + 1911, int.Parse(text.Substring(3, 2)), int.Parse(text.Substring(5, 2)));

                var year = int.Parse(date1.Substring(0, 3)) + 1911;
                var month = int.Parse(date1.Substring(3, 2));
                var day = DateTime.DaysInMonth(year, month);

                if (input_day <= new DateTime(year, month, day))
                {
                    if (date2 != "0")
                    {
                        if (input_day >= new DateTime(int.Parse(date2.Substring(0, 3)) + 1911, int.Parse(date2.Substring(3, 2)), int.Parse(date2.Substring(5, 2))))
                            return "OK";
                        else
                            return "日期應大於等於入院日期";
                    }
                    else
                        return "OK";
                }
                else
                    return "日期應小於等於費用年月的最後一日";
            }
            catch
            {
                return "日期格式錯誤";
            }
        }

        // 身分證 / 居留證
        // text: 文字內容；len: 文字長度
        public string Idcard(string text, int len)
        {
            string pattern = @"^[0-9A-Z]{" + len + "}$";
            if (Regex.IsMatch(text, pattern))
            {
                var check_code = "ABCDEFGHJKLMNPQRSTUVXYWZIO";
                var check_num = check_code.IndexOf(text.Substring(0, 1)) + 10;
                var foreign_num = check_code.IndexOf(text.Substring(1, 1));

                var count = (check_num / 10) * 1 + (check_num % 10) * 9;

                if (foreign_num == -1)
                    count += int.Parse(text.Substring(1, 1)) * 8;
                else
                    count += (foreign_num % 10) * 8;

                for (int i = 1; i < 8; i++)
                    count += int.Parse(text.Substring((i + 1), 1)) * (8 - i);

                if (10 - (count % 10) == int.Parse(text.Substring(9, 1)))
                    return "OK";
                else if (count % 10 == 0 && int.Parse(text.Substring(9, 1)) == 0)
                    return "OK";
            }

            return "身分證或居留證格式錯誤";
        }

        // 代碼
        // text: 文字內容；len: 文字長度
        public string Code(string text, int len, int type)
        {
            string pattern1 = @"^[0-9]{" + len + "}$";
            string pattern2 = @"^[0-9A-Z]{" + len + "}$";
            string pattern3 = @"^[0-9A-Z]*$";
            string pattern4 = @"^[0-9\.]*$";
            string pattern5 = @"^[0-9]*$";
            string pattern6 = @"^[0-9\-]*$";
            string pattern7 = @"^[0-9A-Z\.&]*$"; //用於判斷給藥頻率

            switch (type)
            {
                case 1:
                    if (Regex.IsMatch(text, pattern1))
                        return "OK";
                    break;
                case 2:
                    if (Regex.IsMatch(text, pattern2))
                        return "OK";
                    break;
                case 3:
                    if (text != null && text.Length <= len)
                        if (Regex.IsMatch(text, pattern3))
                            return "OK";
                        else
                            return "只能輸入數值和英文大寫";
                    break;
                case 4:
                    if (text != null && text.Length <= len)
                        if (Regex.IsMatch(text, pattern4))
                            return "OK";
                        else
                            return "只能輸入數值和小數點";
                    break;
                case 5:
                    if (text != null && text.Length <= len)
                        if (Regex.IsMatch(text, pattern5))
                            return "OK";
                        else
                            return "只能輸入數值";
                    break;
                case 6:
                    if (text != null && text.Length <= len)
                        if (Regex.IsMatch(text, pattern6))
                            return "OK";
                        else
                            return "只能輸入數值和負號";
                    break;
                case 7:
                    if (text != null && text.Length <= len)
                        if (Regex.IsMatch(text, pattern7))
                            return "OK";
                        else
                            return "特殊符號只接受(&.)";
                    break;
                default:
                    return "switch error";
            }

            if (text.Length > len)
                return "字串長度超過" + len + "位元(符號算1位元)";
            else if (text.Length < len)
                return "字串長度不足" + len + "位元(符號算1位元)";
            else
                return "else error";
        }

        // 錯誤_空值錯誤計算
        // text: 文字內容
        public string ErrorNull_count(string text)
        {
            MatchCollection count;
            Regex str = new Regex("欄位不得為空值");
            count = str.Matches(text);
            return count.Count.ToString();
        }

        // 錯誤_內容錯誤計算
        // text: 文字內容
        public string ErrorText_count(string text)
        {
            MatchCollection count;
            Regex str = new Regex("內容錯誤");
            count = str.Matches(text);
            return count.Count.ToString();
        }

        public string Rule_ad1(string text)
        {
            string[] rule = new string[] { "01", "02", "03", "04", "05", "06", "07", "08", "09", "11", "12", "13", "14", "15", "16", "17", "19", "21", "22", "23", "24", "25", "28", "29", "30", "31", "A1", "A2", "A3", "A5", "A6", "A7", "B1", "B6", "B7", "B8", "B9", "BA", "C1", "C4", "D1", "D2", "D4", "E1", "DF", "E2", "E3" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_ad4(string text)
        {
            string[] rule = new string[] { "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "D0", "P1", "P2", "P3", "P4", "P5", "P6", "P7", "P8", "F2", "F3", "F4", "FC", "FD", "FE", "FF", "FG", "FH", "FI", "FJ", "FK", "FL", "FM", "FN", "FS", "FT", "FU", "FV", "FX", "FY", "FZ", "L1", "L2", "L3", "L4", "L5", "L6", "L7", "L8", "L9", "LA", "LB", "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "CC", "CD", "CE", "CF", "CG", "J1", "J2", "J3", "J4", "J7", "J9", "JC", "JD", "JE", "JF", "JG", "JH", "JI", "JJ", "JK", "JL", "JM", "JN", "E1", "E2", "E4", "E5", "E6", "E8", "EA", "EB", "EC", "ED", "G4", "G5", "G6", "G8", "G9", "H1", "H2", "H3", "H4", "H6", "H7", "H8", "H9", "HA", "HB", "HC", "HD", "HE", "HF", "HG", "HH", "HI", "JA", "JB", "K1" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_ad8(string text)
        {
            string[] rule = new string[] { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "22", "23", "40", "60", "81", "82", "83", "84", "2A", "2B", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "HA", "BA", "BB", "BC", "BD", "CA", "CB", "DA", "EA", "FA", "FB", "GA", "AK" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_ap3(string text)
        {
            string[] rule = new string[] { "0", "1", "2", "3", "4", "5", "6", "8", "9", "A", "D", "E", "F", "G", "H" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_bd1(string text)
        {
            string[] rule = new string[] { "1", "2", "3", "4", "5", "6", "7", "A1", "A2", "A3", "A4", "A5", "A6", "A7", "A8", "A9", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "AJ", "AK", "AL", "AM", "AN", "AO", "AP", "AQ", "AR", "AS", "AT", "AU", "AV", "AW", "AX", "AY", "AZ", "B1", "B2", "B3", "B4", "B5", "B6", "B7", "B8", "B9", "BA", "BB", "BC", "BD", "BE", "BF", "BG", "BH", "BI", "BJ", "BK", "BL", "BM", "BN", "BO", "BP", "BQ", "BR", "BS", "BT", "BU", "BV", "BW", "BX", "BY", "BZ", "C1", "C2", "C3", "C4", "C5", "C6", "C7", "C8", "C9", "CA", "CB", "CC", "CD", "CE", "CF", "CG", "CH", "CI", "CJ", "CK", "CL", "CM", "CN", "CO", "CP", "CQ", "CR", "CS", "CT", "CU", "CV", "CW", "CX", "CY", "CZ", "D1", "D2", "D3", "D4", "D5", "D6", "D7", "D8", "D9", "DA", "DB", "DC", "DD", "DE", "DF", "DG", "DH", "DI", "DJ", "DK", "DL", "DM", "DN", "DO", "DP", "DQ", "DR", "DS", "DT", "DU", "DV", "DW", "DX", "DY", "DZ" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_bd9(string text)
        {
            string[] rule = new string[] { "00", "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12", "13", "14", "15", "22", "23", "40", "60", "81", "82", "83", "84", "2A", "2B", "AA", "AB", "AC", "AD", "AE", "AF", "AG", "AH", "AI", "BA", "BB", "BC", "BD", "CA", "CB", "DA", "EA", "FA", "FB", "GA", "AJ", "HA", "AK" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_bd24(string text)
        {
            string[] rule = new string[] { "1", "2", "3", "4", "5", "6", "7", "8", "0", "A", "B", "D", "E", "F", "G", "H", "I", "J", "K" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_bp2(string text)
        {
            string[] rule = new string[] { "1", "2", "3", "4", "7", "8", "A", "B", "C", "D", "E", "F", "G", "H", "K", "Z", "Y", "X" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_h3(string text)
        {
            string[] rule = new string[] { "11", "12", "13", "14", "15", "19", "21", "22", "29", "50" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf2_3(string text)
        {
            string[] rule = new string[] { "0", "1", "2", "3", "5", "7", "8", "9" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf2_3_1(string text)
        {
            string[] rule = new string[] { "1", "2", "3", "5", "7", "8" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf2_10_1(string text)
        {
            string[] rule = new string[] { "1", "2", "3", "4", "5", "8", "9", "A", "B", "C", "D", "E", "H", "L", "M", "S", "X" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf2_13_1(string text)
        {
            string[] rule = new string[] { "0", "1", "7", "8", "9" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_4(string text)
        {
            string[] rule = new string[] { "X", "2", "0", "2A", "A", "2A1", "IS", "2A2", "ISU", "2B", "ISD", "2C", "ISDC", "2D", "ISPA", "3", "1M", "3A", "1", "3B", "1A", "3C", "1A1", "3D", "1A2", "3E", "1B", "4", "1B1", "4A", "1B2", "4B", "1C", "4C", "1C1", "4D", "1C2", "4E", "1C3", "8888", "1D", "9999" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_5(string text)
        {
            string[] rule = new string[] { "X", "0", "0A", "0B", "1M", "1", "1A", "1B", "1C", "2", "2M", "2A", "2B", "", "2C", "3", "3A", "3B", "3C", "888", "999" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_6(string text)
        {
            string[] rule = new string[] { "X", "B", "0", "0B", "1", "1A", "1A0", "1A1", "1B", "1B0", "1B1", "1C", "1C0", "1C1", "1D", "1D0", "1D1", "1E", "888", "999" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_7(string text)
        {
            string[] rule = new string[] { "0", "2E", "0A", "2BU", "0IS", "3", "1", "3A", "1A", "3A1", "1A1", "3A2", "1A2", "3B", "1A3", "3C", "1B", "3C1", "1B1", "3C2", "1B2", "4", "1C", "4A", "1E", "4A1", "1S", "4A2", "2", "4B", "2A", "4C", "2A1", "OC", "2A2", "888", "2B", "999", "2C", "BBB" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_8(string text)
        {
            string[] rule = new string[] { "0", "3", "9" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_10(string text)
        {
            string[] rule = new string[] { "X", "2", "0", "2A", "A", "2A1", "IS", "2A2", "ISU", "2B", "ISD", "2C", "ISDC", "2D", "ISPA", "3", "1M", "3A", "1", "3B", "1A", "3C", "1A1", "3D", "1A2", "4", "1B", "4A", "1B1", "4B", "1B2", "4C", "1C", "4D", "1C1", "4E", "1C2", "8888", "1C3", "9999", "1D" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_11(string text)
        {
            string[] rule = new string[] { "X", "1M", "0", "2M", "0A", "2", "0A", "2A", "0B", "2B", "0C", "2C", "0D", "3", "1", "3A", "1A", "3B", "1AS", "3C", "1B", "888", "1C", "999" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_12(string text)
        {
            string[] rule = new string[] { "X", "0", "B", "1", "1A", "1A0", "1A1", "1B", "1B0", "1B1", "1C", "1C0", "1C1", "1D", "1D0", "1D1", "1E", "C", "CA", "CA0", "CA1", "CB", "CB0", "CB1", "CC", "CC0", "CC1", "CD", "CD0", "CD1", "CE", "888", "999" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_13(string text)
        {
            string[] rule = new string[] { "0", "3", "0A", "3A", "0IS", "3A1", "1", "3A2", "1A", "3B", "1A1", "3C", "1A2", "3C1", "1A3", "3C2", "1B", "3D", "1B1", "4", "1B2", "4A", "1C", "4A1", "1S", "4A2", "2", "4B", "2A", "4C", "2A1", "OC", "2A2", "888", "2B", "999", "2C", "BBB" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_14(string text)
        {
            string[] rule = new string[] { "0", "3", "4", "6", "9" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf3_17(string text)
        {
            string[] rule = new string[] { "00", "01", "02", "06", "07", "09", "11", "12", "13" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf4_1_5(string text)
        {
            string[] rule = new string[] { "0", "1", "2", "3", "4", "5", "7", "8", "9", "A", "B", "C", "D", "E" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf4_2_1_5(string text)
        {
            string[] rule = new string[] { "-9", "-8", "-7", "-6", "-1", "0", "1", "2", "3", "4", "5", "6", "7" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf4_2_1_6(string text)
        {
            string[] rule = new string[] { "-9", "-8", "-7", "-1", "0", "1", "2", "3", "4", "5", "6", "7" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf4_2_3_1(string text)
        {
            string[] rule = new string[] { "-9", "-1", "0", "2", "4", "8", "16", "32", "64" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf4_2_3_2(string text)
        {
            string[] rule = new string[] { "-9", "-1", "0", "1", "2", "3", "4", "5", "6", "9", "10", "12", "17", "18", "20", "33", "34", "36", "65", "66", "68", "97", "98", "99" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf5_2(string text)
        {
            string[] rule = new string[] { "00", "04", "06", "10", "13", "14", "15", "16", "17", "20", "21", "22", "25", "26", "27", "30", "36", "40", "46", "51", "52", "53", "54", "55", "56", "57", "58", "59", "60", "62", "70", "88", "99" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }
        public string Rule_crlf7_6(string text)
        {
            string[] rule = new string[] { "000", "001", "002", "003", "004", "005", "100", "104", "204", "209", "303", "304", "309", "403", "409", "502", "503", "509", "602", "609", "701", "702", "709", "801", "809", "900", "901", "909", "988", "999" };
            for (int i = 0; i < rule.Length; i++)
                if (text == rule[i])
                    return "OK";

            return "代碼錯誤";
        }


        //釋放資料庫資料
        protected override void Dispose(bool disposing)
        {
            if (disposing)
            {
                db.Dispose();
            }
            base.Dispose(disposing);
        }
    }
}