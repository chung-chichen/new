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
                            //data = GUID_Converter(workbook, fileName_excel);
                            data = "轉檔成功";

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

        public string TOTFA_EXCELtransfer(string path, string XMLfileName)
        {
            TOTFA_xml(path);
            string report = "";

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

        public string TOTFB_EXCELtransfer(string path, string XMLfileName)
        {
            TOTFB_xml(path);
            //string report = TOTFB_error(path, XMLfileName);
            string report = "";
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

        //每日檢驗申報
        public string LABD_EXCELtransfer(string path, string XMLfileName)
        {
            LABD_xml(path);
            string report = "";


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
            //string report = LABM_error(path, XMLfileName);
            string report = "";

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
                            //data = GUID_Converter(workbook, fileName_excel);
                            data = "轉檔成功";

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
                        //error = Error_CRLF(i, crlf, error);
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
                        //error = Error_CRSF(i, crsf_mdata, error);
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