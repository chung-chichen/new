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
using System.Globalization;
using System.Xml.Linq;
using System.Linq;

namespace hospital.Controllers
{
    public class ConvController : Controller
    {

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
                            //data = "轉檔成功";

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
                        if (row.GetCell(0).ToString() != "" && (row.GetCell(1) != null || row.GetCell(1).ToString() != ""))
                        {
                            patient_Hash.Patient_Id = row.GetCell(0).ToString();
                            code_check_error = Regex.IsMatch(row.GetCell(0).ToString(), @"^[A-Z]{1}[A-Z1-2]{1}[0-9]{8}$") ? "OK" : "Error";
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
                            code_check_error = Regex.IsMatch(row.GetCell(1).ToString(), @"^[0-9]{7}$") ? "OK" : "Error";
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

        public string TOTFA_EXCELtransfer(string path, string XMLfileName)
        {
            try
            {
                XDocument xmlDocument = XDocument.Load(path);

                //var xElements = xmlDocument.Root?.Elements("ddata").Where(x =>
                //    !patient_Hash_Guids.Any(s => x.Element("dbody").Element("d3").Value == s.Patient_Id && x.Element("dbody").Element("d11").Value == s.Patient_Birth)).ToList();
                //xElements.Remove();
                xmlDocument.Root?.Elements("ddata").Where(x =>
                    !patient_Hash_Guids.Any(s => x.Element("dbody").Element("d3").Value == s.Patient_Id && x.Element("dbody").Element("d11").Value == s.Patient_Birth)).Remove();
                // XML路徑
                var path_XML = Server.MapPath("~/data_excel/");
                // 若資料夾不存在則建立
                if (!Directory.Exists(path_XML))
                {
                    Directory.CreateDirectory(path_XML);
                }
                // System.IO.File.Delete(path);
                xmlDocument.Save(path_XML + XMLfileName);

                ZIP(XMLfileName);

                return "轉檔成功";
            }
            catch (Exception e)
            {
                return "Error";
            }

        }

        public string TOTFB_EXCELtransfer(string path, string XMLfileName)
        {
            try
            {
                XDocument xmlDocument = XDocument.Load(path);

                xmlDocument.Root?.Elements("ddata").Where(x =>
                    !patient_Hash_Guids.Any(s => x.Element("dbody").Element("d3").Value == s.Patient_Id && x.Element("dbody").Element("d6").Value == s.Patient_Birth)).Remove();
                // XML路徑
                var path_XML = Server.MapPath("~/data_excel/");
                // 若資料夾不存在則建立
                if (!Directory.Exists(path_XML))
                {
                    Directory.CreateDirectory(path_XML);
                }
                // System.IO.File.Delete(path);
                xmlDocument.Save(path_XML + XMLfileName);

                ZIP(XMLfileName);

                return "轉檔成功";
            }
            catch (Exception e)
            {
                return "Error";
            }

        }

        //每日檢驗申報
        public string LABD_EXCELtransfer(string path, string XMLfileName)
        {
            try
            {
                XDocument xmlDocument = XDocument.Load(path);

                xmlDocument.Root?.Elements("hdata").Where(x =>
                    !patient_Hash_Guids.Any(s => x.Element("h9").Value == s.Patient_Id && x.Element("h10").Value == s.Patient_Birth)).Remove();
                // XML路徑
                var path_XML = Server.MapPath("~/data_excel/");
                // 若資料夾不存在則建立
                if (!Directory.Exists(path_XML))
                {
                    Directory.CreateDirectory(path_XML);
                }
                // System.IO.File.Delete(path);
                xmlDocument.Save(path_XML + XMLfileName);

                ZIP(XMLfileName);

                return "轉檔成功";
            }
            catch (Exception e)
            {
                return "Error";
            }

        }

        //每月檢驗申報
        public string LABM_EXCELtransfer(string path, string XMLfileName)
        {
            try
            {
                XDocument xmlDocument = XDocument.Load(path);

                xmlDocument.Root?.Elements("hdata").Where(x =>
                    !patient_Hash_Guids.Any(s => x.Element("h9").Value == s.Patient_Id && x.Element("h10").Value == s.Patient_Birth)).Remove();
                // XML路徑
                var path_XML = Server.MapPath("~/data_excel/");
                // 若資料夾不存在則建立
                if (!Directory.Exists(path_XML))
                {
                    Directory.CreateDirectory(path_XML);
                }
                // System.IO.File.Delete(path);
                xmlDocument.Save(path_XML + XMLfileName);

                ZIP(XMLfileName);

                return "轉檔成功";
            }
            catch (Exception e)
            {
                return "Error";
            }

        }

        public void ZIP(string XMLfileName)
        {
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

            // 檔名
            fileZip = XMLfileName.Substring(0, XMLfileName.Length - 4) + "_ZipFile.zip";

            // CSV壓縮ZIP
            using (FileStream file_zip = new FileStream(path_zip + fileZip, FileMode.OpenOrCreate))
            {
                using (ZipArchive archive = new ZipArchive(file_zip, ZipArchiveMode.Update))
                {
                    ZipArchiveEntry readmeEntry;
                    readmeEntry = archive.CreateEntryFromFile(path_csv + "/" + XMLfileName, XMLfileName);
                }
            }
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
        class CR_ID_BIRTH
        {
            public int start { get; set; }
            public int length { get; set; }
            public int end { get; set; }
        }

        Dictionary<string, CR_ID_BIRTH> CR_Id_Birth_Dic = new Dictionary<string, CR_ID_BIRTH>
        {
            {"ID", new CR_ID_BIRTH{start =  31, length = 10, end = 40}},
            {"BIRTH", new CR_ID_BIRTH{start =  42, length = 8, end = 49}},
        };

        public string Tw_date (string date) 
        {
            DateTime dt = DateTime.ParseExact(date, "yyyyMMdd", CultureInfo.InvariantCulture).AddYears(-1911); ;
            date = dt.ToString("yyyMMdd");
            if (date.Length != 7 && date.Length < 7)
            {
                date = "0" + date;
            }
            return date;
        }

        public string CRLF_TXT_Read(string path, string fileName)
        {
            try
            {
                // XML路徑
                var path_XML = Server.MapPath("~/data_excel/");
                // 若資料夾不存在則建立
                if (!Directory.Exists(path_XML))
                {
                    Directory.CreateDirectory(path_XML);
                }

                var tempFile = Path.GetTempFileName();
                var linesToKeep = System.IO.File.ReadLines(path).Where
                    (l =>
                    patient_Hash_Guids.Any(
                        s => s.Patient_Id == l.Substring(CR_Id_Birth_Dic["ID"].start - 1, CR_Id_Birth_Dic["ID"].length).Trim() &&
                        s.Patient_Birth == Tw_date(l.Substring(CR_Id_Birth_Dic["BIRTH"].start - 1, CR_Id_Birth_Dic["BIRTH"].length).Trim())
                        )
                    );

                System.IO.File.WriteAllLines(tempFile, linesToKeep);

                // System.IO.File.Delete(path);
                System.IO.File.Move(tempFile, path_XML + fileName);
                


                ZIP(fileName);

                return "轉檔成功";
            }
            catch (Exception e)
            {
                return "Error";
            }
        }

        public string CRSF_TXT_Read(string path, string fileName)
        {
            try
            {
                // XML路徑
                var path_XML = Server.MapPath("~/data_excel/");
                // 若資料夾不存在則建立
                if (!Directory.Exists(path_XML))
                {
                    Directory.CreateDirectory(path_XML);
                }

                var tempFile = Path.GetTempFileName();
                var linesToKeep = System.IO.File.ReadLines(path).Where
                    (l =>
                    patient_Hash_Guids.Any(
                        s => s.Patient_Id == l.Substring(CR_Id_Birth_Dic["ID"].start - 1, CR_Id_Birth_Dic["ID"].length).Trim() &&
                        s.Patient_Birth == Tw_date(l.Substring(CR_Id_Birth_Dic["BIRTH"].start - 1, CR_Id_Birth_Dic["BIRTH"].length).Trim())
                        )
                    );

                System.IO.File.WriteAllLines(tempFile, linesToKeep);

                // System.IO.File.Delete(path);
                System.IO.File.Move(tempFile, path_XML + fileName);



                ZIP(fileName);

                return "轉檔成功";
            }
            catch (Exception e)
            {
                return "Error";
            }
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
    }
}