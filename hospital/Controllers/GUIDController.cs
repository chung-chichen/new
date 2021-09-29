using hospital.Models;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Security.Cryptography;
using System.Text;
using System.Text.RegularExpressions;
using System.Threading.Tasks;
using System.Web;
using System.Web.Mvc;

namespace hospital.Controllers
{
    public class GUIDController : Controller
    {
        //連資料庫
        //private Patient_Guid db = new Patient_Guid();

        // 變數
        string alert = "0"; // 提醒 0為錯誤 1為成功
        string result = ""; // 上傳結果訊息
        string data = ""; // 轉檔結果訊息
        string fileZip = ""; // ZIP檔名
        string fileError = ""; // Error檔名

        // GET: GUID
        public ActionResult Index(string form, string alert, string result, string fileZip, string fileError)
        {
            ViewBag.form = form;
            ViewBag.alert = alert;
            ViewBag.result = result;
            ViewBag.fileZip = fileZip;
            ViewBag.fileError = fileError;
            return View();
        }

        [HttpPost]
        public ActionResult Index(HttpPostedFileBase file)
        {
            alert = "0";
            result = "";
            data = "";
            fileZip = "";
            fileError = "";
            if (file != null)
            {
                if (file.ContentLength > 0)
                {
                    string extension = Path.GetExtension(file.FileName);
                    if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase) || extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
                    {
                        var fileName = Path.GetFileName(file.FileName);
                        var path = Server.MapPath("~/GUID_Upload/");
                        if (!Directory.Exists(path))
                        {
                            Directory.CreateDirectory(path);
                        }
                        DateTime myDate = DateTime.Now;
                        fileName = myDate.ToString("yyyyMMddHHmmss") + "_" + fileName;
                        var path_upload = Path.Combine(path, fileName);
                        file.SaveAs(path_upload);

                        IWorkbook workbook;
                        string filepath = Server.MapPath("~/GUID_Upload/" + fileName);

                        using (FileStream fileStream = new FileStream(filepath, FileMode.Open, FileAccess.Read))
                        {
                            if (extension.Equals(".xls", StringComparison.OrdinalIgnoreCase))
                            {
                                workbook = new HSSFWorkbook(fileStream);
                            }
                            else if (extension.Equals(".xlsx", StringComparison.OrdinalIgnoreCase))
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
                            data = GUID_Converter(workbook, fileName);

                            if (data == "轉檔成功")
                            {
                                alert = "1";
                                result = "檔案已上傳成功！";
                            }
                            else
                            {
                                alert = "0";
                                result = "上傳檔案內容錯誤！";
                            }
                        }
                        else
                        {
                            alert = "0";
                            result = "上傳檔案內容錯誤！";
                        }
                    }
                    else
                    {
                        alert = "0";
                        result = "只能接受EXCEL檔案！";
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
            return RedirectToAction("Index", new { form = "Index", alert = alert, result = result, fileZip = fileZip, fileError = fileError });
        }

        public string GUID_Converter(IWorkbook workbook, string fileName)
        {
            //DataTable table = new DataTable();

            ISheet sheet = workbook.GetSheetAt(0);
            IRow headerRow = sheet.GetRow(0);
            int cellCount = headerRow.LastCellNum;
            int rowCount = sheet.LastRowNum;
            string error = "";
            var code_check_error = "";
            Patient_Hash_Guid patient_Hash = new Patient_Hash_Guid();
            List<Patient_Hash_Guid> patient_Hash_Guids = new List<Patient_Hash_Guid>();

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
                    //出生年月日 不足位數補零
                    if (row.GetCell(1) != null)
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

                    patient_Hash_Guids.Add(patient_Hash);
                }
                else
                {
                    patient_Hash.Patient_Id = row.GetCell(0).ToString();
                    patient_Hash.Patient_Birth = row.GetCell(1).ToString();
                    //將欄位名稱加入 List
                    patient_Hash_Guids.Add(patient_Hash);
                }
            }

            //產生excel檔案
            // excel路徑
            var path_excel = Server.MapPath("~/Patint_ID_excel/");
            // 若資料夾不存在則建立
            if (!Directory.Exists(path_excel))
            {
                Directory.CreateDirectory(path_excel);
            }
            var excel_name = fileName.Split('.')[0] + "_Ha.xlsx";
            fileZip = excel_name;

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (ExcelPackage p = new ExcelPackage())
            {
                ExcelWorksheet sheet_N = p.Workbook.Worksheets.Add("Hash");

                for (int i = 0; i < patient_Hash_Guids.Count; i++)
                {
                    sheet_N.Cells[(i + 1), 1].Value = patient_Hash_Guids[i].Patient_Id;
                    sheet_N.Cells[(i + 1), 2].Value = patient_Hash_Guids[i].Patient_Birth;
                }

                p.SaveAs(new FileInfo(path_excel + excel_name));
            }


            var path_error = Server.MapPath("~/data_error/");
            if (!Directory.Exists(path_error))
            {
                Directory.CreateDirectory(path_error);
            }

            var date = fileName.Substring(0, 15);
            fileError = fileName.Substring(0, fileName.Length - 4) + "_Error.txt";

            if (error == "")
                error = "沒有錯誤";

            error = "欄位不得為空值有" + ErrorNull_count(error) + "個\n內容錯誤有" + ErrorText_count(error) + "個\n\n以下為詳細報告\n" + error;

            using (var file = new StreamWriter(path_error + fileError, false, System.Text.Encoding.UTF8))
            {
                file.WriteLine(error);
            }

            return "轉檔成功";
        }


        // 下載Excel檔案
        public ActionResult DownloadExcel(string fileZip)
        {
            // 下載的檔案位置
            string filepath = Server.MapPath("~/Patint_GUID_excel/" + fileZip);
            // 回傳檔案名稱
            return File(filepath, "application/octet-stream", fileZip);
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

        // text: 文字內容；len: 文字長度
        public string Date(string text, int len)
        {
            if (text.Trim().Length == len)
            {
                var year = int.Parse(text.Substring(0, 3)) + 1911;
                var month = int.Parse(text.Substring(3, 2));
                var day = 1;
                var hour = 0;
                var min = 0;
                if (len == 7)
                    day = int.Parse(text.Substring(5, 2));
                if (len == 11)
                {
                    hour = int.Parse(text.Substring(7, 2));
                    min = int.Parse(text.Substring(9, 2));
                }
                if (len == 8)
                {
                    year = int.Parse(text.Substring(0, 4));
                    month = int.Parse(text.Substring(4, 2));
                    day = int.Parse(text.Substring(6, 2));
                    if (text == "00000000" || text == "99999999")
                        return "OK";
                }

                try
                {
                    new DateTime(year, month, day, hour, min, 0);
                    return "OK";
                }
                catch
                {
                    return "日期格式錯誤";
                }
            }

            return "日期格式不足位數需補0";
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



        //protected override void Dispose(bool disposing)
        //{
        //    if (disposing)
        //    {
        //        db.Dispose();
        //    }
        //    base.Dispose(disposing);
        //}
    }
}