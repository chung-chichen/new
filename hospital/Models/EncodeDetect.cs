using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Web;

namespace hospital.Models
{
    public class EncodeDetect
    {
        //偵測byte[]是否為BIG5編碼
        public static bool IsBig5Encoding(byte[] bytes)
        {
            Encoding big5 = Encoding.GetEncoding(950);
            //將byte[]轉為string再轉回byte[]看位元數是否有變
            var a = bytes.Length;
            var b = big5.GetByteCount(big5.GetString(bytes));
            return bytes.Length ==
                big5.GetByteCount(big5.GetString(bytes));
        }
        //偵測檔案否為BIG5編碼
        public static Encoding IsBig5Encoding(string file)
        {
            //return IsBig5Encoding(File.ReadAllBytes(file));
            if (IsBig5Encoding(File.ReadAllBytes(file)))
                return Encoding.GetEncoding(950);
            else
                return Encoding.UTF8;
        }
    }
}