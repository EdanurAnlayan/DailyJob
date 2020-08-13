using MySql.Data.MySqlClient;
using OfficeOpenXml;
using System;
using System.Collections.Generic;
using System.Data;
using System.Globalization;
using System.IO;
using System.Net;
using System.Net.Mail;
using System.Text.Json;
using System.Text.Json.Serialization;

namespace DailyJob
{
    class Program
    {
        static string path1 = "";
        static string path2 = "";
        static MySqlConnectionStringBuilder builder;

        static double sum_foreign = 0;
        static double sum_native = 0;

        static void Main(string[] args)
        {
            builder = new MySqlConnectionStringBuilder
            {
                UserID = "userid",
                Password = "password",
                Database = "database",
                Server = "server"
            };
            string date = DateTime.Now.Date.ToString("dd MMMM yyyy dddd", CultureInfo.CreateSpecificCulture("tr-TR"));
            fillDates(date);
            var endDate = DateTime.Now.Date;
            var startDate = DateTime.Now.AddDays(-1).Date;
            DateTime epoch = DateTime.UnixEpoch;
            TimeSpan ts = startDate.Subtract(epoch);
            double startEpoch = ts.TotalMilliseconds;
            TimeSpan ts2 = endDate.Subtract(epoch);
            double endEpoch = ts2.TotalMilliseconds;
            string path = "date.json";
            var file = new FileInfo(path);
            var jsonString = File.ReadAllText(path);
            Config config = JsonSerializer.Deserialize<Config>(jsonString);
            if (!File.Exists(path))
            {
                File.Create(path);
            }
            else
            {
                if (config.startTime == "")
                {
                    MrkMuseums((long)startEpoch, (long)endEpoch);
                    OtherMuseums((long)startEpoch, (long)endEpoch);
                    sendEmail("eanlynnea@gmail.com", "excel deneme", "deneme");
                }
                else
                {
                    DateTime st = DateTime.Parse(config.startTime);
                    DateTime et = DateTime.Parse(config.endTime);
                    TimeSpan timeSpan = st.Subtract(epoch);
                    double startepoch = timeSpan.TotalMilliseconds;
                    TimeSpan timeSpan1 = et.Subtract(epoch);
                    double endepoch = timeSpan1.TotalMilliseconds;
                    string s = startepoch.ToString();
                    long sepoch = long.Parse(s.Substring(0, s.Length - 3));
                    string e = endepoch.ToString();
                    long eepoch = long.Parse(e.Substring(0, e.Length - 3));
                    List<Config> dates = new List<Config>();

                    dates.Add(new Config
                    {
                        startTime = st.ToString(),
                    });
                    foreach (var item in dates)
                    {
                        date = st.ToString("dd MMMM yyyy dddd", CultureInfo.CreateSpecificCulture("tr-TR"));

                    }
                    st = st.AddDays(+1);
                    dates.Add(new Config
                    {
                        endTime = st.ToString()
                    });
                    foreach (var item in dates)
                    {
                        DateTime sst = DateTime.Parse(config.startTime);
                        DateTime ett = sst.AddDays(+1);
                        TimeSpan timeeSpan = sst.Subtract(epoch);
                        double starttepoch = timeeSpan.TotalMilliseconds;
                        TimeSpan timeeSpan1 = ett.Subtract(epoch);
                        double enddepoch = timeeSpan1.TotalMilliseconds;
                        string ss = starttepoch.ToString();
                        long ssepoch = long.Parse(ss.Substring(0, ss.Length - 3));
                        string ee = enddepoch.ToString();
                        long eepochh = long.Parse(ee.Substring(0, ee.Length - 3));
                        fillDates(date);
                        MrkMuseums(ssepoch, eepochh);
                        OtherMuseums(ssepoch, eepochh);

                    }
                    do
                    {
                        foreach (var item in dates)
                        {
                            date = st.ToString("dd MMMM yyyy dddd", CultureInfo.CreateSpecificCulture("tr-TR"));
                            DateTime ett = st.AddDays(+1);
                            TimeSpan timeeSpan = st.Subtract(epoch);
                            double starttepoch = timeeSpan.TotalMilliseconds;
                            TimeSpan timeeSpan1 = ett.Subtract(epoch);
                            double enddepoch = timeeSpan1.TotalMilliseconds;
                            string ss = starttepoch.ToString();
                            long ssepoch = long.Parse(ss.Substring(0, ss.Length - 3));
                            string ee = enddepoch.ToString();
                            long eepochh = long.Parse(ee.Substring(0, ee.Length - 3));
                            fillDates(date);
                            MrkMuseums(sepoch, eepoch);
                            OtherMuseums(sepoch, eepoch);
                        }
                        st = st.AddDays(+1);
                        dates.Add(new Config
                        {
                            startTime = st.ToString()
                        });
                        foreach (var item in dates)
                        {
                            date = st.ToString("dd MMMM yyyy dddd", CultureInfo.CreateSpecificCulture("tr-TR"));
                            DateTime ett = st.AddDays(+1);
                            TimeSpan timeeSpan = st.Subtract(epoch);
                            double starttepoch = timeeSpan.TotalMilliseconds;
                            TimeSpan timeeSpan1 = ett.Subtract(epoch);
                            double enddepoch = timeeSpan1.TotalMilliseconds;
                            string ss = starttepoch.ToString();
                            long ssepoch = long.Parse(ss.Substring(0, ss.Length - 3));
                            string ee = enddepoch.ToString();
                            long eepochh = long.Parse(ee.Substring(0, ee.Length - 3));
                            fillDates(date);
                            MrkMuseums(sepoch, eepoch);
                            OtherMuseums(sepoch, eepoch);
                        }
                        st = st.AddDays(+1);
                        dates.Add(new Config
                        {
                            endTime = st.ToString()
                        });
                    } while (st <= et);

                }

            }
        }
        public static void fillDates(string date)
        {
            path1 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + date + " MRK.xlsx";
            path2 = Environment.GetFolderPath(Environment.SpecialFolder.Desktop) + "\\" + date + " Other.xlsx";

        }
        public static void MrkMuseums(long startDate, long endDate)
        {
            MySqlConnection con;
            con = new MySqlConnection(builder.ToString());
            con.Open();
            string query = @"query";
            MySqlCommand cmd = new MySqlCommand(query, con);
            MySqlDataReader dr = cmd.ExecuteReader();
            FillMuseumMRK(dr);
        }
        public static void OtherMuseums(long startDate, long endDate)
        {
            MySqlConnection con;
            con = new MySqlConnection(builder.ToString());
            con.Open();
            string query = @"query";
            MySqlCommand cmd = new MySqlCommand(query, con);
            MySqlDataReader dr = cmd.ExecuteReader();
            FillMuseumOther(dr);
        }
        public class Museums
        {
            public string muze_kodu { get; set; }
            public string bolum_kodu { get; set; }
            public string bolum_adi { get; set; }
            public string islem { get; set; }
            public string uyruk { get; set; }
            public string toplam_sayi { get; set; }
        }
        public static List<Museums> FillMuseumMRK(MySqlDataReader dr)
        {
            List<Museums> excelList = new List<Museums>();
            excelList.Add(new Museums
            {
                muze_kodu = "Müze Kodu",
                bolum_kodu = "Bölüm Kodu",
                bolum_adi = "Bölüm Adı",
                islem = "İşlem",
                uyruk = "Uyruk",
                toplam_sayi = "Toplam Sayı"
            });
            while (dr.Read())
            {
                excelList.Add(new Museums
                {
                    muze_kodu = dr["Müze_Kodu"].ToString(),
                    bolum_kodu = dr["Bölüm_Kodu"].ToString(),
                    bolum_adi = dr["Bölüm_Adı"].ToString(),
                    islem = dr["işlem"].ToString(),
                    uyruk = dr["uyruk"].ToString(),
                    toplam_sayi = dr["Toplam_Sayı"].ToString()

                });
            }
            dr.Close();
            var stream = new MemoryStream();
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage(stream))
            {
                ExcelWorksheet workSheet = excel.Workbook.Worksheets.Add("Yerli Yabancı Listesi");
                ExcelWorksheet workSheet1 = excel.Workbook.Worksheets.Add("Rapor");
                workSheet.Cells.LoadFromCollection(excelList);
                workSheet1.Cells[1, 1].Value = "Toplam Yerli Sayısı:";
                workSheet1.Cells[2, 1].Value = "Toplam Yabancı Sayısı: ";
                for (int row = 1; row <= workSheet.Dimension.Rows; row++)
                {
                    if (workSheet.Cells[row, 5].Value.ToString() == "Yabancı")
                    {
                        sum_foreign += double.Parse((string)workSheet.Cells[row, 6].Value);
                        workSheet1.Cells[2, 4].Value = sum_foreign;

                    }
                    else if (workSheet.Cells[row, 5].Value.ToString() == "Yerli")
                    {
                        sum_native += double.Parse((string)workSheet.Cells[row, 6].Value);
                        workSheet1.Cells[1, 4].Value = sum_native;
                    }
                }
                Stream str = File.Create(path1);
                excel.SaveAs(str);
                str.Close();
            }
            return excelList;

        }
        public static List<Museums> FillMuseumOther(MySqlDataReader dr)
        {
            List<Museums> excelList = new List<Museums>();
            excelList.Add(new Museums
            {
                muze_kodu = "Müze Kodu",
                bolum_kodu = "Bölüm Kodu",
                islem = "İşlem",
                uyruk = "Uyruk",
                toplam_sayi = "Toplam_Sayı"
            });
            while (dr.Read())
            {
                excelList.Add(new Museums
                {
                    muze_kodu = dr["Müze_Kodu"].ToString(),
                    bolum_kodu = dr["Bölüm_Kodu"].ToString(),
                    islem = dr["işlem"].ToString(),
                    uyruk = dr["uyruk"].ToString(),
                    toplam_sayi = dr["Toplam_Sayı"].ToString()

                });
            }
            dr.Close();
            var stream = new MemoryStream();

            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;
            using (var excel = new ExcelPackage(stream))
            {
                var workSheet = excel.Workbook.Worksheets.Add("Yerli Yabancı Listesi");
                ExcelWorksheet workSheet1 = excel.Workbook.Worksheets.Add("Rapor");
                workSheet.Cells.LoadFromCollection(excelList);
                workSheet1.Cells[1, 1].Value = "Toplam Yerli Sayısı:";
                workSheet1.Cells[2, 1].Value = "Toplam Yabancı Sayısı: ";
                for (int row = 1; row <= workSheet.Dimension.Rows; row++)
                {
                    if (workSheet.Cells[row, 5].Value.ToString() == "Yabancı")
                    {
                        sum_foreign += double.Parse((string)workSheet.Cells[row, 6].Value);
                        workSheet1.Cells[2, 4].Value = sum_foreign;

                    }
                    else if (workSheet.Cells[row, 5].Value.ToString() == "Yerli")
                    {
                        sum_native += double.Parse((string)workSheet.Cells[row, 6].Value);
                        workSheet1.Cells[1, 4].Value = sum_native;
                    }
                }
                Stream str = File.Create(path2);
                excel.SaveAs(str);
                str.Close();
            }
            return excelList;
        }
        public static void sendEmail(string to, string subject, string bodyText)
        {
            using (MailMessage mail = new MailMessage())
            {
                mail.From = new MailAddress("eanlynnea@gmail.com");
                mail.To.Add(to);
                mail.Subject = subject;
                mail.Body = bodyText;
                mail.IsBodyHtml = true;
                mail.Attachments.Add(new Attachment(path1));
                mail.Attachments.Add(new Attachment(path2));
                using (SmtpClient smtp = new SmtpClient("smtp.gmail.com", 587))
                {
                    smtp.Credentials = new NetworkCredential("eanlynnea@gmail.com", "password");
                    smtp.EnableSsl = true;
                    smtp.Send(mail);
                }
            }
        }
    }
}
