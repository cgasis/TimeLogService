using System;
using System.Collections.Generic;
using System.Data;
using System.Data.OleDb;
using System.Globalization;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;

namespace TimeLogService
{
    public static class Helper
    {
        public static List<T> DataTableToList<T>(this DataTable table) where T : class, new()
        {
            try
            {
                List<T> list = new List<T>(table.Rows.Count);

                foreach (var row in table.AsEnumerable())
                {
                    T obj = new T();
                    foreach (var prop in obj.GetType().GetProperties())
                    {
                        try
                        {
                            PropertyInfo propertyInfo = obj.GetType().GetProperty(prop.Name);
                            var propertyCustomAttributes = propertyInfo.GetCustomAttributes().ToList()[0];
                            propertyInfo.SetValue(obj,
                                Convert.ChangeType(row[((ColumnAttribute) (propertyCustomAttributes)).Name],
                                    propertyInfo.PropertyType), null);
                        }
                        catch
                        {
                            // ignored
                        }
                    }

                    list.Add(obj);
                }

                return list;
            }
            catch { return null; }
        }
    }

    public class Attendance
    {
        [Column(Name = "Name")]
        public string Name { get; set; }

        [Column(Name = "Date/Time")]
        public DateTime DateTime { get; set; }

        [Column(Name = "Status")]
        public string Status { get; set; }

        [Column(Name = "Location ID")]
        public string Location { get; set; }

        [Column(Name = "VerifyCode")]
        public string VerifyCode { get; set; }
    }

    internal class Program
    {
        private static string AppPath
        {
            get
            {
                return System.Configuration.ConfigurationSettings.AppSettings["path"].ToString();
            }
        }

        private static string DefaultEmployeeName
        {
            get
            {
                return System.Configuration.ConfigurationSettings.AppSettings["DefaultEmployeeName"].ToString();
            }
        }

        private static void Main(string[] args)
        {
            while (true)
            {
                WriteLogs(DefaultEmployeeName);
                Console.Write("Enter Employee Name: ");
                string line = Console.ReadLine();
                if (!string.IsNullOrEmpty(line))
                    WriteLogs(line);
            }
        }

        public static string ByteArrayToHexString(byte[] Bytes)
        {
            StringBuilder Result = new StringBuilder(Bytes.Length * 2);
            string HexAlphabet = "0123456789ABCDEF";

            foreach (byte B in Bytes)
            {
                Result.Append(HexAlphabet[(int)(B >> 4)]);
                Result.Append(HexAlphabet[(int)(B & 0xF)]);
            }

            return Result.ToString();
        }

        private static void WriteLogs(string line)
        {
            if (!string.IsNullOrEmpty(line))
            {
                Console.WriteLine();
                Console.WriteLine("******************* Timelogs for {0} *********************", DefaultEmployeeName);
                Console.WriteLine();

                int intMonthLatest = 0;
                string[] folders = Directory.GetDirectories(AppPath);

                foreach (string folder in folders)
                {
                    DirectoryInfo di = new DirectoryInfo(folder);
                    DateTime dateResult;

                    if (DateTime.TryParse(string.Format("{0} 01, 2016", di.Name.Split(' ')[1]), out dateResult))
                    {
                        int m = Convert.ToDateTime(string.Format("{0} 01, 2016", di.Name.Split(' ')[1])).Month;
                        if (m == DateTime.Now.Month)
                        {
                            intMonthLatest = m;
                            break;
                        }
                    }
                }

                string monthName = new DateTime(2010, intMonthLatest, 1).ToString("MMMM", CultureInfo.CurrentCulture); 
                string path = string.Format("{0}\\{1} {2}", AppPath, intMonthLatest.ToString("0#"), monthName);
                List<string> fileNames = new List<string>();
                string[] files = Directory.GetFiles(path);

                foreach (string file in files)
                {
                    FileInfo fileInfo = new FileInfo(file);
                    fileNames.Add(fileInfo.Name);
                }

                foreach (string s in fileNames.Take(2).OrderByDescending(x => x))
                {
                    //string pref
                    Console.WriteLine();
                    Console.WriteLine(s);
                    Console.WriteLine();
                    ParseExcel(path, s);
                }
            }
        }

        private static void ParseExcel(string path, string fileName)
        {
            string sheet = fileName.Replace("ttendanceLogs.xlsx", "");
            var fle = string.Format("{0}\\{1}", path, fileName);

            string connString = "Provider=Microsoft.ACE.OLEDB.12.0;Data Source=" + fle + ";Extended Properties=Excel 12.0";

            OleDbConnection oledbConn = new OleDbConnection(connString);
            try
            {
                // Open connection
                oledbConn.Open();
                string strQuery = String.Format("select * from [{0}$]", sheet);

                // Create OleDbCommand object and select data from worksheet Sheet1
                OleDbCommand cmd = new OleDbCommand(strQuery, oledbConn);

                // Create new OleDbDataAdapter
                OleDbDataAdapter oleda = new OleDbDataAdapter();

                oleda.SelectCommand = cmd;

                // Create a DataSet which will hold the data extracted from the worksheet.
                DataSet ds = new DataSet();

                // Fill the DataSet from the data extracted from the worksheet.
                oleda.Fill(ds, "sheetdt");

                List<Attendance> attendance = ds.Tables[0].DataTableToList<Attendance>().Where(x => x.Name.Equals(DefaultEmployeeName)).OrderByDescending(d => d.DateTime).ToList();

                DateTime currentDay;

                currentDay = DateTime.Now.AddDays(1);
                bool isDay = false;

                foreach (Attendance a in attendance)
                {
                    //Console.WriteLine(string.Format("{0}: {1}, {2}", a.Name, a.DateTime, a.Status));

                    if (currentDay.Day != a.DateTime.Day)
                    {
                        currentDay = a.DateTime;
                        isDay = false;
                    }
                    else
                    {
                        if (!isDay)
                        {
                            var query = attendance.FindAll(x => x.DateTime.Day == currentDay.Day);
                            var first = query.First();
                            var last = query.Last();

                            Console.WriteLine("{0} {1}", last.DateTime, last.Status);
                            Console.WriteLine("         {0} {1}", first.DateTime.ToLongTimeString(), first.Status);
                            if (query.Any(x => x.Status.Equals("C/Out")))
                            {
                                Console.WriteLine("         Total Hours Worked: {0}", Math.Round((first.DateTime - last.DateTime).TotalHours, 2));
                            }
                            ParseBreak(query.OrderBy(x => x.DateTime.ToLocalTime()).ToList());
                            isDay = true;
                        }
                    }

                    //Console.WriteLine(string.Format("{0}, {1}", query.Count(), a.DateTime));
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.Message);
            }
            finally
            {
                // Close connection
                oledbConn.Close();
                oledbConn.Dispose();
            }
        }

        public static void ParseBreak(List<Attendance> query)
        {
            double intBreak = 0;
            bool isOut = false;
            Console.WriteLine();
            for (int a = 0; a < query.Count; a++)
            {
                if (query[a].Status.Equals("C/In"))
                {
                    if (isOut)
                    {
                        intBreak += (query[a - 1].DateTime - query[a].DateTime).TotalMinutes;

                        Console.WriteLine();
                        Console.WriteLine("         {0} {1}", query[a - 1].DateTime.ToLongTimeString(), query[a - 1].Status);
                        Console.WriteLine("         {0} {1}", query[a].DateTime.ToLongTimeString(), query[a].Status);
                        Console.WriteLine("         break: {0} min(s)", Math.Round((query[a - 1].DateTime - query[a].DateTime).TotalMinutes, 2));
                        Console.WriteLine();
                    }

                    isOut = false;
                }

                if (query[a].Status.Equals("C/Out"))
                {
                    isOut = true;
                }
            }

            if (!query.Any(x => x.Status.Equals("C/Out")))
            {
                Console.WriteLine("         **No out**");
                Console.WriteLine();
            }

            Console.WriteLine("         Total Break Min(s): {0}", Math.Round(intBreak, 2));
            Console.WriteLine();
        }
    }
}