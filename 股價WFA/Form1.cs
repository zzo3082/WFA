using Newtonsoft.Json.Linq;
using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Data.SqlClient;
using System.Threading;
using System.IO;

namespace 股價WFA
{
    public partial class Form1 : Form
    {
        public Form1()
        {
            InitializeComponent();
        }

        // 日期
        List<string> datesTW = new List<string>();
        List<string> datesUS = new List<string>();

        // 寫入上櫃資料1年
        private void button1_Click(object sender, EventArgs e)
        {
            SqlConnection cn = new SqlConnection(@"Data Source=.\sqlexpress;Initial Catalog=Lab;Integrated Security=True");
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cn.Open();
            int cnt = 0;
            foreach (string date in datesTW)
            {
                //https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php?l=zh-tw&o=JSON&d=111/03/01&s=0,asc,0
                string tpexUrl = "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php";
                string download_url = tpexUrl + $"?l=zh-tw&o=JSON&d={date}&s=0,asc,0";
                string downloadedData = "";
                // 網頁回傳
                using (WebClient wClient = new WebClient())
                {
                    wClient.Encoding = Encoding.UTF8;
                    downloadedData = wClient.DownloadString(download_url);
                }
                JObject JSONobject = JObject.Parse(downloadedData);
                JToken data = JSONobject.SelectToken("iTotalRecords");
                if (((Newtonsoft.Json.Linq.JValue)data).Value.ToString() != "0")
                {
                    JToken stocks = JSONobject.SelectToken("aaData");
                    for (int i = 0; i < stocks.Count(); i++)
                    {
                        if (stocks[i][0].ToString().Length == 4)
                        {                                                                                                                 // 代號               名稱                 成交股數            成交筆數        成交金額            開盤價           最高價               最低價             收盤價             日期西元
                            cmd.CommandText = $"INSERT INTO stocks VALUES('{stocks[i][0]}', '{stocks[i][1]}','{stocks[i][8]}','{stocks[i][10]}','{stocks[i][9]}','{stocks[i][4]}','{stocks[i][5]}','{stocks[i][6]}','{stocks[i][2]}', '{datesUS[cnt]}')";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                // 等10秒 在進行下一次呼叫證交所API
                Thread.Sleep(10000);
                cnt++;
            }
            cn.Close();
        }
        public class stockInfo
        {
            public string stockID { get; set; }
            public string stockName { get; set; }
            public float stockPrice { get; set; }

        }

        // 寫入資料到SQL (今天)
        private void button2_Click(object sender, EventArgs e)
        {
            SqlConnection cn = new SqlConnection(@"Data Source=.\sqlexpress;Initial Catalog=Lab;Integrated Security=True");
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cn.Open();

            // 要改日期
            string today = "20220308";
            string todayTW = "111/03/08";
            string dbTable = "stockPrice";


            //https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date={date.ToString(%22yyyyMMdd%22)}&type=ALLBUT0999&_=1586529875476
            string twseUrl = "https://www.twse.com.tw/exchangeReport/MI_INDEX";
            string download_urlTwse = twseUrl + $"?response=json&date={today}&type=ALL";
            string downloadedDataTwse = "";
            string tpexUrl = "https://www.tpex.org.tw/web/stock/aftertrading/daily_close_quotes/stk_quote_result.php";
            string download_urltpex = tpexUrl + $"?l=zh-tw&o=JSON&d={todayTW}&s=0,asc,0";
            string downloadedDatatpex = "";

            using (WebClient wClient = new WebClient())
            {
                // 網頁回傳
                wClient.Encoding = Encoding.UTF8;
                downloadedDataTwse = wClient.DownloadString(download_urlTwse);
                downloadedDatatpex = wClient.DownloadString(download_urltpex);
                //downloadedData.Replace("--", "null");
            }
            #region 寫入上市
            if (downloadedDataTwse.Contains("field"))
            {
                JObject JSONobjectTwse = JObject.Parse(downloadedDataTwse);
                JToken stocks = JSONobjectTwse.SelectToken("data9");
                // 股票代碼幾碼 > stocks[0][0].ToString().Length;
                for (int i = 0; i < stocks.Count(); i++)
                {
                    if (stocks[i][0].ToString().Length == 4)
                    { 
                        string numOfSharesTrade = replaceDot(stocks[i][2]);
                        string numOfTrade = replaceDot(stocks[i][3]);
                        string moneyOfDeal = replaceDot(stocks[i][4]);
                        string openPrice = replaceDot(stocks[i][5]);
                        string highPrice = replaceDot(stocks[i][6]);
                        string lowPrice = replaceDot(stocks[i][7]);
                        string endPrice = replaceDot(stocks[i][8]);
                                                                                                                                        // 代號               名稱                 成交股數            成交筆數        成交金額            開盤價           最高價               最低價             收盤價             日期西元
                        cmd.CommandText = $"INSERT INTO {dbTable} VALUES('{stocks[i][0]}', '{stocks[i][1]}','{numOfSharesTrade}','{numOfTrade}','{moneyOfDeal}','{openPrice}','{highPrice}','{lowPrice}','{endPrice}', '{today}')";
                        // int p = 0;
                        cmd.ExecuteNonQuery();
                    }
                }
            }
            #endregion

            #region 寫入上櫃
            // 寫入上櫃
            JObject JSONobjecttpex = JObject.Parse(downloadedDatatpex);
            JToken datatpex = JSONobjecttpex.SelectToken("iTotalRecords");
            if (((Newtonsoft.Json.Linq.JValue)datatpex).Value.ToString() != "0")
            {
                JToken stocks = JSONobjecttpex.SelectToken("aaData");
                for (int i = 0; i < stocks.Count(); i++)
                {
                    if (stocks[i][0].ToString().Length == 4)
                    {
                        string numOfSharesTrade = replaceDot(stocks[i][8]);
                        string numOfTrade = replaceDot(stocks[i][10]);
                        string moneyOfDeal = replaceDot(stocks[i][9]);
                        string openPrice = replaceDot(stocks[i][4]);
                        string highPrice = replaceDot(stocks[i][5]);
                        string lowPrice = replaceDot(stocks[i][6]);
                        string endPrice = replaceDot(stocks[i][2]);
                        //3,4,5,6,7,8,9                                                                                     // 代號               名稱                 成交股數                         成交筆數                成交金額                開盤價           最高價          最低價         收盤價       日期西元
                        cmd.CommandText = $"INSERT INTO {dbTable} VALUES('{stocks[i][0]}', '{stocks[i][1]}','{numOfSharesTrade}','{numOfTrade}','{moneyOfDeal}','{openPrice}','{highPrice}','{lowPrice}','{endPrice}', '{today}')";
                        cmd.ExecuteNonQuery();
                    }
                }
                #endregion

                cn.Close();

            }

            button2.Text = "ok";
        }

        // 寫入上市資料到SQL (1年)
        private void button3_Click(object sender, EventArgs e)
        {
            #region 日期
            List<string> dates = new List<string>();
            string[] years = { "2021", "2022" };
            string[] months = { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" };
            List<string> days = new List<string>();
            for (int i = 1; i < 32; i++)
            {
                if (i < 10)
                {
                    days.Add("0" + i.ToString());
                }
                else
                {
                    days.Add(i.ToString());
                }
            }
            //days.ToArray();
            int q = 0;
            foreach (var year in years)
            {
                if (q == 1)
                {
                    break;
                }
                foreach (var month in months)
                {
                    if (q == 1)
                    {
                        break;
                    }
                    foreach (var day in days)
                    {
                        string a = year + month + day;
                        if (a == DateTime.Now.ToString("yyyyMMdd"))
                        {
                            q = 1;
                            break;
                        }
                        else
                        {
                            dates.Add(a);
                        }
                    }
                }
            }
            dates.RemoveRange(0, 200);
            // int z = 0;
            #endregion
            SqlConnection cn = new SqlConnection(@"Data Source=.\sqlexpress;Initial Catalog=Lab;Integrated Security=True");
            SqlCommand cmd = new SqlCommand();
            cmd.Connection = cn;
            cn.Open();
            foreach (var date in dates)
            {
                //https://www.twse.com.tw/exchangeReport/MI_INDEX?response=json&date={date.ToString(%22yyyyMMdd%22)}&type=ALLBUT0999&_=1586529875476
                string twseUrl = "https://www.twse.com.tw/exchangeReport/MI_INDEX";
                string download_url = twseUrl + $"?response=json&date={date}&type=ALL";
                string downloadedData = "";
                using (WebClient wClient = new WebClient())
                {
                    // 網頁回傳
                    wClient.Encoding = Encoding.UTF8;
                    downloadedData = wClient.DownloadString(download_url);
                    //downloadedData.Replace("--", "null");
                }
                if (downloadedData.Contains("field"))
                {
                    JObject JSONobject = JObject.Parse(downloadedData);
                    JToken stocks = JSONobject.SelectToken("data9");
                    // 股票代碼幾碼 > stocks[0][0].ToString().Length;
                    for (int i = 0; i < stocks.Count(); i++)
                    {
                        if (stocks[i][0].ToString().Length == 4)
                        { // 0, 1, 2, 3, 4, 5, 6, 7, 8, 9, 10, 15, fdate
                            var upDowm = ((Newtonsoft.Json.Linq.JValue)stocks[i][9]).Value.ToString().Contains("-") ? "-" : "+";
                            cmd.CommandText = $"INSERT INTO stocks VALUES('{stocks[i][0]}', '{stocks[i][1]}','{stocks[i][2]}','{stocks[i][3]}','{stocks[i][4]}','{stocks[i][5]}','{stocks[i][6]}','{stocks[i][7]}','{stocks[i][8]}','{upDowm}','{stocks[i][10]}', '{stocks[i][15]}', '{date}')";
                            cmd.ExecuteNonQuery();
                        }
                    }
                }
                // 等7秒 在進行下一次呼叫證交所API
                Thread.Sleep(10000);
            }
            cn.Close();
        }

        // 處理日期
        private void button4_Click(object sender, EventArgs e)
        {
            #region 日期

            string[] yearsTW = { "110", "111" };
            string[] yearsUS = { "2021", "2022" };
            string[] months = { "01", "02", "03", "04", "05", "06", "07", "08", "09", "10", "11", "12" };
            List<string> days = new List<string>();
            for (int i = 1; i < 32; i++)
            {
                if (i < 10)
                {
                    days.Add("0" + i.ToString());
                }
                else
                {
                    days.Add(i.ToString());
                }
            }
            //days.ToArray();
            int q = 0;
            foreach (var year in yearsTW)
            {
                foreach (var month in months)
                {

                    foreach (var day in days)
                    {
                        //string a = year + month + day;
                        if ($"{int.Parse(year) + 1911}/{month}/{day}" == DateTime.Now.ToString("yyyy/MM/dd"))
                        {
                            //datesTW.RemoveRange(0, 200);
                            //datesUS.RemoveRange(0, 200);
                            return;
                        }
                        else
                        {
                            datesTW.Add($"{year}/{month}/{day}");
                            datesUS.Add($"{yearsUS[q]}{month}{day}");
                        }
                    }
                }
                q++;
            }

            // int t = 0;
            #endregion
        }

        private void testbtn_Click(object sender, EventArgs e)
        {
            string a = "1,222,111.25";
            string b = a.Replace(",", "");
            testbtn.Text = b;

            //decimal c = decimal.Parse(b);
            //testbtn.Text = a.Count(',');

        }


        // 刪除資料逗號
        public string replaceDot(JToken token)
        {
            string strTk = token.ToString();
            while (strTk.Contains(","))
            {
                strTk = strTk.Remove(strTk.IndexOf(","), 1);
            }
            return strTk;
        }


    }
}


