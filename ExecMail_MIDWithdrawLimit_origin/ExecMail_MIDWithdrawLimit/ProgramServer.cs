using MailLimitOfWithdraw.Models;
using System;
using System.Data.SqlClient;
using System.Data;
using System.Diagnostics;
using System.Globalization;
using System.Net.Mail;
using System.Configuration;

namespace MailLimitOfWithdraw.Servers
{
    internal class ProgramServer
    {
        public LimitData newLimitData = new LimitData();
        #region 修改資料寄出狀態
        private void DBStateUpdate(string isMailedstate, string UpdateMID)
        {
            SqlConnection db = new SqlConnection();
            string connectString = ConfigurationManager.AppSettings["DBconnectString_Coins"];
            string strMsgUpdate = @"update dbo.AllPay_MIDUsedLimitIsMailed set isMailed = @value1 where MID = @value2";
            SqlCommand upDate = new SqlCommand(strMsgUpdate, db);

            try
            {
                db.ConnectionString = connectString;
                db.Open();
                upDate.Parameters.AddWithValue("@value1", isMailedstate);
                upDate.Parameters.AddWithValue("@value2", UpdateMID);
                upDate.ExecuteNonQuery();
                db.Close();
            }
            catch (Exception e)
            {
                //抓錯誤訊息
                Console.WriteLine(e.Message, "警告");
            }
            finally
            {
                //清除
                upDate.Cancel();
                db.Close();
                db.Dispose();
            }
        }
        #endregion

        #region 新增Log資料
        private void DBCreateMailLog(string Email, string MID, string MerchantName, string AllCashUsedLimit, string UsedAllCash, int ExecTime)
        {
            string connectString = ConfigurationManager.AppSettings["DBconnectString_Share"];
            //db
            SqlConnection db = new SqlConnection(connectString);
            SqlCommand cmdLog = new SqlCommand("ausp_Share_SendMailContentLog_I", db)
            {
                CommandType = CommandType.StoredProcedure
            };
            cmdLog.Parameters.Add("@MailID", SqlDbType.BigInt).Value = 11111;
            cmdLog.Parameters.Add("@MailKey", SqlDbType.VarChar, 50).Value = "Mail_LogIn_Notification";
            cmdLog.Parameters.Add("@JsonReplaceStrTitle", SqlDbType.NVarChar, -1).Value = "{\"Title\":\"每月提領額度通知\"}";
            cmdLog.Parameters.Add("@JsonReplaceStrlayoutTitle", SqlDbType.NVarChar, -1).Value = "{\"Title\":\"每月提領額度通知\"}";
            cmdLog.Parameters.Add("@JsonReplaceStrBody", SqlDbType.NVarChar, -1).Value = "{\"<strong>每月提領額度通知</strong>\":\"\",\"MID\":\"}" + MID + "{\",\"MerchantName\":\"}" + MerchantName + "{\",\"AllCashUsedLimit\":\"}" + AllCashUsedLimit + "{\",\"UsedAllCash\":\"}" + UsedAllCash + "{\"}";
            cmdLog.Parameters.Add("@Sender", SqlDbType.VarChar, 200).Value = "test@mail.com";
            cmdLog.Parameters.Add("@Recever", SqlDbType.VarChar, 500).Value = Email;
            cmdLog.Parameters.Add("@SendMailType", SqlDbType.TinyInt).Value = 1;
            cmdLog.Parameters.Add("@NoticeStatus", SqlDbType.TinyInt).Value = 2;
            cmdLog.Parameters.Add("@DurationMs", SqlDbType.BigInt).Value = ExecTime;
            cmdLog.Parameters.Add("@SubMailID_Layout", SqlDbType.BigInt).Value = 5317;
            cmdLog.Parameters.Add("@Priority", SqlDbType.TinyInt).Value = 0;
            //執行SP
            db.Open();
            SqlDataReader reader = cmdLog.ExecuteReader();
            reader.Close();
            db.Close();
            db.Dispose();
        }
        #endregion

        #region 連接DB取得資料
        public void DBconnect()
        {
            string connectString = ConfigurationManager.AppSettings["DBconnectString_Coins"];

            //db
            SqlConnection db = new SqlConnection(connectString);
            //cmd
            SqlCommand cmd = new SqlCommand("ausp_AllPay_Coins_GetCustomerCoinsBalance_S", db)
            {
                CommandType = CommandType.StoredProcedure
            };
            cmd.Parameters.Add("@MID", SqlDbType.BigInt).Value = DBNull.Value;
            cmd.Parameters.Add("@Date", SqlDbType.VarChar, 15).Value = "2022-06-14";
            //取得資料
            db.Open();
            SqlDataReader reader = cmd.ExecuteReader();
            while (reader.Read())
            {
                ReadSingleRow(reader);
            }
            reader.Close();
            db.Close();
            db.Dispose();

            //取出單筆資料，寄信，Log func()
            void ReadSingleRow(IDataRecord dataRecord)
            {
                newLimitData.UsedAllCash = String.Format("{0}", dataRecord[4]);
                newLimitData.AllCashUsedLimit = String.Format("{0}", dataRecord[3]);
                newLimitData.MID = String.Format("{0}", dataRecord[21]);
                newLimitData.MerchantName = String.Format("{0}", dataRecord[22]);
                newLimitData.IsMailed = String.Format("{0}", dataRecord[24]);
                newLimitData.Email = String.Format("{0}", dataRecord[25]);
                MailState();
            }
        }
        #endregion

        #region 判斷並寄送
        private void MailState()
        {
            string UsedAllCash = newLimitData.UsedAllCash;
            string AllCashUsedLimit = newLimitData.AllCashUsedLimit;
            string MID = newLimitData.MID;
            string MerchantName = newLimitData.MerchantName;
            string isMailed = newLimitData.IsMailed;
            string Email = newLimitData.Email;
            //設定變數
            float UsedPercentage = (int.Parse(UsedAllCash, NumberStyles.AllowThousands) / int.Parse(AllCashUsedLimit, NumberStyles.AllowThousands)) * 100;
            //isMailed = "0";
            //UsedPercentage = 75;
            string Title = "";
            string partContent = "";
            if (UsedPercentage >= 70 && UsedPercentage < 90)
            {
                Title = "【每月提領額度通知】每月提領額度即將額滿！";
                partContent = "本月份的提領金額即將達到上限";
            }
            else if (UsedPercentage >= 90)
            {
                Title = "【每月提領額度通知】每月提領額度已達上限！";
                partContent = "本月份的提領金額已達到上限";
            }
            string bodycontent = "<body> <h1 style= width: 100%; text-align: center; padding: 30px 0 >每月提領額度通知</h1> <br /> <h2>親愛的會員，您好：</h2> <br /> <div style= padding-left:20px; ><div style= padding-left:20px; >  <p>提醒您，" + partContent + "，若需提高您的每月提領額度請聯絡客服或所屬業務。</p>  <br />  <p style= color:red >【請注意！若提領金額超過額度上限將無法提領】</p>  <br /><br />  <p style=  >會員編號: " + MID + "</p>  <p style=  >商店名稱: " + MerchantName + "</p>  <p style=  >本月提領上限：" + AllCashUsedLimit + " </p>  <p style=  >本月已提領金額：" + UsedAllCash + "</p>  <br /><br /><br /><p style=  >如有任何問題與FunPoint金流客服中心進行聯繫。謝謝！</p></div><p style= color:red >※ 本信件為系統自動發送 ( 請勿回信 )。</p> </div> <br /> <hr /> <br /> <a style= display: flex; justify-content: center  href= https://www.teat.com.tw/ >測試科技股份有限公司</a> <p style= display: flex; justify-content: center >服務時間：平日09:00-18:00 │ 若有任何問題，請洽詢您的業務專員或使用客服信箱。</p></body>";
            void mail()
            {
                string GoogleMailAcount = ConfigurationManager.AppSettings["GoogleMailAcount"];
                string GoogleMailPasswords = ConfigurationManager.AppSettings["GoogleMailPasswords"];
                System.Net.Mail.MailMessage msg = new System.Net.Mail.MailMessage();
                msg.To.Add("aldenwangf2e@gmail.com");
                msg.From = new MailAddress("XXX@gmail.com", "測試寄信", System.Text.Encoding.UTF8);
                msg.Subject = Title;
                msg.SubjectEncoding = System.Text.Encoding.UTF8;
                msg.Body = bodycontent;
                msg.BodyEncoding = System.Text.Encoding.UTF8;
                msg.IsBodyHtml = true;
                SmtpClient client = new SmtpClient
                {
                    Credentials = new System.Net.NetworkCredential(GoogleMailAcount, GoogleMailPasswords), //這裡要填正確的帳號跟密碼
                    Host = "smtp.gmail.com", //設定smtp Server
                    Port = 587, //設定Port
                    EnableSsl = true //gmail預設開啟驗證
                };
                client.Send(msg);
                client.Dispose();
                msg.Dispose();
                Console.WriteLine("郵件寄送成功!");
            }






            int ExecTime;
            //判斷並修改狀態
            if (isMailed == "0" && UsedPercentage >= 70 && UsedPercentage < 90)
            {
                var stopWatch = Stopwatch.StartNew(); //開始計算執行時間
                mail();
                DBStateUpdate("1", MID);
                stopWatch.Stop(); //結束計算執行時間
                ExecTime = Convert.ToInt32(stopWatch.ElapsedMilliseconds);
                DBCreateMailLog(Email, MID, MerchantName, AllCashUsedLimit, UsedAllCash, ExecTime);
                Console.WriteLine("成功寄出達70%信");

            }
            else if (isMailed == "1" && UsedPercentage >= 90)
            {
                var stopWatch = Stopwatch.StartNew();
                mail();
                DBStateUpdate("2", MID);
                stopWatch.Stop();
                ExecTime = Convert.ToInt32(stopWatch.ElapsedMilliseconds);
                Console.WriteLine(stopWatch.ElapsedMilliseconds);
                DBCreateMailLog(Email, MID, MerchantName, AllCashUsedLimit, UsedAllCash, ExecTime);
                Console.WriteLine("成功寄出達90%信");
            }
        }
        #endregion
    }
}
