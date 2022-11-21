using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace MailLimitOfWithdraw.Models
{
    public class LimitData
    {
        //已提領限額
        public string UsedAllCash { get; set; }
        //提領限額
        public string AllCashUsedLimit { get; set; }
        //特店編號
        public string MID { get; set; }
        //特店名稱
        public string MerchantName { get; set; }
        //寄件狀態
        public string IsMailed { get; set; }
        //特店信箱
        public string Email { get; set; }
        //執行時間
        //public int ExecTime { get; set; }
    }
}


