using System;
using System.Data;
using System.Diagnostics;
using System.Net.Mail;
using System.Data.SqlClient;
using System.Globalization;
using MailLimitOfWithdraw.Models;
using MailLimitOfWithdraw.Servers;


namespace ExecMail_MIDWithdrawLimit
{
    internal class Program
    {
        static void Main(string[] args)
        {
            ProgramServer newServer = new ProgramServer();
            newServer.DBconnect();
        }
    }

}
