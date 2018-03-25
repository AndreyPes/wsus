using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Data.SqlClient;
using System.IO;
using System.Data;
namespace WSUS
{
    public  class pccount
    {
       public pccount(int _count,string _pcname)
        {
            count = _count;
            pcname = _pcname;
        }
        public int count { get; set; }
        public string pcname { get; set; }

    }

 
    class Program
    {
       
        private static int lenght;
        private static string pcname;
        private static List<pccount> ls = new List<pccount>();
        private static List<pccount> lsNull = new List<pccount>();
        private static string constr1 = "SELECT left(tbComputerTarget.FullDomainName, 30) as [Machine Name]" +
       " , count(tbComputerTarget.FullDomainName) as [# of Missing patches]" +
        ", tbComputerTarget.LastSyncTime as [Last Sync Time]" +
        ", tbComputerTarget.LastReportedStatusTime as [L Time]" +
"FROM tbUpdateStatusPerComputer INNER JOIN tbComputerTarget ON tbUpdateStatusPerComputer.TargetID =" +
       "  tbComputerTarget.TargetID and tbComputerTarget.TargetID = any(SELECT TargetID" +
" FROM[SUSDB].[dbo].[tbTargetInTargetGroup]" +
" where  TargetGroupID = 'EB51EBB4-2A0F-4D96-A8B8-F7654B9090B0')" +
"WHERE";
        private static string constr2 = "   tbUpdateStatusPerComputer.LocalUpdateID IN (SELECT LocalUpdateID FROM dbo.tbUpdate WHERE UpdateID IN" +
           " (SELECT UpdateID FROM PUBLIC_VIEWS.vUpdateApproval WHERE Action = 'Install')" +
          "  )" +

           " GROUP BY tbComputerTarget.FullDomainName, tbComputerTarget.LastSyncTime, tbComputerTarget.LastReportedStatusTime";

        public static string FridayString()
        {
      
            DateTime dt = DateTime.Now;
            if (dt.DayOfWeek.ToString() == "Friday")
            {
                return "В период " + dt.AddDays(-7).ToShortDateString() + " - " + dt.ToShortDateString() + " число обновлений, ожидающих установки:";
            }
            else
            {
                while(dt.DayOfWeek.ToString() != "Friday")
                {
                    dt = dt.AddDays(+1);
                   
                }            
                return "В период " + dt.AddDays(-7).ToShortDateString() + " - " + dt.ToShortDateString() + " число обновлений, ожидающих установки:";
            }
        }


        static string QueryStringCreator(bool fl, string str1, string str2)
        {
            if (fl == false)
            {
                return str1 + "((tbUpdateStatusPerComputer.SummarizationState = 5)) AND" + str2;
            }
            else
            {
                return str1 + "(not(tbUpdateStatusPerComputer.SummarizationState in ('1', '4'))) AND" + str2;
            }
        }
     
        static void Main(string[] args)
        {
            SqlConnection scon = new SqlConnection(@"Data Source=\\.\pipe\MICROSOFT##WID\tsql\query;Database=SUSDB;Trusted_Connection=yes");
           

            SqlCommand scom = new SqlCommand(QueryStringCreator(true, constr1, constr2), scon);
            try
            {
                Console.WriteLine("начало операции");
                scon.Open();
                Console.WriteLine("Соединение установлено");
                scom.CommandTimeout = 15;
                Console.WriteLine("Счётчик по существующим данным");
                using (SqlDataReader reader = scom.ExecuteReader())
                {
                    int i = 0;
                    while (reader.Read() )
                    {
                     
                      reader.GetInt32(1);
                      lenght = reader.GetString(0).Length;
                      pcname =  reader.GetString(0).Substring(0, lenght - 11);
                        Console.WriteLine("Add: " + pcname);
                        ls.Add(new pccount(reader.GetInt32(1), pcname));
                    }
                }
                Console.WriteLine("Done!");
                scom.CommandText= QueryStringCreator(false, constr1, constr2); ;
                scom.CommandType = CommandType.Text;
                scom.CommandTimeout = 15;
                Console.WriteLine("Счётчик по нулевым данным");
                using (SqlDataReader reader = scom.ExecuteReader())
                {

                    while (reader.Read())
                    {
                     
                        reader.GetInt32(1);
                        lenght = reader.GetString(0).Length;
                        pcname = reader.GetString(0).Substring(0, lenght - 11);
                        Console.WriteLine("Add: " + pcname);
                        lsNull.Add(new pccount(reader.GetInt32(1), pcname));
                    }
                }
                scon.Close();
                Console.WriteLine("Done!");
            }



              

            catch (Exception ex)
            {
                Console.WriteLine("Catch" + "\r\n");
                scon.Close();
                Console.WriteLine(ex.ToString());
                Console.ReadKey();
            }



            string filepath = AppDomain.CurrentDomain.BaseDirectory+"UpdatePcCount"+".txt";
            using (StreamWriter file = new StreamWriter(@"" + filepath, false))
            {
                file.WriteLine(FridayString());

                foreach (pccount s in ls)
                {

                    if (lsNull.Any(p => p.pcname == s.pcname))
                    {
                        file.Write(s.pcname + " - " + 0 + ";");
                        Console.WriteLine(s.pcname + " - " + 0 + "- Null");
                    }
                    else
                    {
                        file.Write(s.pcname + " - " + s.count + ";");
                        Console.WriteLine(s.pcname + " - " + s.count + "- Has value");
                    }
                }

                file.Write("Gate.energo.by - 0;Ns1.energo.by - 0;Ns3.energo.by - 0;WWW.energo.by - 1.\r\n\r\n");

            }

            Console.WriteLine("Done");
            Console.ReadKey();


        }
    }
}
