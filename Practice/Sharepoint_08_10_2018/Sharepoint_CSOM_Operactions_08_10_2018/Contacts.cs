using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_CSOM_Operactions_08_10_2018
{
    class Contacts
    {
        static void Main(string[] args)
        {
            String UserName = "bharat.naidu@acuvate.com";
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext con = new ClientContext("https://acuvatehyd.sharepoint.com"))
            {
                con.Credentials = new SharePointOnlineCredentials(UserName, Password);

                Site S = con.Site;
                con.Load(S);

                Web w = S.RootWeb;
                con.Load(w);

                List L = w.Lists.GetByTitle("Documents");
                con.Load(L);
                con.ExecuteQuery();
                //Console.WriteLine();
                Console.WriteLine(w.Title);



                Console.ReadKey();
            }
        }
        private static SecureString GetPassword()
        {
            ConsoleKeyInfo ck;
            SecureString SC = new SecureString();
            do
            {
                ck = Console.ReadKey(true);
                if (ck.Key != ConsoleKey.Enter)
                {
                    SC.AppendChar(ck.KeyChar);
                    Console.Write("*");
                }
            } while (ck.Key != ConsoleKey.Enter);
            return SC;
        }
    }
}
