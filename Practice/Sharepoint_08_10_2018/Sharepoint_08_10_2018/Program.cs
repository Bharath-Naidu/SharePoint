using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_08_10_2018
{
    class Program
    {
        static void Main(string[] args)
        {
            String UserName = "bharat.naidu@acuvate.com";
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext con=new ClientContext("https://acuvatehyd.sharepoint.com/teams/My_Site"))
            {
                con.Credentials = new SharePointOnlineCredentials(UserName,Password);

                Web WebApplication = con.Web;

                //ListCreationInformation newList = new ListCreationInformation();
                //newList.Title = "ListCreatedUsingC#TestingHere";

                //newList.TemplateType = (int)ListTemplateType.Announcements;
                //List L =WebApplication.Lists.Add(newList);
                //L.Description = "This list is created using the C# codeing";
                //L.Update();
                //con.ExecuteQuery();

                //WebApplication.Description = "This is description about my site";
                //WebApplication.Update();

                //con.Load(WebApplication, p => p.Title, p => p.Description);
                //con.ExecuteQuery();

                con.Load(WebApplication.Lists, p1 => p1.Include(p2 => p2.Title, p3 => p3.Id));
                con.ExecuteQuery();

                foreach (List l in WebApplication.Lists)
                {
                    Console.WriteLine("title is:" + l.Title + "\nid is:" + l.Id);
                }
                Console.WriteLine("Please enter the Title to delate");
                String Tname = "";
                ConsoleKeyInfo CK;
                do
                {
                    CK = Console.ReadKey(true);
                    if (CK.Key != ConsoleKey.Enter)
                    {
                        Tname = Tname + CK.KeyChar;
                        Console.Write(CK.KeyChar);
                    }
                } while (CK.Key != ConsoleKey.Enter);
                //Console.WriteLine(Tname);
                List L = WebApplication.Lists.GetByTitle(Tname);
                ListItem item = L.GetItemById(2);
                //Console.WriteLine(L.GetItemById(2).ti);
                item.DeleteObject();
                //Console.WriteLine("Title is:" + L.Title + "\nDec:" + L.Description);

               
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
