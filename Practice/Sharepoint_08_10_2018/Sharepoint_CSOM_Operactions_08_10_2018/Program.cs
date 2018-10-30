using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using Microsoft.SharePoint.Client;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using System.IO;


namespace Sharepoint_CSOM_Operactions_08_10_2018
{
    class Program
    {
        static void Main(string[] args)
        {
            String UserName = "bharat.naidu@acuvate.com";
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext con = new ClientContext("https://acuvatehyd.sharepoint.com/teams/My_Site"))
            {
                con.Credentials = new SharePointOnlineCredentials(UserName,Password);
                Web WebApplication = con.Web;

                //Site site = con.Site;
                //con.Load(site);

                //Web web = con.Web;
                //con.Load(web);

                //Folder f = web.RootFolder;
                //con.Load(f);

                //con.ExecuteQuery();
                //Console.WriteLine(web.Title);

                //Console.WriteLine(f.Name);


                List targetList = con.Web.Lists.GetByTitle("Documents");
                con.ExecuteQuery();
                FileCreationInformation fci = new FileCreationInformation();
                fci.Content = System.IO.File.ReadAllBytes(@"..\..\SampleFile2.txt");
                fci.Url = "SampleFile2.txt";
                fci.Overwrite = true;
                SP. File fileToUpload = targetList.RootFolder.Files.Add(fci);
                con.Load(fileToUpload);
                con.ExecuteQuery();



                //List targetList = con.Web.Lists.GetByTitle("Documents");
                //con.ExecuteQuery();
                
                //FileCreationInformation fci = new FileCreationInformation();
                //fci.Content = System.IO.File.ReadAllBytes(@"..\..\SampleFile.txt");
                //fci.Url = Path.Combine("Documents/NewFolder", Path.GetFileName("SampleFile.txt"));
                //fci.Overwrite = true;

                //var fileToUpload = targetList.RootFolder.Files.Add(fci);
                //con.Load(fileToUpload);
                //con.ExecuteQuery();


                //SP.List list = WebApplication.Lists.GetByTitle("list1");
                //con.Load(list.Fields);

                //con.ExecuteQuery();
                //foreach (Field f in list.Fields)
                //{
                //    Console.WriteLine(f.InternalName);
                //}
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
