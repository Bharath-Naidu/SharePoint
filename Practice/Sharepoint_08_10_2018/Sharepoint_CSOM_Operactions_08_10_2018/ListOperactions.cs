using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;

namespace Sharepoint_CSOM_Operactions_08_10_2018
{
    class ListOperactions
    {
        static void Main(string[] args)
        {
            String UserName = "bharat.naidu@acuvate.com";
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext con = new ClientContext("https://acuvatehyd.sharepoint.com/teams/My_Site"))
            {
                con.Credentials = new SharePointOnlineCredentials(UserName, Password);

                //ListCreationInformation List1 = new ListCreationInformation();
                //List1.Title = "NewList";
                //List1.Description = "The updated List";
                //List1.TemplateType = (int)ListTemplateType.Contacts;
                //List1.QuickLaunchOption = QuickLaunchOptions.On;
                //List L = con.Web.Lists.Add(List1);
                //con.Load(con.Web.Lists,l=>l.Include(ll=>ll.Title,ll=>ll.Id));

                //con.ExecuteQuery();
                //foreach(List ListItems in con.Web.Lists)
                //{
                //    Console.WriteLine(ListItems.Title);
                //}


                //--------------------------------Adding fields ----------------------------------------------------
                //List ListName=con.Web.Lists.GetByTitle("NewList");

                //Field field = ListName.Fields.AddFieldAsXml("<Field DisplayName='Alternative Phone number' Type='Number'/>",true,AddFieldOptions.DefaultValue);
                //FieldNumber number = con.CastTo<FieldNumber>(field);
                //number.MaximumValue = 10;
                //number.MinimumValue = 10;
                //number.Update();
                //con.ExecuteQuery();
                //---------------------------------Insert data on Fields----------------------------------------------------

                List list = con.Web.Lists.GetByTitle("All Employees");
                ListItem Li = list.AddItem(new ListItemCreationInformation());
                Li["Title"] = "A";
                Li["Department"] = "Placement";
                Li.Update();
                con.ExecuteQuery();
                //---------------------------------Update the fields---------------------------------------------
                //List contactsList = con.Web.Lists.GetByTitle("NewList");
                //ListItem itemToUpdate = contactsList.GetItemById(2);
                //itemToUpdate["Company"] = "ACUVATE";
                //itemToUpdate.Update();
                //---------------------------------Retriveing the data from list-------------------------------------------

                //List Listname = con.Web.Lists.GetByTitle("list1");
                //con.Load(Listname);
                //con.ExecuteQuery();
                //String Address = "vij";
                //CamlQuery query = new CamlQuery();
                //query.ViewXml = "<View><Query><Where><Eq><FieldRef Name='Address'/><Value Type='Text'>"+Address+"</Value></Eq></Where></Query></View>";

                //ListItemCollection listItems = Listname.GetItems(query);
                //con.Load(listItems);

                //con.ExecuteQuery();

                //ListItem li1;

                //foreach(ListItem li in listItems)
                //{
                //    Console.WriteLine("Title="+li["Title"]+"\naddress"+li["Address"]+"\nPhone"+li["Phone"]);
                //    li1 = Listname.GetItemById(li["ID"].ToString());
                //    li1.DeleteObject();
                //}
                //con.ExecuteQuery();
                //Console.ReadKey();
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

