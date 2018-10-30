using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_CSOM_Operactions_08_10_2018
{
    class UsersGroups
    {
        static void Main(string[] args)
        {
            String UserName = "bharat.naidu@acuvate.com";
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext con = new ClientContext("https://acuvatehyd.sharepoint.com/teams/My_Site"))
            {
                con.Credentials = new SharePointOnlineCredentials(UserName, Password);
                GroupCollection siteGroups = con.Web.SiteGroups;
                con.Load(siteGroups);
                con.ExecuteQuery();
                foreach(Group g in siteGroups)
                {
                    Console.WriteLine(g.Title+"\n"+g.Id);
                    Console.WriteLine();
                }
                Group membersGroup = siteGroups.GetByName("only ADD Permissions");
                con.Load(membersGroup.Users);
                con.ExecuteQuery();
                foreach (User member in membersGroup.Users)
                {
                    Console.WriteLine(member.Title+"\n");
                }
                UserCreationInformation NewUser = new UserCreationInformation();
                NewUser.Email = "arvind.torvi@acuvate.com";
                NewUser.LoginName = "nrfsecw.com";
                NewUser.Title = "Mr.Ussrerc";

                 membersGroup = siteGroups.GetByName("only ADD Permissions");
                bool IsUserNotExist = false ;
                try
                {
                    User Checkuser = membersGroup.Users.GetByEmail("arvind.torvi@acuvate.com");
                    con.Load(Checkuser);
                    con.ExecuteQuery();
                }
                catch (Exception e)
                {
                    IsUserNotExist = true;
                }
                if (IsUserNotExist)
                {
                    User member = con.Web.EnsureUser("arvind.torvi@acuvate.com");
                    member.Title = "thisthitle";
                    member.Update();
                    User U = membersGroup.Users.AddUser(member);
                    con.ExecuteQuery();
                }


                //User U=membersGroup.Users.Add(NewUser);
                ////con.Load(U);
                //con.ExecuteQuery();
                con.Load(membersGroup.Users);
                con.ExecuteQuery();
                foreach (User member in membersGroup.Users)
                {
                    Console.WriteLine(member.Title+"\t"+member.Email+"\t"+member.LoginName + "\n");
                }

            }
            Console.ReadKey();
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
