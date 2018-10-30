using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;

namespace Sharepoint_CSOM_Operactions_08_10_2018
{
    class FolderOperactions
    {
        static void Main(string[] args)
        {
            String UserName = "bharat.naidu@acuvate.com";
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext con = new ClientContext("https://acuvatehyd.sharepoint.com/teams/My_Site"))
            {
                con.Credentials = new SharePointOnlineCredentials(UserName, Password);

                //---------------------------create the folder and upload the file on folder-----------------------
                List list = con.Web.Lists.GetByTitle("Documents");
                con.Load(list);
                con.ExecuteQuery();
                Folder newFolder = list.RootFolder;
                newFolder.Folders.Add("FileUpload");
                FolderCollection Allfolders = newFolder.Folders;
                con.Load(Allfolders);
                con.ExecuteQuery();
                foreach(Folder f in Allfolders)
                {
                    String url = f.ServerRelativeUrl;
                    Folder folder = con.Web.GetFolderByServerRelativeUrl(url);
                    con.Load(folder);
                    con.ExecuteQuery();
                    Console.WriteLine(f.Name);
                    if (f.Name.Equals("FileUpload"))
                    {
                        FileCreationInformation fci = new FileCreationInformation();
                        fci.Content = System.IO.File.ReadAllBytes(@"..\..\NewFile.txt");
                        fci.Url = "NewFile.txt";
                        fci.Overwrite = true;
                        File fileToUpload = folder.Files.Add(fci); 
                        con.Load(fileToUpload);
                        con.ExecuteQuery();
                    }
                }

                

                //-------------------get all folders-------------------------------------

                //List list = con.Web.Lists.GetByTitle("Documents");
                //con.Load(list);
                //con.ExecuteQuery();
                //FolderCollection Allfolders = list.RootFolder.Folders;
                //con.Load(Allfolders);
                //foreach(Folder f in Allfolders)
                //{
                //    Console.WriteLine(f.Name);
                //}
                //con.ExecuteQuery();
                //------------------delete the folder-----------------------------
                //List list = con.Web.Lists.GetByTitle("Documents");
                //CamlQuery caml = CamlQuery.CreateAllFoldersQuery();
                //Folder rootFolder = list.RootFolder;
                //FolderCollection folders = rootFolder.Folders;
                
                //con.Load(folders);
                //con.ExecuteQuery();
                //foreach (Folder folder in folders)
                //{
                //    if (folder.Name.Equals("NEW"))
                //    {
                //        String s = folder.ServerRelativeUrl;
                //        Console.WriteLine(s);
                //        Folder f = con.Web.GetFolderByServerRelativeUrl(s);
                //        con.Load(f);
                //        con.ExecuteQuery();
                //        Console.WriteLine(list.RootFolder.Folders.GetByUrl(s)+"\n"+ con.Web.GetFolderByServerRelativeUrl(s));
                //        Console.WriteLine(f.Name);
                //        f.DeleteObject();
                //        con.ExecuteQuery();
                //    }
                    
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
