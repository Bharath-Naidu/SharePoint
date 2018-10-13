using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;
using System.Text;
using System.Threading.Tasks;
using SP = Microsoft.SharePoint.Client;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
namespace Practice
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
                con.Credentials = new SharePointOnlineCredentials(UserName, Password);
                Web web = con.Web;
                List RList = con.Web.Lists.GetByTitle("Documents");
                con.Load(RList);
                con.ExecuteQuery();
                Folder Rfolder = RList.RootFolder;
                FileCollection Allfile = Rfolder.Files;
                con.Load(Allfile);
                con.ExecuteQuery();
                foreach(File fs in Allfile)
                {
                    File Rfile = fs;
                    con.Load(Rfile);
                    con.ExecuteQuery();
                    if (Rfile.Name.Equals("FilesInformation.xlsx"))
                        InsertIntoDataTable(con, Rfile.Name);
                }
            }
            Console.ReadKey();
        }
        public static void InsertIntoDataTable(ClientContext con,String fileName)
        {
            string strErrorMsg = string.Empty;
            string lstDocName = "Documents";
            try
            {
                DataTable dataTable = new DataTable("FileInformation");
                List list = con.Web.Lists.GetByTitle(lstDocName);
                con.Load(list.RootFolder);
                con.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                Microsoft.SharePoint.Client.File file = con.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                con.Load(file);
                con.ExecuteQuery();
                using (System.IO.MemoryStream mStream = new System.IO.MemoryStream())
                {
                    if (data != null)
                    {
                        data.Value.CopyTo(mStream);
                        using (SpreadsheetDocument document = SpreadsheetDocument.Open(mStream, false))
                        {
                            WorkbookPart workbookPart = document.WorkbookPart;
                            IEnumerable<Sheet> sheets = document.WorkbookPart.Workbook.GetFirstChild<Sheets>().Elements<Sheet>();
                            string relationshipId = sheets.First().Id.Value;
                            WorksheetPart worksheetPart = (WorksheetPart)document.WorkbookPart.GetPartById(relationshipId);
                            Worksheet workSheet = worksheetPart.Worksheet;
                            SheetData sheetData = workSheet.GetFirstChild<SheetData>();
                            IEnumerable<Row> rows = sheetData.Descendants<Row>();
                            foreach (Cell cell in rows.ElementAt(0))
                            {
                                string str = GetCellValue(con, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(con, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                            dataTable.Rows.RemoveAt(0);
                        }
                    }
                }
                Display(con,dataTable, fileName);

            }
            catch (Exception e)
            {

                strErrorMsg = e.Message;
            }
        }
       public static void Display(ClientContext con,DataTable data,String fileName)
        {
            List list = con.Web.Lists.GetByTitle("Documents");
            con.Load(list);
            con.ExecuteQuery();
            Folder newFolder = list.RootFolder;
            newFolder.Folders.Add("FileUpload");
            con.Load(newFolder);
            con.ExecuteQuery();
            //String URL = newFolder.ServerRelativeUrl;
            Console.WriteLine(fileName);
            foreach (DataRow row in data.Rows)
            {
                try
                {
                    String FilePath = row[0].ToString();
                    String FileName = row[1].ToString();
                    Console.WriteLine(FilePath);
                    Console.WriteLine(FileName);
                    String url = newFolder.ServerRelativeUrl;
                    Folder folder = con.Web.GetFolderByServerRelativeUrl(url);
                    FileCreationInformation fci = new FileCreationInformation();
                    fci.Content = System.IO.File.ReadAllBytes(FilePath);
                    fci.Url = FileName;
                    fci.Overwrite = true;
                    File fileToUpload = folder.Files.Add(fci);
                    con.Load(fileToUpload);
                    con.ExecuteQuery();
                    Console.WriteLine("Done");
                    break;
                }
                catch(Exception ex)
                {
                    Console.WriteLine(ex);
                    break;
                }
                
            }
        }
        private static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
        {
           
            string strErrorMsg = string.Empty;
            string value = string.Empty;
            try
            {
                if (cell != null)
                {
                    SharedStringTablePart stringTablePart = document.WorkbookPart.SharedStringTablePart;
                    if (cell.CellValue != null)
                    {
                        value = cell.CellValue.InnerXml;
                        if (cell.DataType != null && cell.DataType.Value == CellValues.SharedString)
                        {
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                            {
                               
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                            }
                        }
                        else
                        {
                            
                            return value;
                        }
                    }
                }
               
                return string.Empty;
            }
            catch (Exception e)
            {
                
                strErrorMsg = e.Message;
            }
           
            return value;
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