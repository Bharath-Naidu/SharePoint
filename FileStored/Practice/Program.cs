using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;

namespace Practice
{
    class Program
    {
        public static DataTable OriginaldataTable=new DataTable("ExcelData"); //for storing the excel data
        static void Main(string[] args)
        {
            
            Console.WriteLine("Please enter the User Name");
            String UserName = Console.ReadLine();
            Console.WriteLine("please enter the password");
            SecureString Password = GetPassword();
            using (ClientContext clientContext = new ClientContext("https://acuvatehyd.sharepoint.com/teams/My_Site"))
            {
                clientContext.Credentials = new SharePointOnlineCredentials(UserName, Password);
                Web web = clientContext.Web;
                List RList = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(RList);
                clientContext.ExecuteQuery();
                Folder Rootfolder = RList.RootFolder;
                FileCollection Allfile = Rootfolder.Files;
                clientContext.Load(Allfile);
                clientContext.ExecuteQuery();
                foreach(File PresentFile in Allfile)
                {
                    File file = PresentFile;
                    clientContext.Load(file);
                    clientContext.ExecuteQuery();
                    if (file.Name.Equals("FilesInformation.xlsx")) //compare the each file on root folder
                    {
                        InsertIntoDataTable(clientContext, file.Name);//if it found then read the data from excel
                        break;
                    }
                }
            }
            
            Console.ReadKey();
        }
        static void InsertIntoDataTable(ClientContext clientContext, String fileName)//this method is used to read the
        {                                                                           //data from excel then stored on datatble
            string strErrorMsg = string.Empty;
            
            try
            {
                DataTable dataTable = new DataTable("FileInformation");//Use temporary datatble 
                List list = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName;
                Microsoft.SharePoint.Client.File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);//Extracting the file
                ClientResult<System.IO.Stream> data = file.OpenBinaryStream();
                clientContext.Load(file);
                clientContext.ExecuteQuery();
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
                            foreach (Cell cell in rows.ElementAt(0))            //reading the file colums
                            {
                                string str = GetCellValue(clientContext, document, cell);
                                dataTable.Columns.Add(str);
                            }
                            foreach (Row row in rows)                           //reading the file rows
                            {
                                if (row != null)
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue(clientContext, document, row.Descendants<Cell>().ElementAt(i));
                                    }
                                    dataTable.Rows.Add(dataRow);
                                }
                            }
                             dataTable.Rows.RemoveAt(0);
                        }
                    }
                }
                OriginaldataTable = dataTable.Copy();                       //copied to temporary datatble to original datatable
                
                UploadFile(clientContext, dataTable, fileName);             //Here Upload the all files to sharepoint
                string FileName=ExportToExcelSheet();
                UploadExcelFileToSharepoint(clientContext, FileName+".xlsx"); //Upload the updated sharepoint 
            }
            catch (Exception e)
            {

                strErrorMsg = e.Message;
            }
        }
        //public void display()
        //{
        //    foreach (DataRow DR in OriginaldataTable.Rows)
        //    {
        //        foreach(var item in DR.ItemArray)
        //            Console.WriteLine(item);
        //        Console.WriteLine();

        //    }
        //    ExportToExcelSheet();
        //}
        static void UpdateDataTable(String FilePath,String Reason,String Status)
        {
            int Column = OriginaldataTable.Columns.Count;
            foreach (DataRow DR in OriginaldataTable.Rows)
            {
                if(DR[0].Equals(FilePath))
                {
                    DR[Column - 2] = Status;
                    DR[Column - 1] = Reason;
                }
            }
        }
        static void UploadFile(ClientContext clientContext, DataTable data,String fileName)
        {
            String Reason = "";
            String UploadStatus = "";
            String FilePath = "";
            List list = clientContext.Web.Lists.GetByTitle("Documents");
            clientContext.Load(list);
            clientContext.ExecuteQuery();
            Folder newFolder = list.RootFolder;
            clientContext.Load(newFolder);
            clientContext.ExecuteQuery();
            
            foreach (DataRow row in data.Rows)
            {
                bool flag = true;
                long bytes = 0;
               
                try
                {
                    FilePath = row[0].ToString();
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilePath);
                    string FileType = fileInfo.Extension;
                    if (fileInfo.Exists)    //read the file size
                        bytes = fileInfo.Length;
                    else                    // if the file not exit then status will be failed
                    {
                        UploadStatus = "Failed";
                        Reason = "File not exist on given path";
                        flag = false;
                    }
                    if (flag == true && (bytes < 10000000)) //checking the file size before upload
                    {

                        int last = FilePath.LastIndexOf("\\");
                        string Filename = FilePath.Substring(last + 1);
                        
                        string CreatedBy = row[1].ToString();
                        int depart = Convert.ToInt32(row[2]);
                        string Status = (row[3]).ToString();
                        //Console.WriteLine(FilePath);
                        //Console.WriteLine(CreatedBy);

                        FileCreationInformation File = new FileCreationInformation(); //Uploading the file here
                        File.Content = System.IO.File.ReadAllBytes(FilePath);
                        File.Url = Filename;
                        File.Overwrite = true;

                        File fileToUpload = newFolder.Files.Add(File);
                        clientContext.Load(fileToUpload);
                        clientContext.ExecuteQuery();

                        var newItem = fileToUpload.ListItemAllFields;
                        newItem["CreatedByThisFile"] = CreatedBy.ToString();
                        newItem["Size"] = fileToUpload.Length;
                        newItem["Dept"] = depart;
                        Microsoft.SharePoint.Client.Field field = list.Fields.GetByInternalNameOrTitle("Status"); //here reading the status of the file
                        clientContext.Load(field);
                        clientContext.ExecuteQuery();
                        FieldChoice fieldChoices = clientContext.CastTo<FieldChoice>(field);
                        string[] StatusArray = Status.Split(':');
                        string finalyStatus="";
                        for(int count=0;count<StatusArray.Length;count++) 
                        {
                            if (fieldChoices.Choices.Contains(StatusArray[count]))
                            if (count == StatusArray.Length - 1)
                                finalyStatus = StatusArray[count];
                            else
                                finalyStatus = StatusArray[count]+";";
                        }

                        newItem["Status"] = finalyStatus;
                        newItem["TypeOfFile"] = FileType.ToString();
                        newItem.Update();
                        clientContext.ExecuteQuery();
                        
                        Console.WriteLine(Filename+" is Done");
                        UploadStatus = "Success";
                        Reason = "None";
                    }
                    else if (flag)
                    {
                        Reason = "File size is to high";
                        UploadStatus = "Failed";
                    }
                    UpdateDataTable(FilePath,Reason,UploadStatus); //after taking the all fields from file then update the excel sheet througth datatble
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex);
                    break;
                }   
            }   
        }
        static string GetCellValue(ClientContext clientContext, SpreadsheetDocument document, Cell cell)
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
        static SecureString GetPassword()
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
        static string ExportToExcelSheet() //now the all information in the datatble now it is converted to excel
        {
            DataTable Table = OriginaldataTable.Copy();
            string ExcelFilePath = "D:\\File\\FilesInformation";
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet workSheet = excelApp.ActiveSheet;
                for (int i = 0; i < Table.Columns.Count; i++)//retriving the column from datatable and stored on excel
                {
                    workSheet.Cells[1, (i + 1)] = Table.Columns[i].ColumnName;
                }
                for (int i = 0; i < Table.Rows.Count; i++)
                {
                    for (int j = 0; j < Table.Columns.Count; j++)
                    {
                        workSheet.Cells[(i + 2), (j + 1)] = Table.Rows[i][j];
                    }
                }
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(ExcelFilePath+".xlsx");
                if (fileInfo.Exists)
                    fileInfo.Delete(); //delete the existing one
                workSheet.SaveAs(ExcelFilePath); //saved on given path
                excelApp.Quit(); 
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
            return ExcelFilePath;
        }
        static void UploadExcelFileToSharepoint(ClientContext clientContext, String FileLocation)
        {
            try
            {

                List list = clientContext.Web.Lists.GetByTitle("Documents");
                clientContext.Load(list);
                clientContext.ExecuteQuery();
                Folder Root = list.RootFolder;
                clientContext.Load(Root);
                clientContext.ExecuteQuery();
                int last = FileLocation.LastIndexOf("\\");
                String Filename = FileLocation.Substring(last + 1);
                FileCreationInformation fci = new FileCreationInformation();
                fci.Content = System.IO.File.ReadAllBytes(FileLocation);
                fci.Url = Filename;
                fci.Overwrite = true;
                File fileToUpload = Root.Files.Add(fci);
                clientContext.Load(fileToUpload);
                clientContext.ExecuteQuery();
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
            Console.WriteLine("Uploading successfully completed");
        }
    }
}
