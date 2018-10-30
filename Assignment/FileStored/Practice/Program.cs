using Microsoft.SharePoint.Client;
using System;
using System.Collections.Generic;
using System.Data;
using System.Linq;
using System.Security;
using DocumentFormat.OpenXml.Spreadsheet;
using DocumentFormat.OpenXml.Packaging;
using EX=System.IO;

namespace Practice
{
    class Program
    {
        public static DataTable OriginaldataTable=new DataTable("ExcelData"); //for storing the excel data
        static void Main(string[] args)
        {
            Console.WriteLine(Constant.EnterUserName);
            String UserName = Console.ReadLine();
            Console.WriteLine(Constant.EnterUserPassword);
            SecureString Password = GetPassword();
            using (ClientContext clientContext = new ClientContext(Constant.SiteURL)) 
            {
                clientContext.Credentials = new SharePointOnlineCredentials(UserName, Password);
                Web web = clientContext.Web;
                List RootList = clientContext.Web.Lists.GetByTitle(Constant.RootFolder);
                clientContext.Load(RootList);
                clientContext.ExecuteQuery();
                Folder Rootfolder = RootList.RootFolder; //Set the folder 
                FileCollection Allfile = Rootfolder.Files; //getting all file from the folder after compare the each one with required file
                clientContext.Load(Allfile);
                clientContext.ExecuteQuery();
                foreach(File PresentFile in Allfile)
                {
                    File file = PresentFile;
                    clientContext.Load(file);
                    clientContext.ExecuteQuery();
                    if (file.Name.Equals(Constant.FileInSharepoint)) //compare the each file on root folder
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
                DataTable dataTable = new DataTable("ExcelInformation");//Use temporary datatble 
                List list = clientContext.Web.Lists.GetByTitle(Constant.RootFolder);
                clientContext.Load(list.RootFolder);
                clientContext.ExecuteQuery();
                string fileServerRelativeUrl = list.RootFolder.ServerRelativeUrl + "/" + fileName; 
                File file = clientContext.Web.GetFileByServerRelativeUrl(fileServerRelativeUrl);//Extracting the file
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
                                string str = GetCellValue(document, cell); //getting cell data using the cell value(column name)
                                dataTable.Columns.Add(str); //after adding to the datatable
                            }
                            foreach (Row row in rows)                           //rows contains the all information now take each one from roes then added to the datatable
                            {
                                if (row != null) //if row does't contains the data then no need to insert the data
                                {
                                    DataRow dataRow = dataTable.NewRow();
                                    for (int i = 0; i < row.Descendants<Cell>().Count(); i++)
                                    {
                                        dataRow[i] = GetCellValue( document, row.Descendants<Cell>().ElementAt(i)); //here reading the entire row from excel sheet
                                    }
                                    dataTable.Rows.Add(dataRow); //finally addin the each row to the datatble
                                }
                            }
                             dataTable.Rows.RemoveAt(0);//deleting the column heading 
                        }
                    }
                }
                OriginaldataTable = dataTable.Copy();                       //copied to temporary datatble to original datatable   
               // display();
                UploadFile(clientContext, dataTable, fileName);             //Here Upload the all files to sharepoint
                string FileNameAfterChange = ExportToExcelSheet();
                UploadExcelFileToSharepoint(clientContext, FileNameAfterChange + ".xlsx"); //Upload the updated sharepoint 
            }
            catch (Exception e)
            {
                LogClass.RecordException(e);
                strErrorMsg = e.Message;
            }
        }
       //static void display()
       // {
       //     foreach(DataRow dr in OriginaldataTable.Rows)
       //     {
       //         foreach(var item in dr.ItemArray)
       //             Console.WriteLine(item);
       //     }
       // }
        static void UpdateDataTable(String FilePath,String Reason,String Status)
        {
            int Columns = OriginaldataTable.Columns.Count;
            foreach (DataRow DR in OriginaldataTable.Rows)
                if(DR[0].Equals(FilePath))  //here comparing the each file path to required one 
                {                           //initially the status & reason will be the null after uploading it will be modified
                    DR[Columns - 2] = Status;
                    DR[Columns - 1] = Reason;
                }
        }
        static void UploadFile(ClientContext clientContext, DataTable data,String fileName)
        {
            string Reason = "";
            string UploadStatus = "";
            string FilePath = "";
            List list = clientContext.Web.Lists.GetByTitle(Constant.RootFolder);
            clientContext.Load(list);
            clientContext.ExecuteQuery();
            Folder newFolder = list.RootFolder;
            clientContext.Load(newFolder);
            clientContext.ExecuteQuery();
            foreach (DataRow row in data.Rows)  //here feteching the all rows from datatble 
            {
                bool flag = true;
                long bytes = 0;
                try
                {
                    FilePath = row[0].ToString();
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilePath);
                    string FileType = fileInfo.Extension;
                    //try
                    //{
                        if (fileInfo.Exists)                         //read the file size
                            bytes = fileInfo.Length;
                        else
                            throw new EX.FileNotFoundException("File not found");
                    //}
                    //catch (Exception ex)
                    //{
                    //    UploadStatus = "Failed";
                    //    Reason = "File not exist on given path";
                    //    flag = false;
                    //}
                    if (flag == true && (bytes < 10000000)) //checking the file size before upload & file size is given range or not
                    {
                        int FileNameStarts = FilePath.LastIndexOf("\\");
                        string Filename = FilePath.Substring(FileNameStarts + 1);  //spliting the file name from file path
                        string CreatedBy = row[1].ToString();       //adding the column to the files 
                        string depart = row[2].ToString();
                        string Status = (row[3]).ToString();
                        FileCreationInformation File = new FileCreationInformation(); //Uploading the file to sharepoint library
                        File.Content = System.IO.File.ReadAllBytes(FilePath);
                        File.Url = Filename;
                        File.Overwrite = true;
                        File fileToUpload = newFolder.Files.Add(File);
                        clientContext.Load(fileToUpload);
                        clientContext.ExecuteQuery();
                        var newItem = fileToUpload.ListItemAllFields;   //creating the column items to the file 
                        newItem["CreatedByThisFile"] = CreatedBy.ToString();
                        newItem["Size"] = fileToUpload.Length;
                        newItem["Departement"] = depart;
                        Microsoft.SharePoint.Client.Field field = list.Fields.GetByInternalNameOrTitle("Status"); //here reading the status of the file
                        clientContext.Load(field);
                        clientContext.ExecuteQuery();
                        FieldChoice fieldChoices = clientContext.CastTo<FieldChoice>(field);
                        string[] StatusArray = Status.Split(',');
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
                       // Console.WriteLine(Filename+" is Done");
                        UploadStatus = "Success";
                        Reason = "";
                    }
                    else if (flag)
                    {
                        throw new Exception("File Size to high");
                        //Reason = "File size is to high";
                        //UploadStatus = "Failed";
                        
                    }
                    //after taking the all fields from file then update the excel sheet througth datatble
                }
                catch (EX.FileNotFoundException ex)
                {
                    UploadStatus = "Failed";
                    Reason = ex.Message;
                    LogClass.RecordException(ex);
                }
                catch (Exception ex)
                {
                    UploadStatus = "Failed";
                    Reason = ex.Message;
                    LogClass.RecordException(ex);
                }
                UpdateDataTable(FilePath, Reason, UploadStatus);
            }
            
        }
        static string GetCellValue(SpreadsheetDocument document, Cell cell)
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
                            if (stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)] != null)
                                return stringTablePart.SharedStringTable.ChildElements[Int32.Parse(value)].InnerText;
                        else
                            return value;
                    }
                }  
                return string.Empty;
            }
            catch (Exception e)
            {
                strErrorMsg = e.Message;
                LogClass.RecordException(e);
            }
                       return value;
        }
        static SecureString GetPassword() //getting user password from console
        {
            ConsoleKeyInfo ck;
            SecureString SC = new SecureString();
            do
            {
                ck = Console.ReadKey(true);
                if (ck.Key != ConsoleKey.Enter)
                {
                    SC.AppendChar(ck.KeyChar); //each character is reading
                    Console.Write("*");
                }
            } while (ck.Key != ConsoleKey.Enter); //reading  upto press ENTER
            return SC;
        }
        static string ExportToExcelSheet() //now the all information in the datatble now it is converted to excel
        {
            DataTable Table = OriginaldataTable.Copy();
            string ExcelFilePath = Constant.FileOnLocalSystem;
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application(); //creating the excel sheet
                excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel.Worksheet workSheet = excelApp.ActiveSheet;
                for(int count=0;count<Table.Columns.Count;count++)//retriving the column name from datatable and stored on excel
                    workSheet.Cells[1,(count+1)]=Table.Columns[count].ColumnName; 
                for(int Count=0;Count<Table.Rows.Count;Count++) //retriveing the each cell from datatable then it will be added to the excel sheet
                    for (int InnerLoop=0;InnerLoop<Table.Columns.Count;InnerLoop++) 
                        workSheet.Cells[(Count+2),(InnerLoop+1)]=Table.Rows[Count][InnerLoop];
                System.IO.FileInfo fileInfo = new System.IO.FileInfo(ExcelFilePath+".xlsx");
                if (fileInfo.Exists)
                    fileInfo.Delete();                                  //delete the existing one
                workSheet.SaveAs(ExcelFilePath);                           //saved on given path
                excelApp.Quit(); 
            }
            catch (Exception ex)
            { 
               LogClass.RecordException(ex);
                throw new Exception();
            }
            return ExcelFilePath;
        }
        static void UploadExcelFileToSharepoint(ClientContext clientContext, String FileLocation) //this is used to upload the excel to sharepoint 
        {
            
            try
            {
                List list = clientContext.Web.Lists.GetByTitle(Constant.RootFolder);
                clientContext.Load(list);
                clientContext.ExecuteQuery();

                Folder Root = list.RootFolder;
                clientContext.Load(Root);
                clientContext.ExecuteQuery();

                int last = FileLocation.LastIndexOf("\\");
                String Filename = FileLocation.Substring(last + 1);

                FileCreationInformation NewFile = new FileCreationInformation();
                NewFile.Content = System.IO.File.ReadAllBytes(FileLocation);
                NewFile.Url = Filename;
                NewFile.Overwrite = true;
                File fileToUpload = Root.Files.Add(NewFile);

                clientContext.Load(fileToUpload);
                clientContext.ExecuteQuery();
            }
            catch(Exception ex)
            {
                LogClass.RecordException(ex);
                throw new Exception(ex.Message);
            }
            Console.WriteLine("\n\n\nUploading successfully completed.............");
        }
    }
}
