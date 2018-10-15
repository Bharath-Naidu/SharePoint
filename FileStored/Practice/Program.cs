﻿using Microsoft.SharePoint.Client;
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
        public static DataTable OriginaldataTable=new DataTable("ExcelData");
        static void Main(string[] args)
        {
            
            Console.WriteLine("Please enter the User Name");
            String UserName = Console.ReadLine();
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
                    {
                        InsertIntoDataTable(con, Rfile.Name);
                        break;
                    }
                }
            }
            Console.WriteLine("\n\nUploading successfully completed");
            Console.ReadKey();
        }
        static void InsertIntoDataTable(ClientContext con,String fileName)
        {
            string strErrorMsg = string.Empty;
            string RootFile = "Documents";
            try
            {
                DataTable dataTable = new DataTable("FileInformation");
                List list = con.Web.Lists.GetByTitle(RootFile);
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
                OriginaldataTable = dataTable.Copy();
                
                UploadFile(con,dataTable, fileName);
                string FN=ExportToExcelSheet();
                UploadFileToSharepoint(con,FN+".xlsx");
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
            foreach (DataRow DR in OriginaldataTable.Rows)
            {
                if(DR[0].Equals(FilePath))
                {
                    DR[4] = Status;
                    DR[5] = Reason;
                }
            }
        }
        static void UploadFile(ClientContext con,DataTable data,String fileName)
        {
            String Reason = "";
            String UploadStatus = "";
            String FilePath = "";
            List list = con.Web.Lists.GetByTitle("Documents");
            con.Load(list);
            con.ExecuteQuery();
            Folder newFolder = list.RootFolder;
            con.Load(newFolder);
            con.ExecuteQuery();
            
            foreach (DataRow row in data.Rows)
            {
                bool flag = true;
                long bytes = 0;
               
                try
                {
                    FilePath = row[0].ToString();
                    System.IO.FileInfo fileInfo = new System.IO.FileInfo(FilePath);
                    string FileType = fileInfo.Extension;
                    if (fileInfo.Exists)
                        bytes = fileInfo.Length;
                    else
                    {
                        UploadStatus = "Failed";
                        Reason = "File not exist on given path";
                        flag = false;
                    }
                    if (flag == true && (bytes < 10000000))
                    {

                        int last = FilePath.LastIndexOf("\\");
                        string Filename = FilePath.Substring(last + 1);
                        
                        string CreatedBy = row[1].ToString();
                        int depart = Convert.ToInt32(row[2]);
                        string Status = (row[3]).ToString();
                        //Console.WriteLine(FilePath);
                        //Console.WriteLine(CreatedBy);

                        FileCreationInformation fci = new FileCreationInformation();
                        fci.Content = System.IO.File.ReadAllBytes(FilePath);
                        fci.Url = Filename;
                        fci.Overwrite = true;

                        File fileToUpload = newFolder.Files.Add(fci);
                        con.Load(fileToUpload);
                        con.ExecuteQuery();

                        var newItem = fileToUpload.ListItemAllFields;
                        newItem["CreatedByThisFile"] = CreatedBy.ToString();
                        newItem["Size"] = fileToUpload.Length;
                        newItem["Dept"] = depart;
                        newItem["Status"] = Status.ToString();
                        newItem["TypeOfFile"] = FileType.ToString();
                        newItem.Update();
                        con.ExecuteQuery();
                        
                        Console.WriteLine(Filename+" is Done");
                        UploadStatus = "Success";
                        Reason = "None";
                    }
                    else if (flag)
                    {
                        Reason = "File size is to high";
                        UploadStatus = "Failed";
                    }
                    UpdateDataTable(FilePath,Reason,UploadStatus);
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
        static string ExportToExcelSheet()
        {
            DataTable Table = OriginaldataTable.Copy();
            string ExcelFilePath = "D:\\File\\FilesInformation";
            try
            {
                Microsoft.Office.Interop.Excel.Application excelApp = new Microsoft.Office.Interop.Excel.Application();
                excelApp.Workbooks.Add();
                Microsoft.Office.Interop.Excel._Worksheet workSheet = excelApp.ActiveSheet;
                for (int i = 0; i < Table.Columns.Count; i++)
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
                    fileInfo.Delete();
                workSheet.SaveAs(ExcelFilePath);
                excelApp.Quit(); 
            }
            catch (Exception ex)
            {
                throw new Exception("ExportToExcel: \n" + ex.Message);
            }
            return ExcelFilePath;
        }
        static void UploadFileToSharepoint(ClientContext con, String FileLocation)
        {
            try
            {

                List list = con.Web.Lists.GetByTitle("Documents");
                con.Load(list);
                con.ExecuteQuery();
                Folder Root = list.RootFolder;
                con.Load(Root);
                con.ExecuteQuery();
                int last = FileLocation.LastIndexOf("\\");
                String Filename = FileLocation.Substring(last + 1);
                FileCreationInformation fci = new FileCreationInformation();
                fci.Content = System.IO.File.ReadAllBytes(FileLocation);
                fci.Url = Filename;
                fci.Overwrite = true;
                File fileToUpload = Root.Files.Add(fci);
                con.Load(fileToUpload);
                con.ExecuteQuery();
            }
            catch(Exception ex)
            {
                throw new Exception(ex.Message);
            }
           
        }
    }
}
