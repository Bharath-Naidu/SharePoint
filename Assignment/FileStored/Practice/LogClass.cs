using System;
using System.IO;

namespace Practice
{
    static class LogClass
    {
        public static void RecordException(Exception e)
        {
            string message = "--------" + DateTime.Now + Environment.NewLine + "--------" + e.StackTrace + Environment.NewLine + "--------" + e.Message + "--------" + Environment.NewLine+ Environment.NewLine+ Environment.NewLine + Environment.NewLine+ Environment.NewLine;
            string path="C:\\Users\\bharat.naidu\\source\\repos\\SharePoint\\Assignment\\FileStored\\Practice\\LogFile.txt";
            File.AppendAllText(path,message);
        }
    }
}
