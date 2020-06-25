using System;
using Excel = Microsoft.Office.Interop.Excel;
using System.IO;
using System.Text;

namespace xl2txt
{
    class Program
    {
        public static int cntDir = 0;
        public static int cntFile = 0;
        public static Excel.Application xlapp = new Excel.Application();
        public static string filePath;

        static void Main(string[] path)
        {
            filePath = path[0]; //load first elemnt of array to string variable
            string[] subdir = Directory.GetDirectories(filePath);   //to read and to load the subfolders of the selected directory

            if (File.Exists(filePath + "FolderInfo.txt"))   //to check whether the 'FolderInfo.txt' has been already created. The FolderInfo.txt contains some useful data e.g. oroginal path, txt filename, MD5 hash etc..
            {
                File.Move((filePath + "FolderInfo.txt"), (filePath + "FolderInfo" + (File.GetLastWriteTime(filePath + "FolderInfo.txt")).ToString().Replace(" ",string.Empty).Replace(":", string.Empty).Replace(".", string.Empty)+".txt")); //if the FolderInfo exists, it will be renamed.
            }
            try
            {
                foreach (string subdirId in subdir)
                {
                    if (!File.Exists(subdirId + "\\processed")) //the "processed" file is the flag to check the processing status
                    {
                        string[] shortName = subdirId.Split('\\'); //array of pieces of path
                        cntDir++;   //counter for the processing status

                        Console.WriteLine();
                        Console.ForegroundColor = ConsoleColor.Yellow;
                        Console.WriteLine("Folder: {0} | {1}/{2} | {3}%", shortName[shortName.Length - 1], cntDir, subdir.Length, Convert.ToInt16(((float)cntDir / (float)subdir.Length) * 100));
                        Console.ForegroundColor = ConsoleColor.Gray;

                        File.Create(subdirId + "\\processed"); //if the folder has been processed the flag file is created

                        XlsConvert(Directory.GetFiles(subdirId));
                    }
                    else
                    {
                        Console.WriteLine("Folder has been already processed: {0}", subdirId);
                        //Console.ReadKey();
                    }
                }
                if (subdir.Length == 0) //
                {
                    if (!File.Exists(filePath + "\\processed"))
                    {
                        XlsConvert(Directory.GetFiles(filePath));
                        File.Create(filePath + "\\processed");
                    }
                    else
                    {
                        Console.WriteLine("Folder has been already processed: {0}", filePath);
                        //Console.ReadKey();
                    }
                }
                    //Console.ReadKey();
            }
            catch
            {
                Console.WriteLine("Incorrect path!");
                //Console.ReadKey();
            }
            xlapp.Quit();
            
        }

        static private void XlsConvert(string[] _xlsFiles)
        {
            string hash=string.Empty;
            string txtFile;
            StringBuilder dataRow = new StringBuilder();
            cntFile = 0;
            int xlsPartsCnt = _xlsFiles.Length;
            
            foreach (string xlsFile in _xlsFiles)
            {
                  //if (!(xlsFile.Substring(xlsFile.Length - 4, 4) == ".xls" || xlsFile.Substring(xlsFile.Length - 4, 4) == ".xml" || xlsFile.Substring(xlsFile.Length - 5, 5) == ".xlsx"))
                  if (!(xlsFile.Substring(xlsFile.Length - 4, 4) == ".xls" || xlsFile.Substring(xlsFile.Length - 5, 5) == ".xlsx"))
                    xlsPartsCnt--;
            }
            foreach (string xlsFile in _xlsFiles)
            {
                //if (xlsFile.Substring(xlsFile.Length - 4, 4) == ".xls" || xlsFile.Substring(xlsFile.Length - 4, 4) == ".xml" || xlsFile.Substring(xlsFile.Length - 5, 5) == ".xlsx")
                if (xlsFile.Substring(xlsFile.Length - 4, 4) == ".xls" || xlsFile.Substring(xlsFile.Length - 5, 5) == ".xlsx")
                {
                    cntFile++;
                    string[] file = xlsFile.Split('\\');
                    Console.ForegroundColor = ConsoleColor.Green;
                    Console.WriteLine("\nFile: {0} | {1}/{2} | {3}%", file[file.Length - 1], cntFile, xlsPartsCnt, Convert.ToInt16(((float)cntFile / (float)xlsPartsCnt) * 100));
                    Console.ForegroundColor = ConsoleColor.Gray;
                    try
                    {
                        hash = FileHash(xlsFile);
                        Excel.Workbook workbook = xlapp.Workbooks.Open(xlsFile);
                        foreach (Excel.Worksheet worksheet in workbook.Sheets)
                        {
                            txtFile = Path.GetDirectoryName(xlsFile) + "\\" + worksheet.Name.ToString() + "_" +
                                (Path.GetFileName(xlsFile)).
                                    Replace(".xlsx", "").
                                    Replace(".xls", "").
                                    Replace(".xml", "") + ".txt";
                            int rowAmount = worksheet.UsedRange.Rows.Count;

                            dataRow.Append("\"");
                            dataRow.Append(xlsFile);
                            dataRow.Append("\"");
                            dataRow.Append("\t");
                            dataRow.Append("\"");
                            dataRow.Append(hash);
                            dataRow.Append("\"");
                            dataRow.Append("\t");
                            dataRow.Append("\"");
                            dataRow.Append(worksheet.Name.ToString());
                            dataRow.Append("\"");
                            dataRow.Append("\t");
                            dataRow.Append("\"");
                            dataRow.Append(rowAmount);
                            dataRow.Append("\"");
                            dataRow.Append("\t");
                            dataRow.Append("\"");
                            dataRow.Append(txtFile);
                            dataRow.Append("\"");
                            dataRow.Append("\t");
                            dataRow.Append("\"");
                            dataRow.Append(Environment.UserName);
                            dataRow.Append("\"");
                            dataRow.Append("\t");
                            dataRow.Append("\"");
                            dataRow.Append(DateTime.Now.ToString());
                            dataRow.Append("\"");

                            Console.WriteLine(worksheet.Name.ToString());
                            FolderInfo(dataRow);
                            dataRow.Clear();

                            worksheet.SaveAs(txtFile, Excel.XlFileFormat.xlTextWindows);
                        }
                        workbook.Close(0);
                    }
                    catch
                    {
                        Console.ForegroundColor = ConsoleColor.Red;
                        Console.WriteLine("\nCorrupt source file: {0}\n", xlsFile);
                        Console.ForegroundColor = ConsoleColor.Gray;
                        xlapp.Quit();
                    }
                }
            }
        }

        static private void FolderInfo(StringBuilder _dataRow)
        {
            using (StreamWriter writer = new StreamWriter(filePath + "FolderInfo.txt", true))
            {
                writer.WriteLine(_dataRow);
            }
        }

        static private string FileHash(string _xlsFileToHash)
        {
            using (var md5 = System.Security.Cryptography.MD5.Create())
            {
                using (var dataStream = File.OpenRead(_xlsFileToHash))
                {
                    return BitConverter.ToString(md5.ComputeHash(dataStream)).Replace("-",string.Empty);
                }
            }
                
        }
    }
}
