using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Windows.Forms;
using System.IO;
using OfficeOpenXml;
using SwDocumentMgr;

namespace SolidWorksPropertyAppendC
{
    class Program
    {
        const string TecDesc = "TechnicalDescription";
        const string Desc2 = "Description2";
        static ISwDMClassFactory ClasFact;
        static SwDMApplication4 dmDocMgr;


        static List<FileData> Errors = new List<FileData>();

        [STAThread]
        static void Main(string[] args)
        {
            ClasFact = new SwDMClassFactory();
            dmDocMgr = (SwDMApplication4)ClasFact.GetApplication(SwDocumentMgr_Data.sLicenseKey);

            OpenFileDialog OFD = new OpenFileDialog();

            OFD.Filter = "Excel (*.xlsx)|*.xlsx;";
            try
            {
                if (OFD.ShowDialog() == DialogResult.OK)
                {
                    if (File.Exists(OFD.FileName))
                    {
                        Console.WriteLine("Reading file....");
                        List<FileData> Data = ImportData(OFD.FileName);
                        if (Data.Count != 0)
                        {
                            foreach (FileData item in Data)
                            {
                                Console.WriteLine("Updating file : " + item.fileName);
                               UpdateFiles(item);
                            }
                        }
                    }
                    else
                    {
                        Console.WriteLine("File not valid....");
                    }
                }
                else
                {
                    Console.WriteLine("File not valid....");
                }
            }
            finally
            {
                if (Errors.Count != 0)
                {
                    Console.WriteLine("Errors :" + Errors.Count.ToString());
                    string Time = (DateTime.Now).ToString();
                    Time = Time.Replace("/", "-");
                    Time = Time.Replace(":", "-");
                    string FilePath = "C:\\Temp\\SolidworksAttsUpdate - " + Time + ".txt";
                    // System.IO.File.Create(FilePath);
                    using (StreamWriter sw = new StreamWriter(FilePath, false))
                    {
                        foreach (FileData item in Errors)
                        {
                            sw.WriteLine(item.fileName + "," + item.folderPath + "," + item.prefix + "," + item.result + "," + item.Message);
                        }                        
                    }
                    System.Diagnostics.Process.Start(FilePath);
                }
                else
                {
                    Console.WriteLine("No errors, Complete!");
                }
                Console.WriteLine("Push any key to close");
               Console.Read();
            }
        }


        public static List<FileData> ImportData(string fileName)
        {
            List<FileData> AllFiles = new List<FileData>();

            FileInfo FileIn = new FileInfo(fileName);

            using (ExcelPackage _App = new ExcelPackage(FileIn))
            {
                ExcelWorkbook Document = _App.Workbook;

                ExcelWorksheet MainSheet = Document.Worksheets[1];

                // File Name, Folder, Prefix
                const int File = 1;
                const int Folder = 2;
                const int prefix = 3;

                int Row = 2;
                try
                {
                    while (!string.IsNullOrWhiteSpace(Convert.ToString(MainSheet.Cells[Row, File].Value)))
                    {
                        try
                        {
                            FileData newFile = new FileData();
                            newFile.fileName = (Convert.ToString(MainSheet.Cells[Row, File].Value));
                            newFile.folderPath = (Convert.ToString(MainSheet.Cells[Row, Folder].Value));
                            newFile.prefix = (Convert.ToString(MainSheet.Cells[Row, prefix].Value));
                            AllFiles.Add(newFile);
                            Row += 1;
                        }
                        catch (Exception ex)
                        {
                            Console.WriteLine(ex.Message);
                        }
                    }
                }
                catch (Exception ex)
                {
                    Console.WriteLine(ex.Message);
                }
            }
            return AllFiles;
        }

        public static void UpdateFiles(FileData Info)
        {
            SwDmDocumentType Type;
            SwDmDocumentOpenError Errormsg;

            switch ((string)Path.GetExtension(Info.fileName).ToLower())
            {
                case ".sldprt":
                    Type = SwDmDocumentType.swDmDocumentPart;
                    break;
                case ".sldasm":
                    Type = SwDmDocumentType.swDmDocumentAssembly;
                    break;
                case ".slddrw":
                    Type = SwDmDocumentType.swDmDocumentDrawing;
                    break;
                default:
                    Console.WriteLine("Invalid file" + Info.fileName);
                    return;
            }
            SwDMDocument18 ModDoc = (SwDMDocument18)dmDocMgr.GetDocument(Info.folderPath + "\\" + Info.fileName, Type, false, out Errormsg);
            if (ModDoc != null && Errormsg == SwDmDocumentOpenError.swDmDocumentOpenErrorNone)
            {
                try
                {
                    DMCustomProperties Props = new DMCustomProperties(ModDoc);
                    if (Type == SwDmDocumentType.swDmDocumentPart | Type == SwDmDocumentType.swDmDocumentAssembly)
                    {
                        string Value = Props.DocumentPropertyRawValue(TecDesc);
                        if (!CheckIfAlreadyDone(Value))
                        {
                            Props.DocumentSetProperty(TecDesc, Info.prefix + " - " + Value);
                            Console.WriteLine(Info.prefix + " - " + Value);
                        }
                    }
                    if (Type == SwDmDocumentType.swDmDocumentDrawing)
                    {
                        string Value = Props.DocumentPropertyRawValue(Desc2);
                        if (!CheckIfAlreadyDone(Value))
                        {
                            Props.DocumentSetProperty(Desc2, Info.prefix + " - " + Value);
                            Console.WriteLine(Info.prefix + " - " + Value);
                        }
                    }
                    switch (ModDoc.Save())
                    {
                        case SwDmDocumentSaveError.swDmDocumentSaveErrorNone:
                            break;
                        case SwDmDocumentSaveError.swDmDocumentSaveErrorFail:
                            Info.result = false;
                            Info.Message = "Save Failed";
                            Errors.Add(Info);
                            break;
                        case SwDmDocumentSaveError.swDmDocumentSaveErrorReadOnly:
                            Info.result = false;
                            Info.Message = "Save failed : read only";
                            Errors.Add(Info);
                            break;
                    }
                }
                catch (Exception ex)
                {
                    Info.result = false;
                    Info.Message = ex.Message + "|" + ex.StackTrace;
                    Errors.Add(Info);
                }
                finally { ModDoc.CloseDoc(); }
            }
            else
            {
                Info.result = false;
                Info.Message = "Can not find or open file";
                Errors.Add(Info);
            }
        }
        public static Boolean CheckIfAlreadyDone(string Property)
        {
            try
            {
                string Atts = Property.Substring(0, 4);
                if (Convert.ToInt32(Atts).ToString() == Atts)
                {
                    return true;
                }
            }
            catch { }
            return false;
        }

    }
}
