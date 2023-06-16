using System;
using System.IO;
using OfficeOpenXml;
using System.Globalization;



// See https://aka.ms/new-console-template for more information
//Console.WriteLine("Hello, World! My name is Giang");



/*
string filePath = "C:\\Users\\gtd19\\" + 
                            "OneDrive - University of Canterbury\\" + 
                            "NZ2208\\NghienCuu\\SemSem\\Projects\\betaGa2O3MESFETs\\230512_Fab230504to0607\\230519_Fab230509\\Dev02\\" +
                            "I_V Diode Full wo PowComp [02 A01(2) ; 19_05_2023 1_41_45 p.m.].csv";

Console.WriteLine(filePath);
*/

class Program
{
    static void Main()
    {
        
        string csvFilePath = @"../230512_Fab230504to0607/230606_Fab230509IrOxRecess/Dev02/A01/" +
                            @"I_V Diode Full wo PowComp [02 A01(2) ; 19_05_2023 1_41_45 p.m.].csv";
        MeasFile measFile = new MeasFile();
        SBDFolder SBDFolder = new SBDFolder();
        DeviceFolder devFolder = new DeviceFolder();
        /*
        measFile.CsvFilePath = csvFilePath;
        measFile.analyzeFile();
        SBDFolder.FolderPath = measFile.CsvFolderPath;
        SBDFolder.processSBDFolder();        
        */
        devFolder.FolderPath = @"../230512_Fab230504to0607/230606_Fab230509IrOxRecess/Dev02/";
        devFolder.getAllSubFolderPaths();
        devFolder.getSampleID();
        devFolder.getAllSubFolderNames();
        devFolder.getAllSBDFolders();
        devFolder.getAllRev500Dones();
        devFolder.getLogFilePath();
        devFolder.getMemoFilePath();
        devFolder.memoWrite();
        devFolder.processAllSBDFolders();        
    }

}
public class DeviceFolder

{
    public string FolderPath { get; set;}
    
    public string AllSBDIDs { get; set;}
    public string AllSBDTypes { get; set;}
    public string SampleID { get; set;}
    public string LogFilePath {get; set;}
    public string LogMessage {get; set;}
    public string MemoFilePath {get;set;} 
    public string Memo {get; set;}
    public List<SBDFolder> AllSBDFolders { get; set;} = new List<SBDFolder>();
    public List<string> AllSubFolderPaths { get; set;} = new List<string>();
    public List<string> AllSubFolderNames { get; set;} = new List<string>();
    public List<bool> AllRev500Dones { get; set;} = new List<bool>();
    
    public void getAllSubFolderPaths()
    {
        if (!string.IsNullOrEmpty(FolderPath))
        {
            string[] allPaths = Directory.GetDirectories(FolderPath, "*", SearchOption.AllDirectories);
            foreach (string subfolder in allPaths)
            {
                AllSubFolderPaths.Add(subfolder);
                Console.WriteLine($"Subfolder path: {subfolder}");
            }
        }
    }
    public void getSampleID()
    {
        if (!string.IsNullOrEmpty(FolderPath))
        {
            SampleID = FolderPath.Substring(FolderPath.LastIndexOf("_")+1, FolderPath.LastIndexOf("/")-FolderPath.LastIndexOf("_")-1).Replace("/","_");
            Console.WriteLine($"Sample ID: {SampleID}");
        }
    }
    public void getAllSubFolderNames()
    {
        if (AllSubFolderPaths.Count == 0)
        {
            Console.WriteLine("AllSubFolderPaths list is empty");
        } 
        else
        {
            for (int i = 0; i < AllSubFolderPaths.Count; i++)
            {
                Console.WriteLine($"Folder Name: {AllSubFolderPaths[i].Substring(AllSubFolderPaths[i].Length-3)}");
                AllSubFolderNames.Add(AllSubFolderPaths[i].Substring(AllSubFolderPaths[i].Length-3));
            }
        }
    }

    public void getAllSBDFolders()
    {
        if (AllSubFolderPaths.Count == 0)
        {
            Console.WriteLine("AllSubFolderPaths list is empty");
        } 
        else
        {
            for (int i = 0; i < AllSubFolderPaths.Count; i++)
            {
                SBDFolder sbdFolder = new SBDFolder();
                sbdFolder.FolderPath = AllSubFolderPaths[i];
                AllSBDFolders.Add(sbdFolder);
                Console.WriteLine($"Folder Name: {sbdFolder.FolderPath.Substring(AllSubFolderPaths[i].Length-3)}");
            }
        }
    }
    public void processAllSBDFolders()
    {
        if (AllSBDFolders.Count == 0)
        {
            Console.WriteLine("AllSBDFolders list is empty");
        }
        else
        {
            for (int i = 0; i < AllSBDFolders.Count; i++)
            {
                AllSBDFolders[i].processSBDFolder();
                LogMessage = $"Folder {AllSubFolderNames[i]} has been processed";
                logWrite();
            }
        }
    }
    public void getLogFilePath()
    {
        LogFilePath = Path.Combine(FolderPath, "log.txt");
    }

    public void logWrite()
    {
        if (!string.IsNullOrEmpty(LogFilePath)) File.AppendAllText(LogFilePath, $"{DateTime.Now}: {LogMessage}\n");
    }
    public void getAllRev500Dones()
    {
        if (AllSBDFolders.Count == 0)
        {
            Console.WriteLine("AllSBDFolders list is empty");
        }
        else
        {
            for (int i = 0; i < AllSBDFolders.Count; i++)
            {
                AllSBDFolders[i].getRev500Done();
                AllRev500Dones.Add(AllSBDFolders[i].Rev500Done);
            }
        }
    }

    public void getMemoFilePath()
    {
        MemoFilePath = Path.Combine(FolderPath, "memo.txt");
    }

    public void memoWrite()
    {
        if (AllRev500Dones.Count == 0) 
        {
            Console.WriteLine("AllRev500Dones list is empty");
        }
        else
        {
            for (int i = 0; i < AllRev500Dones.Count; i++)
            {
                Memo = $"{AllSubFolderNames[i]},{AllRev500Dones[i]}";
                File.AppendAllText(MemoFilePath, $"{Memo}\n");
                Console.WriteLine(Memo);
            }
            
        }
    }

}    

public class SBDFolder
{
    public string FolderPath { get; set;}
    public string WorkbookPath { get; set; }
    public string UploadDetailTemplatePath{ get; set;}
    public string LogFilePath {get; set;}
    public string LogMessage {get; set;}
    public string SampleID { get; set;}
    public string SBDID { get; set;}
    public string SBDType { get; set;}
    public string[] AllFilePaths { get; set;}
    public bool Rev500Done {get; set;}
    public List<string> AllCsvFileNames { get; set;} = new List<string>();
    public List<string> AllCsvFilePaths { get; set;} = new List<string>();
    public List<string> AllMeasTypes { get; set;} = new List<string>();
    public List<DateTime> AllMeasTimes { get; set;} = new List<DateTime>();
    public List<bool> AllIsLasts { get; set;} = new List<bool>();
    public List<(string, string, int, int)> AllUploadDetails { get; set;} = new List<(string, string, int, int)>();
    public List<(string, string, int, int)> UploadDetailTemplates { get; set;} = new List<(string, string, int, int)>();
    //public ExcelPackage Workbook;
    
    public void processSBDFolder()
    {
        if (!string.IsNullOrEmpty(FolderPath))
        {
            UploadDetailTemplatePath = "../UploadDetailsTemplate.csv";
            getAllCsvFileNames();
            getAllMeasTimes();
            getAllMeasTypes();
            getAllIsLasts();
            getSBDID();
            getSBDType();
            getSamplID();
            createWorkbook();
            getUploadDetailTemplate();
            getLogFilePath();
            getRev500Done();
            Upload();
        }  
    }
    public void getRev500Done()
    {
        if (AllMeasTypes.Contains("Pico Rev 500")) 
            Rev500Done = true; 
        else 
            Rev500Done = false;
    }
    public void getSamplID()
    {
        SampleID = FolderPath.Substring(FolderPath.LastIndexOf("_")+1, FolderPath.LastIndexOf("/")-FolderPath.LastIndexOf("_")-1).Replace("/","_");
        //Console.WriteLine(SampleID);
    }
    public void getSBDID()
    {
        SBDID = FolderPath.Substring(FolderPath.Length-3);
        //Console.WriteLine(SBDID);
    }
    public void getSBDType()
    {
        SBDType = FolderPath.Substring(FolderPath.Length-3,1);
        Console.WriteLine(SBDType);
    }
    public void getLogFilePath()
    {
        LogFilePath = Path.Combine(FolderPath, "log.txt");
    }

    public void logWrite()
    {
        if (!string.IsNullOrEmpty(LogFilePath)) File.AppendAllText(LogFilePath, $"{DateTime.Now}: {LogMessage}\n");
    }
    public void getAllCsvFileNames()
    {
        AllFilePaths =  Directory.GetFiles(FolderPath); 
        foreach (string filePath in AllFilePaths)
        {
            string fileName = Path.GetFileName(filePath);
            if (fileName.Contains("csv")) 
            {
                //Console.WriteLine(fileName);
                AllCsvFileNames.Add(fileName);
                AllCsvFilePaths.Add(filePath);
            }
        }
    }
    public void getAllMeasTypes()
    {
        foreach (string fileName in AllCsvFileNames)
        {
            AllMeasTypes.Add(getMeasType(fileName));
        }
    }

    public void getAllMeasTimes()
    {   
        foreach (string fileName in AllCsvFileNames)
        {
            AllMeasTimes.Add(getMeasTime(fileName));
        }
    }
    public void getAllIsLasts()
    {   
        //List<int> lstIgnoredIndices = new List<int>();
        Console.WriteLine($"Measurement number: {(AllMeasTypes.Count-1).ToString()}");
        for (int i=0; i < AllMeasTypes.Count; i++)
        {
            AllIsLasts.Add(true); 
        }
        
        for (int i=0; i < AllMeasTypes.Count; i++)
        {
            for (int j=i+1; j < AllMeasTypes.Count; j++)
            {
                if (AllMeasTypes[j] == AllMeasTypes[i])
                {
                    //lstIgnoredIndices.Add(j);
                    if (AllMeasTimes[j] < AllMeasTimes[i])
                    {
                        AllIsLasts[j] = false;
                    }
                    else
                    {
                        AllIsLasts[i] = false;
                    }

                }
            } 
        }
        for (int i=0; i < AllMeasTypes.Count; i++)
        {
            Console.WriteLine(AllIsLasts[i]);
        }

        
    }
    private string getMeasType(string csvFileName) // get the measurement type for the file name
    {
        string measType;
        int index = csvFileName.IndexOf(@"[");
        if (index != -1)
        {
            measType = csvFileName.Substring(0, index-1);
            Console.WriteLine(measType);
        }
        else if (csvFileName.Contains("Stop3"))
        {
            measType = "KL For";
        }
        else if (csvFileName.Contains("StopM3"))
        {
            measType = "KL Rev";
        }
        else if (csvFileName.Contains("StopM500"))
        {
            measType = "Pico Rev 500";
        }
        else measType = "No Type";

        Console.WriteLine(measType);

        return measType;

    }
    private DateTime getMeasTime(string csvFileName)
    {
        string strMeasTime;
        DateTime measTime = DateTime.MinValue;
        string dateTimeFormat = "dd/MM/yyyy HH:mm:ss";
        int indexCloseBracket = csvFileName.IndexOf(@"]");
        int indexSemicolon = csvFileName.IndexOf(@";");
        if ( indexCloseBracket!= -1)
        {
            strMeasTime = csvFileName.Substring(indexSemicolon+2,indexCloseBracket-indexSemicolon-2);
            Console.WriteLine(strMeasTime);
        try 
        {
            
            string ap = "a";
            string strHour;
            int firstDotIndex = strMeasTime.IndexOf(@".");
            ap = strMeasTime.Substring(firstDotIndex-1,1);
            strHour = strMeasTime.Substring(firstDotIndex-10,2);
            if (ap=="p" && strHour != "12")
            {
                strHour = (int.Parse(strHour) + 12).ToString();
                if (strMeasTime[firstDotIndex-10] == ' ')
                {
                    strMeasTime = strMeasTime.Remove(firstDotIndex-9,1);
                    strMeasTime = strMeasTime.Insert(firstDotIndex-9, strHour);
                }
                else
                {
                    strMeasTime = strMeasTime.Remove(firstDotIndex-10,2);
                    strMeasTime = strMeasTime.Insert(firstDotIndex-10, strHour);
                }
            }
            strMeasTime = strMeasTime.Substring(0,strMeasTime.Length-5);
            int strMeasTimeLength = strMeasTime.Length;
            strMeasTime = strMeasTime.Remove(strMeasTimeLength-3,1);
            strMeasTime = strMeasTime.Insert(strMeasTimeLength-3,":");
            strMeasTime = strMeasTime.Remove(strMeasTimeLength-6,1);
            strMeasTime = strMeasTime.Insert(strMeasTimeLength-6,":");
            strMeasTime = strMeasTime.Replace("_", "/");

                //Console.WriteLine($"ap: {ap} hour: {strHour} measureTime: {strMeasTime}");
            
            //Console.WriteLine(strMeasTime);
            //measTime=DateTime.Parse(strMeasTime);
            measTime = DateTime.ParseExact(strMeasTime,dateTimeFormat, CultureInfo.InvariantCulture);
            //ParseExact(strMeasTime, dateTimeFormat);
        }
        catch (FormatException ex)
        {
            
            Console.WriteLine($"Invalid date format: {ex.Message}");
        }
        }
        else 
        {
            measTime = File.GetCreationTime(csvFileName);
            //Console.Write(measTime.ToString());
        }
        return measTime;    
    }
    public void createWorkbook()
    {
        string[] files = Directory.GetFiles(FolderPath);

        foreach (string file in files)
        {
            string extension = Path.GetExtension(file);

            if (extension == ".xlsx" || extension == ".xls")
            {
                WorkbookPath = file;
                break;
            }
        }

        if (WorkbookPath == null)
        {
            // Workbook doesn't exist, create one from the template
            string templatePath = "../SBDExcelTemplates";  // Path to the template folder
            
            string[] templateFiles = Directory.GetFiles(templatePath);
            foreach (string filePath in templateFiles)
            {
                if (SBDType == filePath.Substring(filePath.Length -8, 1))
                {
                    templatePath = filePath;
                    //Console.WriteLine(templatePath);
                }
            }

            // Copy the template to the folder
            string workbookID = SampleID + "_" + SBDID + ".xlsx";
            WorkbookPath = Path.Combine(FolderPath, workbookID);
            File.Copy(templatePath, WorkbookPath);
            //Workbook = new ExcelPackage(new FileInfo(WorkbookPath));
            
        }
        
    }
    
    public void Upload()
    {
        try
        {
            for (int i=0; i < AllIsLasts.Count; i++)
            {
                if ((AllIsLasts[i]) && (AllMeasTypes[i] != "No Type"))
                {
                    //Console.WriteLine(AllMeasTypes[i]);
                    for (int j=0; j < UploadDetailTemplates.Count; j++)
                    {
                        if (AllMeasTypes[i] == UploadDetailTemplates[j].Item1)
                        {
                            AllUploadDetails.Add(UploadDetailTemplates[j]);
                            Console.WriteLine($"data to upload: {AllCsvFileNames[i]}, sheet: {UploadDetailTemplates[j].Item2}");
                            uploadToOneSheet(WorkbookPath,Path.Combine(FolderPath,AllCsvFileNames[i]),UploadDetailTemplates[j].Item2,
                                            UploadDetailTemplates[j].Item3,
                                            UploadDetailTemplates[j].Item4);
                            logWrite();
                        }
                    }
                }
            }

        }
        catch(ArgumentOutOfRangeException ex)
        {
            // Handle the exception when the index is out of range
            Console.WriteLine($"List indices not correct {ex.Message}");
        }
        
        /*
        string messString = $"AllCsvFileNames length: {AllCsvFileNames.Count.ToString()}\n" +
                            $"AllCsvFilePaths length: {AllCsvFilePaths.Count.ToString()}\n" +
                            $"AllCsvFileNames length: {AllCsvFileNames.Count.ToString()}\n" +
                            $"AllMeasTypes length: {AllMeasTypes.Count.ToString()}\n" +
                            $"AllMeasTimes length: {AllMeasTimes.Count.ToString()}\n" +
                            $"AllIsLasts length: {AllIsLasts.Count.ToString()}\n";
        Console.WriteLine(messString);*/
    }
    public void getUploadDetailTemplate()
    {
        using (StreamReader reader = new StreamReader(UploadDetailTemplatePath))
            {
                try
                {
                    
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] data = line.Split(',');
                        UploadDetailTemplates.Add((data[0], data[1], int.Parse(data[2]), int.Parse(data[3])));
                    }
                    
                }
                catch (FileNotFoundException)
                {
                    Console.WriteLine("Excel file not found");
                }
                catch (IOException)
                {
                    Console.WriteLine("An error occurred while reading or writing the excel file");
                }
        }
    }
    public void uploadToMultiSheets (string fExcelFilePath, string[] fCsvFilePaths,  string[] fSheetNames, int[] fStartRows, int[] fStartCols)
    {
        Console.WriteLine("Went in");
        for (int i=0; i < fCsvFilePaths.Length; i++)
        {
            uploadToOneSheet(fExcelFilePath, fCsvFilePaths[i],  fSheetNames[i], fStartRows[i], fStartRows[i]);
        }
    
}
    public void uploadToOneSheet (string fExcelFilePath, string fCsvFilePath, string fSheetName, int fStartRow, int fStartCol)
    {
        try
        {
            // Read the CSV file
            using (StreamReader reader = new StreamReader(fCsvFilePath))
            {
                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(fExcelFilePath)))
                    {
                        package.Workbook.CalcMode = ExcelCalcMode.Manual;
                        ExcelWorksheet worksheet = package.Workbook.Worksheets[fSheetName];
                        if (worksheet == null)
                        {
                            Console.WriteLine("Worksheet not found!");
                            return;
                        }
                        int row = fStartRow;
                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine();
                            string[] data = line.Split(',');

                            for (int col = 0; col < data.Length; col++)
                            {
                                bool success = double.TryParse(data[col], out double result);
                                if (success)
                                {
                                    worksheet.Cells[row, col + fStartCol].Value = result;
                                }
                                else
                                {
                                    worksheet.Cells[row, col + fStartCol].Value = data[col];
                                }
                                
                            }

                            row++;
                        }

                        // Save the Excel file
                        //FileInfo excelFile = new FileInfo(fExcelFilePath);
                        LogMessage = $"Data file: {Path.GetFileName(fCsvFilePath)} has been successfully imported to the sheet: {fSheetName} in {Path.GetFileName(fExcelFilePath)}";
                        package.Workbook.CalcMode = ExcelCalcMode.Automatic;
                        package.Save();
                    }
                }
                catch (FileNotFoundException)
                {
                    Console.WriteLine("Excel file not found");
                }
                catch (IOException)
                {
                    Console.WriteLine("An error occurred while reading or writing the excel file");
                }
            }

            Console.WriteLine("Data successfully imported to Excel.");
        }
        catch (FileNotFoundException)
        {
            Console.WriteLine("CSV file not found!");
        }
        catch (IOException)
        {
            Console.WriteLine("An error occurred while reading or writing the csv file!");
        }
    
    }
    
}


public class MeasFile
{
    public string CsvFilePath { get; set; }
    public string CsvFolderPath{get; set;}
    public string CsvFileName{get; set;}
    public string MeasType { get; set; } // Type of measuremnt
    public void analyzeFile() // get the file name and the folder path from the file path
    {
        
        if (!string.IsNullOrEmpty(CsvFilePath))
        {
            if (string.IsNullOrEmpty(CsvFileName)) CsvFileName = Path.GetFileName(CsvFilePath);
            if (string.IsNullOrEmpty(CsvFolderPath)) CsvFolderPath = Path.GetDirectoryName(CsvFilePath);
            Console.WriteLine(CsvFileName);
            Console.WriteLine(CsvFolderPath);
        }
        else
        {
            Console.WriteLine("File path is empty or null.");
        }
            
    } 
    public void getMeasurementType() // get the measurement type fromt the file name
    {
        int index = CsvFileName.IndexOf(@"[");
        if (index != -1)
        {
            MeasType = CsvFileName.Substring(0, index-1);
            Console.WriteLine(MeasType);
        }
        else if (CsvFileName.Contains("Stop3"))
        {
            MeasType = "KL For";
        }
        else if (CsvFileName.Contains("StopM3"))
        {
            MeasType = "KL Rev";
        }
        else if (CsvFileName.Contains("StopM500"))
        {
            MeasType = "Pico Rev 500";
        }
        else MeasType = "No Type";

    }
    
}




