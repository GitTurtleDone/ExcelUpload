using System;
using System.IO;
using OfficeOpenXml;
using OfficeOpenXml.Style;
using System.Globalization;
using System.Threading;




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
        //Upload to Device Folders started
        string[] arrDevFolderPaths = {@"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07",
                                      @"../230512_Fab230504to0607/230609_Fab230509IrOxNonRecess/Dev01", 
                                      @"../230512_Fab230504to0607/230606_Fab230509IrOxRecess/Dev02"};
        for (int i = 0; i < arrDevFolderPaths.Length; i++)
        {
            DeviceFolder devFolder = new DeviceFolder();
            devFolder.FolderPath = arrDevFolderPaths[i];
            devFolder.processDeviceFolder();
        }
        
        //Upload to Device Folders stopped
        
       
        /*
        comparisonFolder.uploadOneSheet((@"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07/E05/Fab230215IrOxNonRecess_Dev07_E05.xlsx", @"Rev500", 3, 1, 509,3),
                                         (@"../IrOxOASRecVsNonRec/E_500um.xlsx", @"Rev500I", 3,1,509,3));
        */
        
         
    
    
        /*
        
        // Process SBDFolder starts
        SBDFolder SBDFolder = new SBDFolder();
        SBDFolder.FolderPath = @"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07/E05";
        SBDFolder.SBDTemplateFolderPath = @"../SBDExcelTemplates";
        SBDFolder.getAllCsvFileNames();
        SBDFolder.processSBDFolder();
        // Process SBDFolder stops

        string csvFilePath = @"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07/E05" +
                            @"I_V Diode For Full After Rev500 wo PowComp [07 E5(4) ; 14_04_2023 4_49_12 p.m.].csv";
        MeasFile measFile = new MeasFile();
        

        //Comparison started
        ComparisonFolder comparisonFolder = new ComparisonFolder();
        //string[] comparisonFolderPaths = {@"../230512_Fab230504to0607/230606_Fab230509IrOxRecess/Dev02/", @"../230512_Fab230504to0607/230609_Fab230509IrOxNonRecess/Dev01/"};
        //comparisonFolder.ComparedFolderPaths = comparisonFolderPaths;
        comparisonFolder.ComparedFolderPaths = new string[] {@"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07",
                                                                @"../230512_Fab230504to0607/230609_Fab230509IrOxNonRecess/Dev01", 
                                                                    @"../230512_Fab230504to0607/230606_Fab230509IrOxRecess/Dev02"};
        comparisonFolder.TemplateFolderPath = @"../ComparisonExcelTemplates";
        comparisonFolder.FolderPath = @"../IrOxOASRecVsNonRec";
        comparisonFolder.SBDType = @"E";
        comparisonFolder.createWorkbook(true);
        comparisonFolder.getAllDeviceFolders();
        comparisonFolder.UploadMultipleSheets(@"../E_0500um.csv", 3, 130);
        //comparisonFolder.uploadOneSheet((@"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07/E05/Fab230215IrOxNonRecess_Dev07_E05.xlsx", @"Rev500", 3, 1, 509,3),
        //                                 (@"../IrOxOASRecVsNonRec/E_500um.xlsx", @"Rev500I", 3,1,509,3));
        //Comparison stopped
        
        //DeviceFolder devFolder = new DeviceFolder();     
        measFile.CsvFilePath = csvFilePath;
        measFile.analyzeFile();
        
        //string[] arrDevFolderPaths = {@"../230216_Fab230215/230414_Fab230215IrOxNonRecess/Dev07/"};
       
        

        
        

        */
        
        
          
    }

}

public class ComparisonFolder
{
    public string FolderPath {get; set;}
    public string[] ComparedFolderPaths { get; set;}
    public string TemplateFolderPath {get; set;}
    public string TemplatePath {get; set;}
    public string SBDType {get; set;}
    public string ComparisonUploadDetailsTemplatePath{get; set;}
    public List<string> ComparedWorkbookPaths {get; set;} = new List<string>();
    public List<DeviceFolder> DeviceFolders {get; set;} = new List<DeviceFolder>();

    public void getAllDeviceFolders()
    {
        if (ComparedFolderPaths.Length == 0)
        {
            Console.WriteLine("Please assign the paths to the ComparedFolderPaths");
        }
        else
        {
            if (DeviceFolders.Count == 0)
            {
                for (int i = 0; i < ComparedFolderPaths.Length; i++)
                {
                    DeviceFolder deviceFolder = new DeviceFolder();
                    deviceFolder.FolderPath = ComparedFolderPaths[i];
                    if (deviceFolder.AllSubFolderPaths.Count == 0) deviceFolder.getAllSubFolderPaths();
                    if (deviceFolder.AllSubFolderNames.Count == 0) deviceFolder.getAllSubFolderNames();
                    if (deviceFolder.AllRev500Dones.Count == 0) deviceFolder.getAllRev500Dones();
                    Console.WriteLine ($"{ComparedFolderPaths[i]} has {deviceFolder.AllRev500Dones.Count} subfolders ");
                    DeviceFolders.Add(deviceFolder);
                }
            }    
                
                
        }
    }
    public void FillUploadDetails()
    {
        if (ComparedFolderPaths.Length != 0)
        {
            for (int i = 0; i < ComparedFolderPaths.Length; i++)
            {
                DeviceFolder deviceFolder = new DeviceFolder();

            }
        }
        else
        {
            Console.WriteLine("Please assign all device folder paths");
        }
    }
    public void createWorkbook(bool OverrideDestination = true)
    {    
        if (OverrideDestination)
        {
            string[] templateFiles = Directory.GetFiles(TemplateFolderPath);
            foreach (string filePath in templateFiles)
            {
                if (SBDType == filePath.Substring(filePath.Length - 13, 1))
                {
                    // Copy the template to the folder
                    File.Copy(filePath, Path.Combine(FolderPath, Path.GetFileName(filePath)), true);
                    ComparedWorkbookPaths.Add(Path.Combine(FolderPath, Path.GetFileName(filePath)));
                    //Console.WriteLine(templatePath);
                }
            }
        }
    }
    public void UploadMultipleSheets(string ComparisonUploadDetailsTemplatePath, int startRow, int stopRow)
    {
    //
        //ExcelPackage compUploadDetPack = new ExcelPackage(new FileInfo(ComparisonUploadDetailsTemplatePath));
        List<(string, string, int, int, int, int)> sources = new List<(string, string, int, int, int, int)>();
        List <(string, string, int, int, int, int)> destinations = new List<(string, string, int, int, int, int)>();

        try
        {
            // Read the CSV file
            using (StreamReader reader = new StreamReader(ComparisonUploadDetailsTemplatePath))
            {
                try
                {
                    int row = 0;
                    
                    while (!reader.EndOfStream)
                    {
                        string line = reader.ReadLine();
                        string[] data = line.Split(',');
                        try
                        { 
                            if (row > 1) //((row >= 17) && (row < 18))
                            {
                                if ((row >= startRow -1) && (row <= stopRow -1))
                                {
                                    if (!string.IsNullOrEmpty(data[1]))
                                    {
                                        sources.Add((data[1], data[2], int.Parse(data[3]), int.Parse(data[4]), int.Parse(data[5]), int.Parse(data[6])));    
                                        destinations.Add((data[7], data[8], int.Parse(data[9]), int.Parse(data[10]), int.Parse(data[11]), int.Parse(data[12])));
                                    }
                                }
                                                  
                            }
                        }
                        catch (FormatException ex)
                        {
                            Console.WriteLine($"An error occurred: {ex.Message}");
                        }

                        
                        row++;
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

            //Console.WriteLine("Data successfully imported to Excel.");
        }
        catch (FileNotFoundException)
        {
            Console.WriteLine("CSV file not found!");
        }
        catch (IOException)
        {
            Console.WriteLine("An error occurred while reading or writing the csv file!");
        }

        
        for (int i = 0; i < sources.Count; i++)
        {
            Console.WriteLine($"Source: {sources[i]}, Destination: {destinations[i]}");
            uploadOneSheet(sources[i],destinations[i]);
        }
    }
    public void uploadOneSheet ((string, string, int, int, int, int) source, (string, string, int, int, int, int) destination)
    {
        try
        {
               

            // Load the source workbook
            using (ExcelPackage sourcePackage = new ExcelPackage(new FileInfo(source.Item1)))
            {
                sourcePackage.Workbook.CalcMode = ExcelCalcMode.Manual;
                // Get the source worksheet
                ExcelWorksheet sourceWorksheet = sourcePackage.Workbook.Worksheets[source.Item2];

                // Load the destination workbook
                using (ExcelPackage destinationPackage = new ExcelPackage(new FileInfo(destination.Item1)))
                {
                    destinationPackage.Workbook.CalcMode = ExcelCalcMode.Manual;
                    // Get the destination worksheet
                    ExcelWorksheet destinationWorksheet = destinationPackage.Workbook.Worksheets[destination.Item2];

                    // Define the range to copy in the source worksheet
                    ExcelRangeBase sourceRange = sourceWorksheet.Cells[source.Item3,source.Item4,source.Item5,source.Item6];

                    // Define the destination range in the destination worksheet
                    ExcelRangeBase destinationRange = destinationWorksheet.Cells[destination.Item3,destination.Item4,destination.Item5,destination.Item6];

                    if ((destinationWorksheet.Cells[destination.Item3-2,destination.Item4] == null || destinationWorksheet.Cells[destination.Item3-2,destination.Item4].Value == null || destinationWorksheet.Cells[destination.Item3-2,destination.Item4].Value is ExcelErrorValue))
                        destinationWorksheet.Cells[destination.Item3-2,destination.Item4].Value = "Source";
                    if ((destinationWorksheet.Cells[destination.Item3-1,destination.Item4] == null || destinationWorksheet.Cells[destination.Item3-1,destination.Item4].Value == null || destinationWorksheet.Cells[destination.Item3-1,destination.Item4].Value is ExcelErrorValue))
                        destinationWorksheet.Cells[destination.Item3-1,destination.Item4].Value = source.Item1;
                    // Copy the values and formatting from the source range to the destination range
                    /*
                    for (int i = source.Item3; i <= source.Item5; i++)
                    {
                        for (int j = source.Item4; j <= source.Item6; j++)
                        {
                            destinationWorksheet.Cells[destination.Item3 + i-source.Item3, destination.Item4 + j-source.Item4].Value = sourceWorksheet.Cells[i, j].Value;
                            
                            //Console.WriteLine($"Destination: {destinationWorksheet.Cells[destination.Item3 + i-source.Item3, destination.Item4 + j-source.Item4].Value}; Source: {sourceWorksheet.Cells[i, j].Value}");
                        }
                        
                    };
                    */
                    destinationRange.Value = sourceRange.Value;
                    //destinationRange.Style.Copy(sourceRange.Style);
                    // Save the changes to the destination workbook
                    sourcePackage.Workbook.CalcMode = ExcelCalcMode.Automatic;
                    destinationPackage.Workbook.CalcMode = ExcelCalcMode.Automatic;
                    destinationPackage.Save();
                    sourcePackage.Save();
                    
                    sourceRange.Dispose();
                    destinationRange.Dispose();
                    destinationWorksheet.Dispose();
                    sourceWorksheet.Dispose();
                    destinationPackage.Dispose();
            
                }
            sourcePackage.Dispose();
            //Thread.Sleep(5000);
            Console.WriteLine($"{source.Item1} successfully imported to the {destination.Item1}");
            }
        }
        catch (FileNotFoundException)
        {
            Console.WriteLine("CSV file not found!");
        }
        catch (IOException)
        {
            Console.WriteLine("An error occurred while reading or writing the excel files!");
        }
    
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

    public void processDeviceFolder(string fUploadDetailTemplatePath = "../UploadDetailsTemplate.csv", 
                                    string fSBDTemplateFolderPath = "../SBDExcelTemplates", bool OverrideDestination = true)
    {
        if (!string.IsNullOrEmpty(FolderPath))
        {
            getAllSubFolderPaths();
            //Console.WriteLine($"After getAllSubFolderPaths AllSubfolderNames.Count: {AllSubFolderNames.Count}, AllRevDones.Count {AllRev500Dones.Count}");
            getSampleID();
            getAllSubFolderNames();
            //Console.WriteLine($"After getAllSubFolderNames AllSubfolderNames.Count: {AllSubFolderNames.Count}, AllRevDones.Count {AllRev500Dones.Count}");
            getAllSBDFolders();
            //Console.WriteLine($"After getAllSBDFolders AllSubfolderNames.Count: {AllSubFolderNames.Count}, AllRevDones.Count {AllRev500Dones.Count}");
            getLogFilePath();
            getAllRev500Dones(); 
            //Console.WriteLine($"After getAllRev500Dones AllSubfolderNames.Count: {AllSubFolderNames.Count}, AllRevDones.Count {AllRev500Dones.Count}");
            getMemoFilePath();
            memoWrite(); 
            processAllSBDFolders(fUploadDetailTemplatePath, fSBDTemplateFolderPath, OverrideDestination);     
            
        }  
    }
    
    public void getAllSubFolderPaths()
    {
        if (!string.IsNullOrEmpty(FolderPath))
        {
            string[] allPaths = Directory.GetDirectories(FolderPath, "*", SearchOption.TopDirectoryOnly);
            foreach (string subfolder in allPaths)
            {
                AllSubFolderPaths.Add(subfolder);
                //Console.WriteLine($"Subfolder path: {subfolder}");
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
                //Console.WriteLine($"Folder Name: {AllSubFolderPaths[i].Substring(AllSubFolderPaths[i].Length-3)}");
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
                //Console.WriteLine($"Folder Name: {sbdFolder.FolderPath.Substring(AllSubFolderPaths[i].Length-3)}");
            }
        }
    }
    public void processAllSBDFolders(string fUploadDetailTemplatePath = "../UploadDetailsTemplate.csv",  
                                     string fSBDTemplateFolderPath = "../SBDExcelTemplates",
                                     bool OverrideDestination = true) 
    {
        if (AllSBDFolders.Count == 0)
        {
            Console.WriteLine("AllSBDFolders list is empty");
        }
        else
        {
            for (int i = 0; i < AllSBDFolders.Count; i++)
            {
                Console.WriteLine($"SBD Folder in process: {AllSubFolderNames[i]}");
                AllSBDFolders[i].processSBDFolder(fUploadDetailTemplatePath, fSBDTemplateFolderPath, OverrideDestination);
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
        //Console.WriteLine($"before getRev500Dones starts AllSubfolderNames.Count: {AllSubFolderNames.Count}, AllRevDones.Count {AllRev500Dones.Count}");
        //getAllSBDFolders();
        if (AllSBDFolders.Count == 0)
        {
            Console.WriteLine("AllSBDFolders list is empty");
        }
        else
        {
            if (AllRev500Dones.Count == 0)
            {
                for (int i = 0; i < AllSBDFolders.Count; i++)
                {
                    AllSBDFolders[i].getRev500Done();
                    AllRev500Dones.Add(AllSBDFolders[i].Rev500Done);
                    //Console.WriteLine($"{i}: {AllRev500Dones.Count}");
                }
            }
            
        }
        //Console.WriteLine($"after getAllRev500Dones finishes AllSubfolderNames.Count: {AllSubFolderNames.Count}, AllRevDones.Count {AllRev500Dones.Count}");
    }

    public void getMemoFilePath()
    {
        MemoFilePath = Path.Combine(FolderPath, "memo.txt");
    }

    public void memoWrite()
    {
        if (File.Exists(MemoFilePath)) File.Delete(MemoFilePath);
        if (AllRev500Dones.Count == 0) 
        {
            Console.WriteLine("AllRev500Dones list is empty");
        }
        else
        {
            for (int i = 0; i < AllRev500Dones.Count; i++)
            {
                Memo = $"{DateTime.Now}: {AllSubFolderNames[i]},{AllRev500Dones[i]}";
                File.AppendAllText(MemoFilePath, $"{Memo}\n");
                //Console.WriteLine(Memo);
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
    public string SBDTemplateFolderPath {get; set;}
    public List<string> AllCsvFileNames { get; set;} = new List<string>();
    public List<string> AllCsvFilePaths { get; set;} = new List<string>();
    public List<string> AllMeasTypes { get; set;} = new List<string>();
    public List<DateTime> AllMeasTimes { get; set;} = new List<DateTime>();
    public List<bool> AllIsLasts { get; set;} = new List<bool>();
    public List<(string, string, int, int)> AllUploadDetails { get; set;} = new List<(string, string, int, int)>();
    public List<(string, string, int, int)> UploadDetailTemplates { get; set;} = new List<(string, string, int, int)>();
    //public ExcelPackage Workbook;
    
    public void processSBDFolder(string fUploadDetailTemplatePath ="../UploadDetailsTemplate.csv",
                                 string fSBDTemplateFolderPath = "../SBDExcelTemplates",
                                 bool OverrideDestination = true)
    {
        if (!string.IsNullOrEmpty(FolderPath))
        {
            UploadDetailTemplatePath = fUploadDetailTemplatePath;
            SBDTemplateFolderPath = fSBDTemplateFolderPath;
            if (AllCsvFileNames.Count == 0) getAllCsvFileNames();
            if (AllMeasTimes.Count == 0) getAllMeasTimes();
            if (AllMeasTypes.Count == 0) getAllMeasTypes();
            if (AllIsLasts.Count == 0) getAllIsLasts();
            getSBDID();
            getSBDType();
            getSamplID();
            createWorkbook(OverrideDestination);
            getUploadDetailTemplate();
            getLogFilePath();
            getRev500Done();
            Upload();
        }  
    }
    public void getRev500Done()
    {
        if (AllMeasTypes.Count == 0)
        {     
            if (AllCsvFileNames.Count == 0)
                getAllCsvFileNames();
            getAllMeasTypes();
        }    
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
        //Console.WriteLine($"Folder Path: {FolderPath}; SBDID: {SBDID}");
    }
    public void getSBDType()
    {
        SBDType = FolderPath.Substring(FolderPath.Length-3,1);
        //Console.WriteLine(SBDType);
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
        //Console.WriteLine($"Measurement number: {(AllMeasTypes.Count-1).ToString()}");
        for (int i=0; i < AllMeasTypes.Count; i++)
        {
            AllIsLasts.Add(true); 
        }
        Console.WriteLine($"AllMeasTypes.Count: {AllMeasTypes.Count}, AllMeasTimes.Count: {AllMeasTimes.Count}");
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
            Console.WriteLine($"MeasType: {AllMeasTypes[i]}, MeasTime: {AllMeasTimes[i]}, IsLast: {AllIsLasts[i]}");
        }

        
    }
    private string getMeasType(string csvFileName) // get the measurement type for the file name
    {
        string measType;
        int index = csvFileName.IndexOf(@"[");
        if (index != -1)
        {
            measType = csvFileName.Substring(0, index-1);
            //Console.WriteLine(measType);
        }
        else if (csvFileName.Contains("Stop3"))
        {
            if (csvFileName.Contains("Rev500"))
                measType = "KL For After Rev500";
            else
                measType = "KL For";
        }
        else if (csvFileName.Contains("StopM3"))
        {
            if (csvFileName.Contains("Rev500"))
                measType = "KL Rev After Rev500";
            else
                measType = "KL Rev";
        }
        else if (csvFileName.Contains("StopM500"))
        {
            measType = "Pico Rev 500";
        }
        else measType = "No Type";

        //Console.WriteLine(measType);

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
            //Console.WriteLine(strMeasTime);
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
    public void createWorkbook(bool OverrideDestination = true)
    {
        string workbookID = SampleID + "_" + SBDID + ".xlsx";
        WorkbookPath = Path.Combine(FolderPath, workbookID);   
        
        if (OverrideDestination)
        {
            if (File.Exists(WorkbookPath)) 
            {
                Console.WriteLine("Workbook sussesfully deleted");
                File.Delete(WorkbookPath);
            }    
            string[] templateFiles = Directory.GetFiles(SBDTemplateFolderPath);
            foreach (string filePath in templateFiles)
            {
                if (SBDType == filePath.Substring(filePath.Length -8, 1))
                {
                    File.Copy(filePath, WorkbookPath, true);
                    Console.WriteLine($"Copy Template successfull to: {WorkbookPath} ");
                    break;
                    
                }
            }
            
        }
        // Copy the template to the folder
        
        //Console.WriteLine($"FolderPath: {FolderPath}; workbookID: {workbookID}");
        
        //Workbook = new ExcelPackage(new FileInfo(WorkbookPath));
        /*
       

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
            //Console.WriteLine($"FolderPath: {FolderPath}; workbookID: {workbookID}");
            File.Copy(templatePath, WorkbookPath);
            //Workbook = new ExcelPackage(new FileInfo(WorkbookPath));
            
            
        }
        */
        
    }
    
    public void Upload()
    {
        try
        {
            ExcelPackage package = new ExcelPackage (new FileInfo(WorkbookPath));
            package.Workbook.CalcMode = ExcelCalcMode.Manual;
            ExcelWorksheet worksheet = package.Workbook.Worksheets["Summary"];
            worksheet.Cells[5,2].Value = SBDID;
            package.Save();

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
                        worksheet.Dispose();
                        package.Dispose();
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
            reader.Dispose();
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
        catch (InvalidDataException ex)
        {
            Console.WriteLine($"File in Use: {ex.Message}");
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




