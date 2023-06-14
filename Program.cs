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
        /*
        
        string csvFilePath = @"I_V Diode Full wo PowComp.csv";
        string excelFilePath = @"/mnt/c/Users/Dell User/" +
                               @"OneDrive - University of Canterbury/" + 
                               @"NZ2208/NghienCuu/SemSem/Projects/betaGa2O3MESFETs/230512_Fab230504to0607/230519_Fab230509/Dev02/" +
                                @"IrOx230509_02_A01.xlsx";
        
        */
        string excelFilePath = @"IrOx230509_02_A01.xlsx";
        string csvFilePath = @"../230512_Fab230504to0607/230519_Fab230509/Dev02/A01/" +
                            @"I_V Diode Full wo PowComp [02 A01(2) ; 19_05_2023 1_41_45 p.m.].csv";
        
        
        
        //uploadToOneSheet (csvFilePath, excelFilePath, "Full", 7, 1);
        MeasFile measFile = new MeasFile();
        SBDFolder SBDFolder = new SBDFolder();
        
        measFile.CsvFilePath = csvFilePath;
        measFile.analyzeFile();
        SBDFolder.FolderPath = measFile.CsvFolderPath;
        SBDFolder.getAllCsvFileNames();
        SBDFolder.getAllMeasTimes();
        SBDFolder.getAllMeasTypes();
        SBDFolder.getAllIsLasts();
        SBDFolder.getSBDID();
        SBDFolder.getSBDType();
        SBDFolder.getSamplID();
        SBDFolder.createWorkbook();
        // DateTime currentDateTime = DateTime.Now;
        // string formattedDateTime = currentDateTime.ToString();  // Format based on system settings
        // Console.WriteLine(formattedDateTime);

        
        //Console.WriteLine($"Here is the folder path: {measFile.CsvFolderPath} ");
        
    }
    
    /*
    public static Person GetPersonDetails()
    {
        int age = 30;
        string name = "John Doe";
        return new Person { Age = age, Name = name };
    }

    // Usage
    Person person = GetPersonDetails();
    int age = person.Age;
    string name = person.Name;
    */
    static void uploadToMultiSheets (string[] fCsvFilePaths, string fExcelFilePath, string[] fSheetNames, int[] fStartRows, int[] fStartCols)
    {
        Console.WriteLine("Went in");
        for (int i=0; i < fCsvFilePaths.Length; i++)
        {
            uploadToOneSheet(fCsvFilePaths[i], fExcelFilePath, fSheetNames[i], fStartRows[i], fStartRows[i]);
        }
        
    }
    static void uploadToOneSheet (string fCsvFilePath, string fExcelFilePath, string fSheetName, int fStartRow, int fStartCol)
    {
        // Console.WriteLine(csvFilePath);
        // Console.WriteLine(excelFilePath);
        // fCsvFilePath = $"\"{fCsvFilePath}\"";
        // fExcelFilePath = $"\"{fExcelFilePath}\"";
        try
        {
            // Read the CSV file
            using (StreamReader reader = new StreamReader(fCsvFilePath))
            {
                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(fExcelFilePath)))
                    {
                        /*
                        ExcelWorkbook workbook = package.Workbook;
                        if (workbook == null)
                        {
                            Console.WriteLine("Workbook not found");
                            return;
                        }
                        else 
                        {
                            Console.WriteLine("I have gone in here");
                            var worksheets = package.Workbook.Worksheets.Select(x => x.Name);
                            //foreach (var sheet in worksheets) Console.WriteLine(sheet);
                            foreach (var ws in worksheets)
                                {
                                    Console.WriteLine("hellow");
                                    Console.WriteLine(ws);
                                }
                        }
                        */
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
                                worksheet.Cells[row, col + fStartCol].Value = data[col];
                            }

                            row++;
                        }

                        // Save the Excel file
                        FileInfo excelFile = new FileInfo(fExcelFilePath);
                        package.SaveAs(excelFile);
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
public class SBDFolder
{
    public string FolderPath { get; set;}
    public string SampleID { get; set;}
    public string SBDID { get; set;}
    public string SBDType { get; set;}
    public string[] AllFilePaths { get; set;}
    public List<string> AllCsvFileNames { get; set;} = new List<string>();
    public List<string> AllCsvFilePaths { get; set;} = new List<string>();
    public List<string> AllMeasTypes { get; set;} = new List<string>();
    public List<DateTime> AllMeasTimes { get; set;} = new List<DateTime>();
    public List<bool> AllIsLasts { get; set;} = new List<bool>();
    public ExcelWorkbook Workbook;
    public string WorkbookPath { get; set; }
    
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
        }
        else strMeasTime = "";
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

            Console.WriteLine($"ap: {ap} hour: {strHour} measureTime: {strMeasTime}");
            
            //Console.WriteLine(strMeasTime);
            //measTime=DateTime.Parse(strMeasTime);
            measTime = DateTime.ParseExact(strMeasTime,dateTimeFormat, CultureInfo.InvariantCulture);
            //ParseExact(strMeasTime, dateTimeFormat);
        }
        catch (FormatException ex)
        {
            
            Console.WriteLine($"Invalid date format: {ex.Message}");
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
            string templatePath = "../SBDExcelTemplates";  // Path to the template file
            
            string[] templateFiles = Directory.GetFiles(templatePath);
            foreach (string filePath in templateFiles)
            {
                if (SBDType == filePath.Substring(filePath.Length -8, 1))
                {
                    templatePath = filePath;
                    Console.WriteLine(templatePath);
                }
            }

            // Copy the template to the folder
            string workbookID = SampleID + "_" + SBDID + ".xlsx";
            WorkbookPath = Path.Combine(FolderPath, workbookID);
            File.Copy(templatePath, WorkbookPath);
            
        }
    }
    
}




