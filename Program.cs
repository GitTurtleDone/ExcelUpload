using System;
using System.IO;
using OfficeOpenXml;

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
        
        string csvFilePath = @"/mnt/c/Users/Dell User/" + 
                            @"OneDrive - University of Canterbury/" + 
                            @"NZ2208/NghienCuu/SemSem/Projects/betaGa2O3MESFETs/230512_Fab230504to0607/230519_Fab230509/Dev02/A01/" +
                            @"I_V Diode Full wo PowComp.csv";
        string excelFilePath = @"/mnt/c/Users/Dell User/" +
                               @"OneDrive - University of Canterbury/" + 
                               @"NZ2208/NghienCuu/SemSem/Projects/betaGa2O3MESFETs/230512_Fab230504to0607/230519_Fab230509/Dev02/" +
                                @"IrOx230509_02_A01.xlsx";
        
        */
        string excelFilePath = @"IrOx230509_02_A01.xlsx";
        string csvFilePath = @"I_V Diode Full wo PowComp.csv";
        
        uploadToOneSheet (csvFilePath, excelFilePath, "Full", 7, 1);
    }
    static void uploadToOneSheet (string fCsvFilePath, string fExcelFilePath, string fSheetName, int fStartRow, int fStartCol)
    {
        
        
        
        // Console.WriteLine(csvFilePath);
        // Console.WriteLine(excelFilePath);
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

