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
        string csvFilePath = @"/mnt/c/Users/gtd19/" + 
                            @"OneDrive - University of Canterbury/" + 
                            @"NZ2208/NghienCuu/SemSem/Projects/betaGa2O3MESFETs/230512_Fab230504to0607/230519_Fab230509/Dev02/A01/" +
                            @"I_V Diode Full wo PowComp.csv";
        string excelFilePath = @"/mnt/c/Users/gtd19/" +
                               @"OneDrive - University of Canterbury/" + 
                               @"NZ2208/NghienCuu/SemSem/Projects/betaGa2O3MESFETs/230512_Fab230504to0607/230519_Fab230509/Dev02/" +
                                @"IrOx230509_02_A01.xlsx";
        Console.WriteLine(csvFilePath);
        Console.WriteLine(excelFilePath);
        try
        {
            // Read the CSV file
            using (StreamReader reader = new StreamReader(csvFilePath))
            {
                try
                {
                    using (ExcelPackage package = new ExcelPackage(new FileInfo(excelFilePath)))
                    {
                        ExcelWorksheet worksheet = package.Workbook.Worksheets["Full"];
                        if (worksheet == null)
                        {
                            Console.WriteLine("Worksheet not found!");
                            return;
                        }
                        int row = 7;
                        while (!reader.EndOfStream)
                        {
                            string line = reader.ReadLine();
                            string[] data = line.Split(',');

                            for (int col = 0; col < data.Length; col++)
                            {
                                worksheet.Cells[row, col + 1].Value = data[col];
                            }

                            row++;
                        }

                        // Save the Excel file
                        FileInfo excelFile = new FileInfo(excelFilePath);
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

