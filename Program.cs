using System;
using System.Collections.Generic;
using System.Data;
using System.Data.SqlClient;
using System.IO;
using DocumentFormat.OpenXml.Office.Word;
using Epplus;
using OfficeOpenXml;


class Program
{
    SearchTermValueModel SearchTermValueModel = new SearchTermValueModel();
    static int Claimid = 28;

    static void Main(string[] args)
    {
        try
        {
            string inputExcelFilePath = @"C:\Users\krish.sah\Downloads\claim_template.xlsx";
            string outputExcelFilePath = @"C:\Users\krish.sah\Downloads\new.xlsx";
            ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

            DeleteExistingOutputFile(outputExcelFilePath);

            string headerSqlQuery = FindHeaderSqlQuery(inputExcelFilePath);

            List<SearchTermValueModel> searchTermValueList = RetrieveDataFromDatabase(headerSqlQuery);

            using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(inputExcelFilePath)))
            {
                foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
                {
                    ProcessWorksheet(worksheet, searchTermValueList);
                }

                excelPackage.SaveAs(new FileInfo(outputExcelFilePath));
            }

            Console.WriteLine("Excel file has been successfully generated.");
        }
        catch (Exception ex)
        {
            Console.WriteLine("An error occurred: " + ex.Message);
        }
        finally
        {
            Console.WriteLine("Press any key to exit...");
            Console.ReadKey();
        }
    }

    static string FindHeaderSqlQuery(string inputExcelFilePath)
    {
        using (ExcelPackage excelPackage = new ExcelPackage(new FileInfo(inputExcelFilePath)))
        {
            foreach (ExcelWorksheet worksheet in excelPackage.Workbook.Worksheets)
            {
                foreach (var cellValue in worksheet.Cells)
                {
                    if (cellValue.Value != null && cellValue.Value.ToString().StartsWith("#HeaderSQL="))
                    {
                        return cellValue.Value.ToString().Replace("#HeaderSQL=", "").Trim();
                    }
                }
            }
            throw new Exception("Header SQL query not found in the Excel file.");
        }
    }

    static void ProcessWorksheet(ExcelWorksheet worksheet, List<SearchTermValueModel> searchTermValueList)
    {
        foreach (var worksheetCell in worksheet.Cells)
        {
            if (worksheetCell.Value != null && worksheetCell.Value.ToString().StartsWith("{") && worksheetCell.Value.ToString().EndsWith("}"))
            {
                string placeholder = worksheetCell.Value.ToString().Trim('{', '}');
                SearchTermValueModel termValue = searchTermValueList.Find(x => x.SearchTerm.ToLower() == placeholder.ToLower());

                if (termValue != null)
                {
                    worksheetCell.Value = termValue.SearchValue;
                }
                else
                {
                    Console.WriteLine($"No data found for placeholder '{placeholder}'.");
                }
            }
        }
    }

    static List<SearchTermValueModel> RetrieveDataFromDatabase(string headerSqlQuery)
    {
        List<SearchTermValueModel> searchTermValueList = new List<SearchTermValueModel>();

        string connectionString = "Data Source=FLXNEPLPT125\\MSSQLSERVER01;Initial Catalog=excel;Integrated Security=True";

        using (SqlConnection connection = new SqlConnection(connectionString))
        {
            connection.Open();

            using (SqlCommand command = new SqlCommand(headerSqlQuery, connection))
            {
                command.Parameters.AddWithValue("@ClaimId", Claimid);

                using (SqlDataReader reader = command.ExecuteReader())

                    while (reader.Read())
                    {

                        for (int i = 0; i < reader.FieldCount; i++)
                        {
                            string columnName = reader.GetName(i).ToString();
                            string columnValue = reader[columnName].ToString();

                            searchTermValueList.Add(new SearchTermValueModel
                            {
                                SearchTerm = columnName,
                                SearchValue = columnValue
                            });
                        }


                    }
            }
        }


        return searchTermValueList;
    }


       
    
    static void DeleteExistingOutputFile(string outputExcelFilePath)
    {
        if (File.Exists(outputExcelFilePath))
        {
            File.Delete(outputExcelFilePath);
            Console.WriteLine("Existing output file deleted.");
        }
    }
}



