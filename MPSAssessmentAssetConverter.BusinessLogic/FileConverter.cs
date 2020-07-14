using ClosedXML.Excel;
using System;
using System.Collections.Generic;
using System.Data;
using System.IO;
using System.Linq;
using System.Net.Mail;
using System.Text.RegularExpressions;

namespace MPSAssessmentAssetConverter.BusinessLogic
{
    public class FileConverter
    {
        ////INPUT 
        ////Serial Number	Cart Number	School	"Deployment Date"	Quantity	"Email to Nancy for Google"	Version	Model	"Emailed to Acer:Warranty Update"	PO	 PURCHASE CODE	Comment 	Assigned To

        ////A: Serial Number 
        ////B: Cart Number    
        ////C: School 
        ////D: Deployment Date   
        ////E: Quantity 
        ////F: Email to Nancy for Google   
        ////G: Version 
        ////H: Model    
        ////I: Emailed to Acer: Warranty Update
        ////J: PO
        ////K: PURCHASE CODE    
        ////L: Comment 
        ////M: Assigned To

        //OUTPUT CSV
        //Category
        //Company
        //Location
        //Manufacturer
        //ModelName
        //SerialNumber
        //PurchaseDate
        //CartNumber
        //TouchScreen 
        //OrderNumber 
        //AssetTag 
        //Notes


        //Assumptions: 
        //First sheet is sheet with applicable data
        //Column Order is always the same 
        

        //HeaderRow
        const int colSerialNumber = 1;
        const int colCartNumber = 2;
        const int colSchool = 3;
        const int colDeploymentDate = 4;
        const int colQuantity = 5;
        const int colEmailNancy = 6;
        const int colVersion = 7;
        const int colModel = 8;
        const int colEamiledAcer = 9;
        const int colPO = 10;
        const int colPurchaseCode = 11;
        const int colComment = 12;
        const int colAssignedTo = 13;

        public void ConvertFile(string inputFileNameAndPath, string outputFolder)
        {
            const string CompanyNameConstant = "MPS";
            const string outputHeader = "Category,Company,Location,Manufacturer,Model Name,Serial Number,Purchase Date,Cart Number,Touch Screen,Order Number,Asset Tag,Notes";
            var wb = new XLWorkbook(inputFileNameAndPath);
            var ws = wb.Worksheet(1);


            
            var lastRowUsed = ws.LastRowUsed().RowNumber();

            List<OutputFileRecord> outputRecords = new List<OutputFileRecord>();

            //First row is header row
            for (int row = 2; row <= lastRowUsed; row++)
            {
                var newOutputRecord = RowToOutputFileRecord(ws.Row(row));
                if (!string.IsNullOrWhiteSpace(newOutputRecord.AssetTag))
                {
                    outputRecords.Add(newOutputRecord);
                }
                
            }

            var schoolList = outputRecords.Select(x => x.Location).Distinct();
            
            //Create output folder
            System.IO.Directory.CreateDirectory(outputFolder);
            
            foreach (var school in schoolList)
            {
                var outputFileRecords = outputRecords.Where(x=>x.Location == school);

                //Remove white space from school name for file
                var outputFileNameAndPath = $"{outputFolder}//{Regex.Replace(school, @"\s+", "")}AssetInventory.csv";
                using (var w = new StreamWriter(outputFileNameAndPath))
                {
                    //var headerLineText = "First Name,Last Name,email,Username,Location,Phone Number,Job Title,Employee Number,Company";
                    w.WriteLine(outputHeader);
                    w.Flush();

                    
                    foreach (var outputFileRecord in outputFileRecords)
                    {

                       
                        var line = $"{outputFileRecord.Category},{outputFileRecord.Company}," +
                            $"{outputFileRecord.Location},{outputFileRecord.Manufacturer}," +
                            $"{outputFileRecord.ModelName},{outputFileRecord.SerialNumber}," +
                            $"{outputFileRecord.PurchaseDate},{outputFileRecord.CartNumber}," +
                            $"{outputFileRecord.TouchScreen},{outputFileRecord.OrderNumber}," +
                            $"{outputFileRecord.AssetTag},{outputFileRecord.Notes}";

                        w.WriteLine(line);
                        w.Flush();
                    }
                }
            }




        }

        private OutputFileRecord RowToOutputFileRecord(IXLRow row)
        {
            var assetTag = row.Cell(colAssignedTo).GetString();

            var isInt = int.TryParse(assetTag, out int assetTagInt);

            //Add only if asset tag is a 6 digit number
            if (isInt  && assetTagInt > 99999 && assetTagInt < 1000000) {
                return new OutputFileRecord
                {
                    CartNumber = EscapeCSVText(row.Cell(colCartNumber).GetString()),
                    AssetTag = assetTag,
                    Category = "Chromebooks",
                    Company = "MPS",
                    Location = EscapeCSVText(row.Cell(colSchool).GetString()).Trim(),
                    Manufacturer = "Acer",
                    ModelName = EscapeCSVText(row.Cell(colModel).GetString()),
                    Notes = EscapeCSVText($"{row.Cell(colComment).GetString()} - {row.Cell(colPurchaseCode).GetString()}"),
                    OrderNumber = EscapeCSVText(row.Cell(colPO).GetString()),
                    PurchaseDate = EscapeCSVText(row.Cell(colDeploymentDate).GetString()),
                    SerialNumber = EscapeCSVText(row.Cell(colSerialNumber).GetString()),
                    TouchScreen = row.Cell(colVersion).GetString() == "Touch" ? "True" : "False"

                };
            }
           
            return new OutputFileRecord();
        }

       

        private string EscapeCSVText(string data)
        {
            if (data.Contains("\""))
            {
                data = data.Replace("\"", "\"\"");
            }

            if (data.Contains(","))
            {
                data = String.Format("\"{0}\"", data);
            }

            if (data.Contains(System.Environment.NewLine))
            {
                data = String.Format("\"{0}\"", data);
            }

            return data;
        }      

    }

    public struct OutputFileRecord
    {
        //OUTPUT CSV

        public string Category { get; set; }
        public string Company { get; set; }
        public string Location { get; set; }
        public string Manufacturer { get; set; }
        public string ModelName { get; set; }
        public string SerialNumber { get; set; }
        public string PurchaseDate { get; set; }
        public string CartNumber { get; set; }
        public string TouchScreen { get; set; }
        public string OrderNumber { get; set; }
        public string AssetTag { get; set; }
        public string Notes { get; set; }
    }
}
