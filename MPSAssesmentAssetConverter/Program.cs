using Microsoft.Extensions.Configuration;
using MPSAssessmentAssetConverter.BusinessLogic;
using System;
using System.IO;

namespace MPSAssesmentAssetConverter
{
    class Program
    {
        static void Main(string[] args)
        {

            var builder = new ConfigurationBuilder()
        .SetBasePath(Directory.GetCurrentDirectory())
        .AddJsonFile("appsettings.json");

            var configuration = builder.Build();

            var root = configuration["root"] ?? @"C:\Files\mps\";
            var fileName = configuration["fileName"] ?? "MPSAssetInventory.xlsx";
            var outputFolder = configuration["outputFolder"] ?? @"C:\Files\mps\asset-converter\Output\";
            var company = configuration["company"] ?? "MPS";
            var outputHeader = configuration["outputHeader"] ?? "Category,Company,Location,Manufacturer,Model Name,Serial Number,Purchase Date,Cart Number,Touch Screen,Order Number,Asset Tag,Notes";
            var category = configuration["category"] ?? "Chromebooks";
            var manufacturer = configuration["manufacturer"] ?? "Acer";

            var fileConverter = new FileConverter();
            fileConverter.ConvertFile(string.Concat(root, fileName), outputFolder, company, outputHeader, category, manufacturer);
        }
    }
}
