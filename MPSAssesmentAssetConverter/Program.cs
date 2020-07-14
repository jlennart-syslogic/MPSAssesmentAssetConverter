using MPSAssessmentAssetConverter.BusinessLogic;
using System;

namespace MPSAssesmentAssetConverter
{
    class Program
    {
        static void Main(string[] args)
        {
            var fileName = "C:\\Users\\jlennart\\Documents\\MPS_ASSET\\MPSAssetInventory.xlsx";
            var outputFolder = "C:\\Users\\jlennart\\Documents\\MPS_ASSET\\Output";
            var fileConverter = new FileConverter();
            fileConverter.ConvertFile(fileName, outputFolder);
        }
    }
}
