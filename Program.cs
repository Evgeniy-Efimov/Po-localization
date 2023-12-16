using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using LocalizePo.Entities;
using LocalizePo.Providers;

namespace LocalizePo
{
    internal class Program
    {
        static void Main(string[] args)
        {
            try
            {
                //parameters
                Encoding encoding = new UTF8Encoding(true);

                //input variables
                var workDirectory = "Data";
                var localizationFileName = "Game.po";
                var workFileName = "Game.xlsx";
                var localizationFilePath = string.Empty;
                var localizationFileText = string.Empty;
                var workFilePath = string.Empty;

                //providers
                IExcelProvider excelProvider = new ExcelProvider();
                IPoProvider poProvider = new PoProvider();

                Console.WriteLine($"1 - Convert data file to excel");
                Console.WriteLine($"2 - Save data from excel file");
                Console.Write($"Choose operation by typing number: ");
                var input = Console.ReadLine();

                switch (input)
                {
                    case "1":

                        //get data file name and path
                        Console.Write($"Type work directory path or press Enter for default ({workDirectory}): ");
                        input = Console.ReadLine();
                        workDirectory = string.IsNullOrWhiteSpace(input) ? workDirectory : input;

                        Console.Write($"Type file name or press Enter for default ({localizationFileName}): ");
                        input = Console.ReadLine();
                        localizationFileName = string.IsNullOrWhiteSpace(input) ? localizationFileName : input;
                        localizationFilePath = Path.Combine(workDirectory, localizationFileName);

                        //check and read file
                        if (!File.Exists(localizationFilePath))
                        {
                            throw new Exception($"File {localizationFilePath} not found");
                        }

                        localizationFileText = File.ReadAllText(localizationFilePath);

                        if (string.IsNullOrWhiteSpace(localizationFileText))
                        {
                            throw new Exception($"File is empty");
                        }

                        var localizationData = poProvider.ReadFile<LocalizationData>(localizationFilePath).ToList();

                        //get result file name + path
                        Console.WriteLine($"Read {localizationData.Count} rows");
                        Console.Write($"Type work file name or press Enter for default ({workFileName}): ");
                        input = Console.ReadLine();

                        workFileName = string.IsNullOrWhiteSpace(input) ? workFileName : input;
                        workFilePath = Path.Combine(workDirectory, workFileName);
                        Console.WriteLine("Processing...");

                        //using excel provider form file and save it
                        using (var memoryStream = excelProvider.ExportToFile(localizationData.Select(ld => ld.Model), "Localization data"))
                        {
                            //go back to the beginning of the file before writing
                            var memoryStreamCopy = new MemoryStream(memoryStream.ToArray());
                            memoryStreamCopy.Seek(0, SeekOrigin.Begin);

                            if (File.Exists(workFilePath)) { File.Delete(workFilePath); }

                            using (var fileStream = new FileStream(workFilePath, FileMode.OpenOrCreate))
                            {
                                memoryStreamCopy.CopyTo(fileStream);
                                fileStream.Flush();
                            }
                        }

                        Console.WriteLine($"Result has been saved to file {workFilePath}");

                        break;

                    case "2":

                        //get data file name and path
                        Console.Write($"Type work directory path or press Enter for default ({workDirectory}): ");
                        input = Console.ReadLine();
                        workDirectory = string.IsNullOrWhiteSpace(input) ? workDirectory : input;

                        Console.Write($"Type localization file name or press Enter for default ({localizationFileName}): ");
                        input = Console.ReadLine();
                        localizationFileName = string.IsNullOrWhiteSpace(input) ? localizationFileName : input;

                        localizationFilePath = Path.Combine(workDirectory, localizationFileName);

                        //check and read file
                        if (!File.Exists(localizationFilePath))
                        {
                            throw new Exception($"File {localizationFilePath} not found");
                        }

                        //split all text by new localization line substring
                        var originalLocalizationData = poProvider.ReadFile<LocalizationData>(localizationFilePath).ToList();

                        Console.Write($"Type work file name or press Enter for default ({workFileName}): ");
                        input = Console.ReadLine();
                        workFileName = string.IsNullOrWhiteSpace(input) ? workFileName : input;

                        workFilePath = Path.Combine(workDirectory, workFileName);

                        if (!File.Exists(workFilePath))
                        {
                            throw new Exception($"File {workFilePath} not found");
                        }

                        Console.WriteLine("Processing...");

                        var workData = new List<RowImportResult<LocalizationData>>();

                        using (var fileStream = File.OpenRead(workFilePath))
                        {
                            workData = excelProvider.ReadFile<LocalizationData>(fileStream).ToList();
                        }

                        var updateResults = new List<UpdateResult>();

                        //find and replace original file localization data
                        foreach (var data in workData.Where(wd => !string.IsNullOrEmpty(wd.Model.Key)))
                        {
                            try
                            {
                                if (!data.IsSuccessfullyProcessed)
                                {
                                    throw new Exception(data.Message);
                                }

                                var localization = originalLocalizationData.FirstOrDefault(old => old.Model?.Key == data.Model.Key && old.Model?.SourceLocation == data.Model.SourceLocation);

                                if (localization == null)
                                {
                                    throw new Exception($"Localization with key {data.Model.Key} and source {data.Model.SourceLocation} not found");
                                }

                                if (localization.Model.LocalizedText != data.Model.LocalizedText)
                                {
                                    localization.SetModelPropertyText(typeof(LocalizationData).GetProperty(nameof(data.Model.LocalizedText)), data.Model.LocalizedText);

                                    updateResults.Add(new UpdateResult()
                                    {
                                        Key = data.Model.Key,
                                        OriginalText = data.Model.OriginalText,
                                        LocalizedText = data.Model.LocalizedText,
                                        IsSuccessfullyProcessed = true,
                                        Message = "Localization has been updated"
                                    });
                                }
                                else
                                {
                                    updateResults.Add(new UpdateResult()
                                    {
                                        Key = data.Model.Key,
                                        OriginalText = data.Model.OriginalText,
                                        LocalizedText = data.Model.LocalizedText,
                                        IsSuccessfullyProcessed = true
                                    });
                                }
                            }
                            catch (Exception updateEx)
                            {
                                updateResults.Add(new UpdateResult()
                                {
                                    Key = data.Model.Key,
                                    OriginalText = data.Model.OriginalText,
                                    LocalizedText = data.Model.LocalizedText,
                                    IsSuccessfullyProcessed = false,
                                    Message = updateEx.Message
                                });
                            }
                        }

                        //save results and logs
                        var updateWorkFileName = $"{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")} {localizationFileName}";
                        var updateWorkFilePath = Path.Combine(workDirectory, updateWorkFileName);
                        poProvider.ExportToFile(originalLocalizationData, updateWorkFilePath, encoding);

                        var updateResultsFileName = $"{DateTime.Now.ToString("yyyy-MM-dd_HH-mm-ss")} Update results.xlsx";

                        using (var memoryStream = excelProvider.ExportToFile(updateResults, "Update results"))
                        {
                            //go back to the beginning of the file before writing
                            var memoryStreamCopy = new MemoryStream(memoryStream.ToArray());
                            memoryStreamCopy.Seek(0, SeekOrigin.Begin);

                            using (var fileStream = new FileStream(Path.Combine(workDirectory, updateResultsFileName), FileMode.OpenOrCreate))
                            {
                                memoryStreamCopy.CopyTo(fileStream);
                                fileStream.Flush();
                            }
                        }

                        Console.WriteLine($"Result has been saved to file {updateWorkFilePath}, logs has been saved to file {updateResultsFileName}.");

                        break;
                    default:
                        throw new Exception("unknown operation");
                }
            }
            catch(Exception ex) 
            {
                Console.WriteLine($"Error: {ex.Message}");
            }

            Console.WriteLine("Press Enter to exit");
            Console.Read();
        }
    }
}
