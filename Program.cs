using System;
using System.IO;
using ExcelDataReader;
using System.Net.Http;
using Newtonsoft.Json;
using System.Collections.Generic;
using Newtonsoft.Json.Linq;
using System.Text;
using OfficeOpenXml;
using System.Reflection;

class Program
{
    static void Main(string[] args)
    {
        string desktopPath = Environment.GetFolderPath(Environment.SpecialFolder.Desktop);
        // Get the directory path of the directory where the executable is running from
        string executablePath = AppDomain.CurrentDomain.BaseDirectory;

        // Get the directory path of the project directory (the parent directory of the bin directory)
        string projectDirectory = Directory.GetParent(executablePath).Parent.Parent.Parent.FullName;

        string excelFileName = "jappreviews.xlsx";
        string excelFilePath = Path.Combine(projectDirectory, excelFileName);
        string outputFileName = "ComplaintsReport.xlsx";

        List<string[]> outputData = new List<string[]>();
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance);
        var encoding = Encoding.GetEncoding(1252); // Change this to the appropriate encoding if needed


        using (var stream = File.Open(excelFilePath, FileMode.Open, FileAccess.Read))
        {
            using (var reader = ExcelReaderFactory.CreateReader(stream, new ExcelReaderConfiguration() { FallbackEncoding = encoding }))
            {
                do
                {
                    while (reader.Read())
                    {
                        // Assuming the specific column is in the first column of each row
                        string msisdnColumnData =  reader.GetDouble(0).ToString();
                        string feedbackColumnData = reader.GetString(1);

                        // Call API
                        string apiResponse = CallApi(feedbackColumnData);

                        // Parse API response
                        string[] parsedData = ParseResponse(apiResponse,msisdnColumnData);

                        // Add parsed data to output
                        outputData.Add(parsedData);
                    }
                } while (reader.NextResult());
            }
        }
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial;

        // Create a new Excel package
        using (var package = new ExcelPackage())
        {
            // Add a new worksheet
            ExcelWorksheet worksheet = package.Workbook.Worksheets.Add("Review Results");
            // Define the header row
            string[] headers = { "Msisdn", "Customer Feedback", "Is Positive", "Negativity Level", "Proposed Reply", "Ticket Category", "Ticket Date" };

            // Write the headers to the first row
            for (int i = 1; i <= headers.Length; i++)
            {
                worksheet.Cells[1, i].Value = headers[i - 1];
            }

            // Populate the worksheet with your data
            for (int i = 0; i < outputData.Count; i++)
            {
                for (int j = 0; j < outputData[i].Length; j++)
                {
                    worksheet.Cells[i+2, j+1].Value = outputData[i][j];
                }
            }

            // Save the Excel package to a file
            FileInfo excelFile = new FileInfo(outputFileName);
            package.SaveAs(excelFile);
        }

        Console.WriteLine("Excel file saved successfully.");

    }

    static string CallApi(string specificColumnData)
    {
        string apiUrl = "http://localhost:3000/api/v1/prediction/1d2bbfc4-bf16-4364-a3bf-57293cf37ea1";

        using (var client = new HttpClient())
        {
            var requestBody = new
            {
                question = specificColumnData,
                // Add other required parameters from the row if needed
            };

            var jsonRequest = JsonConvert.SerializeObject(requestBody);
            var content = new StringContent(jsonRequest, System.Text.Encoding.UTF8, "application/json");

            var response = client.PostAsync(apiUrl, content).Result;

            if (response.IsSuccessStatusCode)
            {
                return response.Content.ReadAsStringAsync().Result;
            }
            else
            {
                throw new Exception("API call failed: " + response.ReasonPhrase);
            }
        }
    }

    static string[] ParseResponse(string apiResponse, string msisdn )
    {
        // Deserialize JSON string into JObject
        JObject jsonObject = JObject.Parse(apiResponse);
        // Extract required data from JObject
        string question = (string)jsonObject["question"];
        bool positive = (bool)jsonObject["json"]["positive"];
        int level = (int)jsonObject["json"]["level"];
        string reply = (string)jsonObject["json"]["reply"];
        //string briefReply = (string)jsonObject["json"]["briefReply"];
        string problemCategory = (string)jsonObject["json"]["problemCategory"];
        string opendate = DateTime.Now.ToString();

        // Construct array with extracted data
        string[] parsedData = {
            msisdn,
            question,
            positive.ToString(),
            level.ToString(),
            reply,
            problemCategory,
            opendate
        };

        return parsedData;
    }


    public class JsonResponse
    {
        public JsonData Json { get; set; }
        public string Question { get; set; }
        public string ChatId { get; set; }
        public string ChatMessageId { get; set; }
    }

    public class JsonData
    {
        public bool Positive { get; set; }
        public int Level { get; set; }
        public string Reply { get; set; }
        public string BriefReply { get; set; }
        public bool ShouldContact { get; set; }
    }
}
