using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Management.Automation;
using System.Management.Automation.Runspaces;
using System.Collections.ObjectModel;

using Data = Google.Apis.Sheets.v4.Data;
using System.Threading;
using System.IO;
using Google.Apis.Util.Store;
using System.Xml.Linq;
using System.Xml;
using System.Text.RegularExpressions;
using Google.Apis.Auth.OAuth2;
using Google.Apis.Http;
using Google.Apis.Services;
using Google.Apis.Sheets.v4;
using Newtonsoft.Json;
using System.Net.Http;

namespace New_Relic_APIcall
{
    class Program
    {
        static void Main(string[] args)
        {



            


            RunspaceConfiguration runspaceConfiguration = RunspaceConfiguration.Create();

            Runspace runspace = RunspaceFactory.CreateRunspace(runspaceConfiguration);

            runspace.Open();

            RunspaceInvoke scriptInvoker = new RunspaceInvoke();
            Pipeline pipeline = runspace.CreatePipeline();

            Command myCommand = new Command("C:\\Users\\udhungel\\Desktop\\ps.ps1");



            pipeline.Commands.Add(myCommand);

            // pipeline.Commands.Add("Out-String");

            var results = pipeline.Invoke();
            List<dynamic> lista = new List<dynamic>();
            foreach (var psobj in results)
            {
                lista.Add(psobj);
            }
            


             runspace.Close();
            //     RunScript();
            appendEmptyColumn();
            getTagName(lista);

        }

        private static string RunScript()
        {
           
               
            
            return null;
        }

        //Gets the tag names of the 1st column in the RAVEX reports
        private static string getTagName(List<dynamic> listOfMetrics)
        {
            //Spreadsheet ID for the google spreadsheet => form the URL link
            string spreadsheetId = "11QWSoBvPXm6_4NQ9t1xomfTHJSHYtYKd0GQP_qNF6EU";

            //Range to look for the tag names
            string range = "RaveX June 2017 Incremental Perf Test Metrics!A:A";


            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });



            SpreadsheetsResource.ValuesResource.GetRequest request =
                  sheetsService.Spreadsheets.Values.Get(spreadsheetId, range); //puts an api call to get the items in the row with the provided range

            Data.ValueRange response = request.Execute();

            IList<IList<Object>> values = response.Values; //generic list to store the row name tags

            string colRange = getColumnRange();

            int rowIndex = 1;
            
                foreach (var rowTag in values)
                {
                string newrowTag = rowTag.ToString();
                foreach (var tagName in listOfMetrics)
                {
                    string newtagName = tagName.ToString();
                    if (!newrowTag.Equals(newtagName))
                    {

                    }
                }
            }


            return null;
        }

        private static string getColumnRange()
        {
            string spreadsheetId = "11QWSoBvPXm6_4NQ9t1xomfTHJSHYtYKd0GQP_qNF6EU";



            string range = "RaveX June 2017 Incremental Perf Test Metrics!A:A";

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });


            bool loop = true;
            var startColNum = 1;
            char rangeConcat = 'R';

            do
            {
                string nullColumnIndexRange = "RaveX June 2017 Incremental Perf Test Metrics!"; //Range to be concatinated with
                string rangeStr = rangeConcat.ToString();

                string finalRange = string.Concat(nullColumnIndexRange, rangeStr, 1);//Concatinates the string to produce the range string for the right column to enter the data
 

                string colData = getColumnData(finalRange);//Loops through the google docs sheet to find the 

                if (colData.Equals("null found"))
                {
                    loop = false;

                }
                rangeConcat++;
                startColNum++;

            } while (loop == true);

            return rangeConcat.ToString();
        }

        private static string getColumnData(string range)
        {
            // The ID of the spreadsheet to update.
            string spreadsheetId = "11QWSoBvPXm6_4NQ9t1xomfTHJSHYtYKd0GQP_qNF6EU";  // TODO: Update placeholder value.

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            SpreadsheetsResource.ValuesResource.GetRequest request =
                   sheetsService.Spreadsheets.Values.Get(spreadsheetId, range);

            Data.ValueRange response = request.Execute();
            IList<IList<Object>> values = response.Values;

            return nullTest(values);
           

        }

        private static string nullTest(dynamic testVar)
        {
            
            if (testVar == null)
            {
                return "null found";
            }
            else
            {
                string tagName = testVar.ToString();
                return tagName ;
            }
        }

        private static void appendEmptyColumn()
        {

            // The A1 notation of a range to search for a logical table of data.
            // Values will be appended after the last row of the table.
            //   string range = "1 - Performance Test Results";  // TODO: Update placeholder value.
            //"'1 - Performance Test Results'!A:G"
            // The ID of the spreadsheet to update.

            // string spreadsheetId = spreadsheetIDTextbox.Text; //Impliment this for the input entered in the textbox

            string spreadsheetId = "11QWSoBvPXm6_4NQ9t1xomfTHJSHYtYKd0GQP_qNF6EU";  // TODO: Update placeholder value.

            SheetsService sheetsService = new SheetsService(new BaseClientService.Initializer
            {
                HttpClientInitializer = GetCredential(),
                ApplicationName = "RaveReportAutomationUtility",
            });

            Data.Request reqBody = new Data.Request
            {
                AppendDimension = new Data.AppendDimensionRequest
                {
                    SheetId = 920811147,
                    Dimension = "COLUMNS",
                    Length = 1

                }
            };

            List<Data.Request> requests = new List<Data.Request>();

            requests.Add(reqBody);


            // TODO: Assign values to desired properties of `requestBody`:
            Data.BatchUpdateSpreadsheetRequest requestBody = new Data.BatchUpdateSpreadsheetRequest();
            requestBody.Requests = requests;

            SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);
            //SpreadsheetsResource.BatchUpdateRequest request = sheetsService.Spreadsheets.BatchUpdate(requestBody, spreadsheetId);


            // To execute asynchronously in an async method, replace `request.Execute()` as shown:
            Data.BatchUpdateSpreadsheetResponse response = request.Execute();
          



        }

        private static UserCredential GetCredential()
        {

            UserCredential credential;
            string credPath = System.Environment.GetFolderPath(
                    System.Environment.SpecialFolder.Personal);
            credPath = Path.Combine(credPath, ".credentials/sheets.googleapis.com-dotnet-quickstart.json");
            string[] Scopes = { SheetsService.Scope.Drive };

            var stream = new FileStream("client_secret.json", FileMode.Open, FileAccess.Read);

            credential = GoogleWebAuthorizationBroker.AuthorizeAsync(
                        GoogleClientSecrets.Load(stream).Secrets,

                        Scopes,
                        "user",
                        CancellationToken.None,
                        new FileDataStore(credPath, true)).Result;
            Console.WriteLine("Credential file saved to: " + credPath);
            stream.Close();
            return credential;






        }
    }
}
