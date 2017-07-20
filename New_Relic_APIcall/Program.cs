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


            getTagName();


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
                         
        }

        private static string RunScript()
        {
           
               
            
            return null;
        }

        //Gets the tag names of the 1st column in the RAVEX reports
        private static string getTagName()
        {
            //Spreadsheet ID for the google spreadsheet => form the URL link
            string spreadsheetId = "1j1XYUFxnAuNN6Po0IyiMu5UWjAr2a1TrOQscfabtl-A";

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






            return null;
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
