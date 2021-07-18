using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi;
using Microsoft.TeamFoundation.WorkItemTracking.WebApi.Models;
using Microsoft.VisualStudio.Services.Common;
using ChoETL;
using System.IO;
using Newtonsoft.Json.Linq;
using System.Net.Http;
using System.Net.Http.Headers;

namespace Work_Items_Query
{


    class Program
    {

        static async Task Main(string[] args)
        {

            List<String> projects_names = new List<String>();

            Console.WriteLine("======================== Extracting Projects Data ========================");

            // function to get the projects
            var st = GetProjects();

            JObject json = JObject.Parse(st);         


            for (int i = 0; i < Int32.Parse(json["count"].ToString()); i++)
            {
                var ids = JObject.Parse(json["value"][i].ToString())["name"];
                projects_names.Add(ids.ToString());

            }

            //Console.WriteLine(JObject.Parse(json["value"][0].ToString())["name"]);
           

            String jsonString = json["value"].ToString();



            String project_json = @"<your json file file path>";

            File.WriteAllText(project_json, jsonString);


            //========================Convert Projects Json to CSV ========================================

            String project_csv = @"<your csv file path>";

            JsonToCSV(project_json, project_csv);

            // =============================== Bugs query===================================================
            //string project_id = "7125b0b8-4193-4a92-9eda-924f6038237b";

            string json_str = "";
            foreach (string project in projects_names)
            {            
                //Console.WriteLine(project);

                string project_name = project;
                
                string personalAccessToken = "<your PAT>";
                string orgName = "<your organization name>";
                Uri uri = new Uri("https://dev.azure.com/" + orgName);
                var credentials = new VssBasicCredential(string.Empty, personalAccessToken);

                var wiql = new Wiql()
                {
                    // NOTE: Even if other columns are specified, only the ID & URL are available in the WorkItemReference
                    Query = "Select [Id] " +
                              "From WorkItems " +
                              "Where [Work Item Type] = 'Bug' " +
                              "And [System.TeamProject] = '" + project_name + "' " +
                              //"And [System.State] <> 'Closed' " +
                              "Order By [State] Asc, [Changed Date] Desc",
                };


                // create instance of work item tracking http client
                using (var httpClient = new WorkItemTrackingHttpClient(uri, credentials))
                {
                    // execute the query to get the list of work items in the results
                    var result = await httpClient.QueryByWiqlAsync(wiql).ConfigureAwait(false);

                    var ids = result.WorkItems.Select(item => item.Id).ToArray();
                    /*foreach (int i in ids)
                    {
                        Console.WriteLine("=====  id =====");
                        Console.Write(i);
                    }*/

                    // build a list of the fields we want to see
                    var fields = new[] { "System.Id", "System.Title", "System.State", "System.CreatedDate", "System.ChangedDate", "Microsoft.VSTS.Common.Priority", "Microsoft.VSTS.Common.Severity" };

                    // get work items for the ids found in query

                    var workItems = await httpClient.GetWorkItemsAsync(ids, fields, result.AsOf).ConfigureAwait(false);


                    //var workItems = await this.QueryOpenBugs(project).ConfigureAwait(false);

                    Console.WriteLine("Query Results: {0} items found", workItems.Count);


                    List<String> json_array = new List<String>();
                    
                    foreach (var workItem in workItems)
                    {                      

                        string str = "";

                        str += "{" + "\"Id\": " + "\"" + workItem.Id + "\", " + "\"Title\": " + "\"" + workItem.Fields["System.Title"].ToString().Replace("\"", "") + "\", "
                            + "\"State\": " + "\"" + workItem.Fields["System.State"] + "\", " + "\"CreatedDate\": " + "\"" + workItem.Fields["System.CreatedDate"] + "\", "
                            + "\"ChangedDate\": " + "\"" + workItem.Fields["System.ChangedDate"] + "\", " + "\"Priority\": " + "\"" + workItem.Fields["Microsoft.VSTS.Common.Priority"]
                            + "\", " + "\"Severity\" : " + "\"" + workItem.Fields["Microsoft.VSTS.Common.Severity"] + "\", " + "\"Project_name\": " + "\"" + project_name + "\"" + "}";

                        //json_array.Add(str);
                        json_str += str + ",";                      
                      
                    }                  
                }
            }

            string final_str = "[" + json_str.TrimEnd(',') + "]";


            //JObject json = JObject.Parse(json_str);
            String Bugs_json_path = @"C:\Users\tp2010023\Desktop\project_data\json\Bugs.json";
            File.WriteAllText(Bugs_json_path, final_str);
            //========================Convert Json to CSV ========================================
            String Bugs_csv = @"C:\Users\tp2010023\Desktop\project_data\csv\Bugs.csv";
            JsonToCSV(Bugs_json_path, Bugs_csv);

            //Console.ReadKey();
        }

        //========= function read json file and convert it to CSV file ==================
        public static void JsonToCSV(string jsonfilePath, string csvOutputFilePath)
        {
            try
            {
                using (var r = new ChoJSONReader(jsonfilePath))
                {
                    using (var w = new ChoCSVWriter(csvOutputFilePath).WithFirstLineHeader())
                    {
                        w.Write(r);
                    }
                }
            }
            catch (Exception ex)
            {
                Console.WriteLine("Error " + ex);
            }
        }

        public static String GetProjects()
        {
            String responseBody = "";
            try
            {
                var personalaccesstoken = "<your PAT>";

                HttpClient client = new HttpClient();

                client.DefaultRequestHeaders.Accept.Add(new System.Net.Http.Headers.MediaTypeWithQualityHeaderValue("application/json"));

                client.DefaultRequestHeaders.Authorization = new AuthenticationHeaderValue("Basic",
                        Convert.ToBase64String(System.Text.ASCIIEncoding.ASCII.GetBytes(string.Format("{0}:{1}", "", personalaccesstoken))));

                Task<string> response = client.GetStringAsync("https://dev.azure.com/<your organization name>/_apis/projects");

                responseBody = response.Result;
                //Console.WriteLine(responseBody);

            }
            catch (Exception ex)
            {
                Console.WriteLine(ex.ToString());
            }
            return responseBody;
        }
    }
}

