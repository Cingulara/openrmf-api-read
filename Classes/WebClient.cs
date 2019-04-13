using System;
using System.Threading.Tasks;
using System.Net.Http;
using System.Net.Http.Headers;
using openstig_read_api.Models;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Xml.Serialization;

namespace openstig_read_api.Classes
{
    public static class WebClient 
    {
        public static async Task<Score> GetChecklistScore(string artifactId)
        {
            // Create a New HttpClient object and dispose it when done, so the app doesn't leak resources
            using (HttpClient client = new HttpClient())
            {
                // Call asynchronous network methods in a try/catch block to handle exceptions
                try	
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Add("Accept", "application/json");
                    string hosturl = Environment.GetEnvironmentVariable("openstig-api-score-server");
                    HttpResponseMessage response = await client.GetAsync(hosturl + "/artifact/" + System.Uri.EscapeUriString(artifactId));
                    response.EnsureSuccessStatusCode();
                    string responseBody = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<Score>(responseBody);
                    return result;
                }
                catch(HttpRequestException e)
                {
                    // log something here
                    return new Score();
                }
                catch (Exception ex) {
                    // log something here
                    return new Score();
                }
            }
        }

        
        public static async Task<List<string>> GetCCIListing(string control)
        {
            // Create a New HttpClient object and dispose it when done, so the app doesn't leak resources
            using (HttpClient client = new HttpClient())
            {
                // Call asynchronous network methods in a try/catch block to handle exceptions
                try	
                {
                    client.DefaultRequestHeaders.Accept.Clear();
                    client.DefaultRequestHeaders.Add("Accept", "application/json");
                    string hosturl = Environment.GetEnvironmentVariable("openstig-api-compliance-server");
                    HttpResponseMessage response = await client.GetAsync(hosturl + "/cci/" + System.Uri.EscapeUriString(control));
                    response.EnsureSuccessStatusCode();
                    string responseBody = await response.Content.ReadAsStringAsync();
                    var result = JsonConvert.DeserializeObject<List<string>>(responseBody);
                    return result;
                }
                catch(HttpRequestException e)
                {
                    // log something here
                    throw e;
                }
                catch (Exception ex) {
                    // log something here
                    throw ex;
                }
            }
        }
    }
}