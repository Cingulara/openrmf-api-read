using System;
using System.Collections.Generic;
using System.IO;
using System.IO.Compression;
using System.Text;
using NATS.Client;
using openrmf_read_api.Models;
using Newtonsoft.Json;

namespace openrmf_read_api.Classes
{
    public static class NATSClient
    {        
        /// <summary>
        /// Get a single checklist back by passing the ID.
        /// </summary>
        /// <param name="id">The artifact ID of the checklist.</param>
        /// <returns></returns>
        public static Score GetChecklistScore(string id)
        {
            Score score = new Score();
            // Create a new connection factory to create a connection.
            ConnectionFactory cf = new ConnectionFactory();

            // Creates a live connection to the default NATS Server running locally
            IConnection c = cf.CreateConnection(Environment.GetEnvironmentVariable("natsserverurl"));

            Msg reply = c.Request("openrmf.score.read", Encoding.UTF8.GetBytes(id), 30000); // publish to get this Artifact checklist back via ID
            c.Flush();
            // save the reply and get back the checklist score
            if (reply != null) {
                score = JsonConvert.DeserializeObject<Score>(Compression.DecompressString(Encoding.UTF8.GetString(reply.Data)));
                return score;
            }
            c.Close();
            return score;
        }

        /// <summary>
        /// Return a list of Vulnerability IDs based on the control passed in. Matches CCI to VULN IDs.
        /// </summary>
        /// <param name="control">The major control to filter the Vulnerability IDs.</param>
        /// <returns></returns>
        public static List<string> GetCCIListing(string control){
            // get the result ready to receive the info and send on
            List<string> listing = new List<string>();
            // Create a new connection factory to create a connection.
            ConnectionFactory cf = new ConnectionFactory();
            // Creates a live connection to the default NATS Server running locally
            IConnection c = cf.CreateConnection(Environment.GetEnvironmentVariable("natsserverurl"));
            // send the message with data of the control as the only payload (small)
            Msg reply = c.Request("openrmf.controls.cci", Encoding.UTF8.GetBytes(control), 30000);
            // save the reply and get back the checklist to score
            if (reply != null) {
                listing = JsonConvert.DeserializeObject<List<string>>(Compression.DecompressString(Encoding.UTF8.GetString(reply.Data)));
                c.Close();
                return listing;
            }
            c.Close();
            return listing;
        }
    }
}