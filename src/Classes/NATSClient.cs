using System;
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
        /// Decompresses the string.
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

            Msg reply = c.Request("openrmf.score.read", Encoding.UTF8.GetBytes(id), 3000); // publish to get this Artifact checklist back via ID
            c.Flush();
            // save the reply and get back the checklist score
            if (reply != null) {
                score = JsonConvert.DeserializeObject<Score>(Compression.DecompressString(Encoding.UTF8.GetString(reply.Data)));
                return score;
            }
            c.Close();
            return score;
        }
    }
}