// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

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
            // add the options for the server, reconnecting, and the handler events
            Options opts = ConnectionFactory.GetDefaultOptions();
            opts.MaxReconnect = -1;
            opts.ReconnectWait = 2000;
            opts.Name = "openrmf-api-read";
            opts.Url = Environment.GetEnvironmentVariable("NATSSERVERURL");
            opts.AsyncErrorEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("NATS client error. Server: {0}. Message: {1}. Subject: {2}", events.Conn.ConnectedUrl, events.Error, events.Subscription.Subject));
            };

            opts.ServerDiscoveredEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("A new server has joined the cluster: {0}", events.Conn.DiscoveredServers));
            };

            opts.ClosedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Closed: {0}", events.Conn.ConnectedUrl));
            };

            opts.ReconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Reconnected: {0}", events.Conn.ConnectedUrl));
            };

            opts.DisconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Disconnected: {0}", events.Conn.ConnectedUrl));
            };
            
            // Creates a live connection to the NATS Server with the above options
            IConnection c = cf.CreateConnection(opts);

            Msg reply = c.Request("openrmf.score.read", Encoding.UTF8.GetBytes(id), 10000); // publish to get this Artifact checklist back via ID
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
            // add the options for the server, reconnecting, and the handler events
            Options opts = ConnectionFactory.GetDefaultOptions();
            opts.MaxReconnect = -1;
            opts.ReconnectWait = 2000;
            opts.Name = "openrmf-api-read";
            opts.Url = Environment.GetEnvironmentVariable("NATSSERVERURL");
            opts.AsyncErrorEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("NATS client error. Server: {0}. Message: {1}. Subject: {2}", events.Conn.ConnectedUrl, events.Error, events.Subscription.Subject));
            };

            opts.ServerDiscoveredEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("A new server has joined the cluster: {0}", events.Conn.DiscoveredServers));
            };

            opts.ClosedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Closed: {0}", events.Conn.ConnectedUrl));
            };

            opts.ReconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Reconnected: {0}", events.Conn.ConnectedUrl));
            };

            opts.DisconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Disconnected: {0}", events.Conn.ConnectedUrl));
            };
            
            // Creates a live connection to the NATS Server with the above options
            IConnection c = cf.CreateConnection(opts);
            
            // send the message with data of the control as the only payload (small)
            Msg reply = c.Request("openrmf.compliance.cci.control", Encoding.UTF8.GetBytes(control), 10000);
            // save the reply and get back the checklist to score
            if (reply != null) {
                listing = JsonConvert.DeserializeObject<List<string>>(Compression.DecompressString(Encoding.UTF8.GetString(reply.Data)));
                c.Close();
                return listing;
            }
            c.Close();
            return listing;
        }

        /// <summary>
        /// Get a listing of all checklists and their scores back from the Score MSG client
        /// </summary>
        /// <param name="systemGroupId">The system ID for all of the checklists.</param>
        /// <returns></returns>
        public static List<Score> GetSystemScores(string systemGroupId)
        {
            List<Score> score = new List<Score>();

            // Create a new connection factory to create a connection.
            ConnectionFactory cf = new ConnectionFactory();
            // add the options for the server, reconnecting, and the handler events
            Options opts = ConnectionFactory.GetDefaultOptions();
            opts.MaxReconnect = -1;
            opts.ReconnectWait = 2000;
            opts.Name = "openrmf-api-read";
            opts.Url = Environment.GetEnvironmentVariable("NATSSERVERURL");
            opts.AsyncErrorEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("NATS client error. Server: {0}. Message: {1}. Subject: {2}", events.Conn.ConnectedUrl, events.Error, events.Subscription.Subject));
            };

            opts.ServerDiscoveredEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("A new server has joined the cluster: {0}", events.Conn.DiscoveredServers));
            };

            opts.ClosedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Closed: {0}", events.Conn.ConnectedUrl));
            };

            opts.ReconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Reconnected: {0}", events.Conn.ConnectedUrl));
            };

            opts.DisconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Disconnected: {0}", events.Conn.ConnectedUrl));
            };
            
            // Creates a live connection to the NATS Server with the above options
            IConnection c = cf.CreateConnection(opts);

            Msg reply = c.Request("openrmf.scores.system", Encoding.UTF8.GetBytes(systemGroupId), 10000); // publish to get this system's scores back
            c.Flush();
            // save the reply and get back the checklist score
            if (reply != null) {
                score = JsonConvert.DeserializeObject<List<Score>>(Compression.DecompressString(Encoding.UTF8.GetString(reply.Data)));
                return score;
            }
            c.Close();
            return score;
        }


        /// <summary>
        /// Return a list of CCI Items to use for generating compliance
        /// </summary>
        /// <returns></returns>
        public static List<CciItem> GetCCIListing(){
            // get the result ready to receive the info and send on
            List<CciItem> cciItems = new List<CciItem>();
            
            // Create a new connection factory to create a connection.
            ConnectionFactory cf = new ConnectionFactory();
            // add the options for the server, reconnecting, and the handler events
            Options opts = ConnectionFactory.GetDefaultOptions();
            opts.MaxReconnect = -1;
            opts.ReconnectWait = 1000;
            opts.Name = "openrmf-api-compliance";
            opts.Url = Environment.GetEnvironmentVariable("NATSSERVERURL");
            opts.AsyncErrorEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("NATS client error. Server: {0}. Message: {1}. Subject: {2}", events.Conn.ConnectedUrl, events.Error, events.Subscription.Subject));
            };

            opts.ServerDiscoveredEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("A new server has joined the cluster: {0}", events.Conn.DiscoveredServers));
            };

            opts.ClosedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Closed: {0}", events.Conn.ConnectedUrl));
            };

            opts.ReconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Reconnected: {0}", events.Conn.ConnectedUrl));
            };

            opts.DisconnectedEventHandler += (sender, events) =>
            {
                Console.WriteLine(string.Format("Connection Disconnected: {0}", events.Conn.ConnectedUrl));
            };
            
            // Creates a live connection to the NATS Server with the above options
            IConnection c = cf.CreateConnection(opts);
            
            // send the message with the subject, no data needed
            Msg reply = c.Request("openrmf.compliance.cci", Encoding.UTF8.GetBytes(JsonConvert.SerializeObject("")), 30000);
            // save the reply and get back the checklist to score
            if (reply != null) {
                cciItems = JsonConvert.DeserializeObject<List<CciItem>>(Compression.DecompressString(Encoding.UTF8.GetString(reply.Data)));
                c.Close();
                return cciItems;
            }
            c.Close();
            return cciItems;
        }
    }
}