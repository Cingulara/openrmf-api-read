// Copyright (c) Cingulara LLC 2020 and Tutela LLC 2020. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;

namespace openrmf_read_api.Models
{
    public class Audit
    {
        public Audit () {
            auditId = Guid.NewGuid();
        }
        public Guid auditId { get; set; }
        public string program { get; set; }
        public DateTime created { get; set; }
        public string action { get; set; }
        public string userid { get; set; }
        public string username { get; set; }
        public string fullname { get; set; }
        public string email { get; set; }
        public string url { get; set; }
        public string message { get; set; }
    }
}