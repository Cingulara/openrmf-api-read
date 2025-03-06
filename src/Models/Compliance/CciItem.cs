// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;
using System.Collections.Generic;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class CciItem
    {
        public CciItem () {
            references = new List<CciReference>();
        }
        public string cciId { get; set; }
        public string status { get; set; }
        public string publishDate { get; set; }
        public string contributor { get; set; }
        public string definition { get; set; }
        public string type { get; set; }
        public string parameter { get; set; }
        public string note { get; set; }
        public List<CciReference> references { get; set; }
    }

    public class CciReference
    {
        public CciReference()
        {}
        
        public string creator { get; set; }
        public string title { get; set; }
        public string version { get; set; }
        public string location { get; set; }
        public string index { get; set; }
        public string majorControl { get; set; }
    }
}