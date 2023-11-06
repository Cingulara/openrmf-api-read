// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System.Collections.Generic;
using openrmf_read_api.Models;
using System.Xml.Serialization;
using System.Xml;

namespace openrmf_read_api.Classes
{
    public static class RecordGenerator
    {        
        // Checklist common routines
        public static string DecodeHTML (string html) {
            if (!string.IsNullOrEmpty(html))
                return System.Web.HttpUtility.HtmlDecode(html);
            else 
                return "";
        }
    }
}