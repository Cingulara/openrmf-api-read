// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.
using System;
using System.Collections.Generic;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class UploadResult
    {
        public UploadResult () {
            failedUploads = new List<string>();
        }

        public int successful { get; set; }
        public int failed { get; set; }
        public int updated { get; set; }
        public List<string> failedUploads { get; set;}
    }
}