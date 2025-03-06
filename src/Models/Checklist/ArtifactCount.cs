// Copyright (c) Cingulara LLC 2019 and Tutela LLC 2019. All rights reserved.
// Licensed under the GNU GENERAL PUBLIC LICENSE Version 3, 29 June 2007 license. See LICENSE file in the project root for full license information.

using System;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class ArtifactCount
    {
        public ArtifactCount () {
        }
        public string stigType { get; set; }

        public int count { get; set; }
    }
}