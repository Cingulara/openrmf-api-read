using System;
using System.Collections.Generic;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class ChecklistSystem
    {
        public string system { get; set; }
        public int checklistCount { get; set; }
    }

}