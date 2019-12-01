using System;
using System.Collections.Generic;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class ChecklistSystem
    {
        public string systemGroupId { get; set; }
        public int checklistCount { get; set; }
    }

}