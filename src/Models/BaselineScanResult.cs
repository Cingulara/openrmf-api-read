using System;
using System.Collections.Generic;
using System.ComponentModel;

namespace openrmf_read_api.Models
{
    [Serializable]
    public class BaselineScanResult
    {
        public BaselineScanResult () {
            checklistType = ChecklistType.XML; // default in our application
        }

        public string rawChecklist { get; set;}
        public string baseTemplateType { get; set; }
        public string scanType { get; set;}
        public ChecklistType checklistType {get; set;}
    }


    public enum ChecklistType {
        [Description("Unknown")]
        Unknown = 0,
        [Description("XML")]
        XML = 10
    }
}