using System;
using System.Collections.Generic;
using System.Linq;
using System.Threading.Tasks;
using openstig_read_api.Models;
using System.IO;
using System.Text;
using System.Xml.Serialization;
using System.Xml;

namespace openstig_read_api.Classes
{
    public static class ChecklistLoader
    {        
        public static CHECKLIST LoadASDChecklist(string rawChecklist) {
            CHECKLIST myChecklist = new CHECKLIST();
            XmlSerializer serializer = new XmlSerializer(typeof(CHECKLIST));
            using (TextReader reader = new StringReader(rawChecklist))
            {
                myChecklist = (CHECKLIST)serializer.Deserialize(reader);
            }
            return myChecklist;
        }
    }
}