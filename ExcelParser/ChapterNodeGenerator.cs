using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
    public class ChapterNodeGeneratorSettings
    {
        public string DisplayName { get; set; }
        public string UrlName { get; set; }
        public string CfaType { get; set; }
        public string ExamPercentage { get; set; }
        public string Description { get; set; }
        public string Locked { get; set; }
    }

    class ChapterNodeGenerator
    {
        public static XmlElement Generate(XmlDocument xml, ChapterNodeGeneratorSettings settings)
        {
            var chapterNode = xml.CreateElement("chapter");
            chapterNode.SetAttribute("display_name", settings.DisplayName != null ? settings.DisplayName : "");
            chapterNode.SetAttribute("url_name", settings.UrlName != null ? settings.UrlName : "");
            chapterNode.SetAttribute("cfa_type", settings.CfaType != null ? settings.CfaType : "topic");
            chapterNode.SetAttribute("exam_percentage", settings.ExamPercentage != null ? settings.ExamPercentage : "");
            chapterNode.SetAttribute("description", settings.Description != null ? settings.Description : "");
            chapterNode.SetAttribute("locked", settings.Locked != null ? settings.Locked : "yes");
            return chapterNode;
        }
    }
}
