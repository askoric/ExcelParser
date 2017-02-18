using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
    class ProgressTestExcelConverter
    {
        public static XmlElement Convert(XmlDocument xml, Excel<TestExcelColumn, TestExcelColumnType> progressTestExcel)
        {
            var chapterNode = xml.CreateElement("chapter");
            chapterNode.SetAttribute("display_name", "Progress Test");
            chapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("ProgressTestChapeterNode", CourseTypes.Topic));
            chapterNode.SetAttribute("cfa_type", "progress_test");
            chapterNode.SetAttribute("cfa_short_name", "Progress Test");

            var topicGroup = new List<List<IExcelColumn<TestExcelColumnType>>>();
            string previousTopicName = null;
            string previousTopicId = null;
            string kStructure = null;
            XmlNode sequentialNode = null;
            var lastRow = progressTestExcel.Rows.Last();

            foreach (var row in progressTestExcel.Rows)
            {

                var topicName = row.First(tn => tn.Type == TestExcelColumnType.TopicName).Value;
                var topicAbbrevation = row.First(c => c.Type == TestExcelColumnType.TopicAbbrevation).Value;

                string topicId = String.Format("{0}-r-progressTest", topicAbbrevation);

                if (previousTopicName != null && previousTopicName != topicName)
                {

                    sequentialNode = GetProgressTestSequantialNode(xml, previousTopicName, previousTopicId, kStructure, topicGroup);
                    chapterNode.AppendChild(sequentialNode);
                    topicGroup = new List<List<IExcelColumn<TestExcelColumnType>>>();
                }

                topicGroup.Add(row);
                previousTopicName = topicName;
                previousTopicId = topicId;
                kStructure = String.Join("|", row.First(c => c.Type == TestExcelColumnType.KStructure).Value.Split('|').Take(2));
            }

            //Append last question group
            sequentialNode = GetProgressTestSequantialNode(xml, previousTopicName, previousTopicId, kStructure, topicGroup);
            chapterNode.AppendChild(sequentialNode);

            return chapterNode;
        }


        private static XmlNode GetProgressTestSequantialNode(XmlDocument xml, string topicName, string topicId, string kStructure, List<List<IExcelColumn<TestExcelColumnType>>> topicGroup)
        {
            var sequentialNode = xml.CreateElement("sequential");
            sequentialNode.SetAttribute("display_name", topicName);
            sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(topicId, CourseTypes.StudySession));
            sequentialNode.SetAttribute("taxon_id", kStructure);

            var verticalNode = xml.CreateElement("vertical");
            verticalNode.SetAttribute("display_name", "Progress test - R");
            verticalNode.SetAttribute("study_session_test_id", "");
            verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(topicId, CourseTypes.Reading));

            sequentialNode.AppendChild(verticalNode);

            var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, topicGroup, new ProblemBuilderNodeSettings
            {
                DisplayName = "Progress test",
                UrlName = CourseConverterHelper.getGuid(topicId, CourseTypes.Question),
                ProblemBuilderNodeElement = "problem-builder-progress-test",
                PbMcqNodeElement = "pb-mcq-progress-test",
                PbChoiceBlockElement = "pb-choice-progress-test",
                PbTipBlockElement = "pb-tip-progress-test"
            });
            verticalNode.AppendChild(problemBuilderNode);

            return sequentialNode;
        }
    }
}
