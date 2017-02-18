using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
    class MockExamExcelConverter
    {
        public static XmlElement Convert(XmlDocument xml, Excel<TestExcelColumn, TestExcelColumnType> mockExamExcel)
        {
            var chapterNode = xml.CreateElement("chapter");
            chapterNode.SetAttribute("display_name", "Mock Exam");
            chapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExamChapterNode", CourseTypes.Topic));
            chapterNode.SetAttribute("cfa_type", "mock_exam");
            chapterNode.SetAttribute("cfa_short_name", "Mock Exam");

            var amRows = new List<List<IExcelColumn<TestExcelColumnType>>>();
            var pmRows = new List<List<IExcelColumn<TestExcelColumnType>>>();
            string amFcmNumber = "";
            string pmFcmNumber = "";

            foreach (var row in mockExamExcel.Rows)
            {
                var fcmNumber = row.First(tn => tn.Type == TestExcelColumnType.FcmNumber).Value;
                if (fcmNumber.Contains("_AM"))
                {
                    amRows.Add(row);
                    amFcmNumber = fcmNumber;
                }
                else if (fcmNumber.Contains("_PM")) {
                    pmRows.Add(row);
                    pmFcmNumber = fcmNumber;
                }
            }

            var amSequentialNode = GetMockExamSequantialNode(xml, "AM", amFcmNumber, amRows);
            var pmSequentialNode = GetMockExamSequantialNode(xml, "PM", pmFcmNumber, pmRows);

            chapterNode.AppendChild(amSequentialNode);
            chapterNode.AppendChild(pmSequentialNode);


            return chapterNode;
        }


        private static XmlNode GetMockExamSequantialNode(XmlDocument xml, string displayName, string fcmNumber, List<List<IExcelColumn<TestExcelColumnType>>> rows)
        {
            var sequentialNode = xml.CreateElement("sequential");
            sequentialNode.SetAttribute("display_name", displayName);
            sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid("sequential-" + displayName, CourseTypes.Mock));
            sequentialNode.SetAttribute("taxon_id", fcmNumber);

            var topicNameGroup = rows.GroupBy(r=> r.First(tn => tn.Type == TestExcelColumnType.TopicName).Value);

            foreach (var topic in topicNameGroup)
            {
                string topicName = topic.Key;
                var verticalNode = xml.CreateElement("vertical");
                verticalNode.SetAttribute("display_name", topicName );
                verticalNode.SetAttribute("study_session_test_id", "");
                verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid("vertical-" + displayName, CourseTypes.Mock));

                sequentialNode.AppendChild(verticalNode);

                var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, topic, new ProblemBuilderNodeSettings
                {
                    DisplayName = String.Format("Mock exam - {0} questions", displayName),
                    UrlName = CourseConverterHelper.getGuid("progress-test-" + displayName, CourseTypes.Mock),
                    ProblemBuilderNodeElement = "problem-builder-mock-exam",
                    PbMcqNodeElement = "pb-mcq-mock-exam",
                    PbChoiceBlockElement = "pb-choice-mock-exam",
                    PbTipBlockElement = "pb-tip-mock-exam"
                });

                verticalNode.AppendChild(problemBuilderNode);
            }

            return sequentialNode;
        }
    }
}
