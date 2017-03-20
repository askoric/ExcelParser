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
        public static List<XmlElement> Convert(XmlDocument xml, Excel<TestExcelColumn, TestExcelColumnType> mockExamExcel)
        {
            List<XmlElement> chapterNodes = new List<XmlElement>();
            List<String> mockExamContainerReferences = new List<String>();
            var mockExamContainerReferencesValues = mockExamExcel.Rows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.ContainerRef).Value);
            foreach (var mockExamContainerReferenceValue in mockExamContainerReferencesValues)
            {
                string containerReferenceValue = mockExamContainerReferenceValue.Key;
                mockExamContainerReferences.Add(containerReferenceValue.Remove(containerReferenceValue.Length - 3));
            }
            mockExamContainerReferences = mockExamContainerReferences.Distinct().ToList();
            foreach (var containerReference in mockExamContainerReferences)
            {
                char index = containerReference.Last();
                var excelRows = mockExamExcel.Rows.Where(r => r.Any(c => c.Type == TestExcelColumnType.ContainerRef && c.Value.Contains(containerReference)));
                if (excelRows.Any())
                {
                    var chapterNode = xml.CreateElement("chapter");
                    chapterNode.SetAttribute("display_name", "Mock Examination " + index);
                    chapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExam" + index + "ChapterNode", CourseTypes.Topic));
                    chapterNode.SetAttribute("cfa_type", "mock_exam");
                    chapterNode.SetAttribute("cfa_short_name", "Mock Exam " + index);

                    var amRows = new List<List<IExcelColumn<TestExcelColumnType>>>();
                    var pmRows = new List<List<IExcelColumn<TestExcelColumnType>>>();
                    string amFcmNumber = "";
                    string pmFcmNumber = "";

                    foreach (var row in excelRows)
                    {
                        var fcmNumber = row.First(tn => tn.Type == TestExcelColumnType.FcmNumber).Value;
                        if (fcmNumber.Contains("_AM"))
                        {
                            amRows.Add(row);
                            amFcmNumber = fcmNumber;
                        }
                        else if (fcmNumber.Contains("_PM"))
                        {
                            pmRows.Add(row);
                            pmFcmNumber = fcmNumber;
                        }
                    }

                    var amSequentialNode = GetMockExamSequantialNode(xml, "AM", amFcmNumber, amRows, index);
                    var pmSequentialNode = GetMockExamSequantialNode(xml, "PM", pmFcmNumber, pmRows, index);

                    chapterNode.AppendChild(amSequentialNode);
                    chapterNode.AppendChild(pmSequentialNode);
                    chapterNodes.Add(chapterNode);
                }
            }
            return chapterNodes;
        }


        private static XmlNode GetMockExamSequantialNode(XmlDocument xml, string displayName, string fcmNumber, List<List<IExcelColumn<TestExcelColumnType>>> rows, char index)
        {
            var pdfAnswers = rows.First().FirstOrDefault(tn => tn.Type == TestExcelColumnType.PdfAnswers).Value;
            var pdfQuestions = rows.First().FirstOrDefault(tn => tn.Type == TestExcelColumnType.PdfQuestions).Value;
            var sequentialNode = xml.CreateElement("sequential");
            sequentialNode.SetAttribute("display_name", displayName);
            sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-{0}-sequential-{1}-{2}", index, displayName, fcmNumber), CourseTypes.Mock));
            sequentialNode.SetAttribute("taxon_id", fcmNumber);
            sequentialNode.SetAttribute("pdf_answers", pdfAnswers);
            sequentialNode.SetAttribute("pdf_questions", pdfQuestions);

            var topicNameGroup = rows.GroupBy(r=> r.First(tn => tn.Type == TestExcelColumnType.TopicName).Value);

            foreach (var topic in topicNameGroup)
            {
                string topicName = topic.Key;
                string topicTaxonId = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicTaxonId).Value;
                var verticalNode = xml.CreateElement("vertical");
                verticalNode.SetAttribute("display_name", topicName );
                verticalNode.SetAttribute("study_session_test_id", "");
                verticalNode.SetAttribute("taxon_id", topicTaxonId);
                verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-{0}-vertical-{1}-{2}", index, displayName, topicName), CourseTypes.Mock));

                sequentialNode.AppendChild(verticalNode);

                var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, topic, new ProblemBuilderNodeSettings
                {
                    DisplayName = String.Format("Mock exam {0} - {1} questions", index, displayName),
                    UrlName = CourseConverterHelper.getGuid(String.Format("mock-{0}-progress-test-{1}-{2}", index, displayName, topicName), CourseTypes.Mock),
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
