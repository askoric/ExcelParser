using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
    class MockExamsExcelConverter
    {
        public static List<XmlElement> Convert(XmlDocument xml, Excel<MockExamExcelColumn, MockExamExcelColumnType> mockExamsExcel)
        {
            List<XmlElement> chapterNodes = new List<XmlElement>();

            //get practice and final mock exam questions
            var practiceMockReference = "practice";
            var finalMockReference = "final";
            var practiceMockRows = mockExamsExcel.Rows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.MockType && c.Value.Contains(practiceMockReference)));
            var finalMockRows = mockExamsExcel.Rows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.MockType && c.Value.Contains(finalMockReference)));

            //work on practice mocks
            List<String> practiceMockExamContainerReferences = new List<String>();
            var practiceMockExamContainerReferencesValues = practiceMockRows.GroupBy(r => r.First(tn => tn.Type == MockExamExcelColumnType.Container1Ref).Value);
            foreach (var practiceMockExamContainerReferencesValue in practiceMockExamContainerReferencesValues)
            {
                string containerReferenceValue = practiceMockExamContainerReferencesValue.Key;
                //remove _AM and _PM from keys
                practiceMockExamContainerReferences.Add(containerReferenceValue.Remove(containerReferenceValue.Length - 3));
            }
            practiceMockExamContainerReferences = practiceMockExamContainerReferences.Distinct().ToList();
            foreach (var containerReference in practiceMockExamContainerReferences)
            {
                var mockRows = practiceMockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.Container1Ref && c.Value.Contains(containerReference)));
                if(mockRows.Any())
                {
                    char index = containerReference.Last();
                    var mockExamChapterFcmNumber = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.FcmNumber).Value;
                    mockExamChapterFcmNumber = mockExamChapterFcmNumber.Remove(mockExamChapterFcmNumber.Length - 3);
                    var mockExamChapterNode = GetMockExamChapterNode(xml, mockRows);
                    mockExamChapterNode.SetAttribute("display_name", "Mock Examination " + index);
                    mockExamChapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExamChapterNode " + mockExamChapterFcmNumber, CourseTypes.Topic));
                    mockExamChapterNode.SetAttribute("cfa_type", "mock_exam");
                    mockExamChapterNode = GetMockExamType(mockExamChapterNode);
                    mockExamChapterNode.SetAttribute("cfa_short_name", "Mock Exam " + index);
                    mockExamChapterNode.SetAttribute("test_duration", "03:00");
                    chapterNodes.Add(mockExamChapterNode);
                }
            }


            //work on final mock exam
            var finalExamChapterNode = GetMockExamChapterNode(xml, finalMockRows);
            var finalExamChapterFcmNumber = finalMockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.FcmNumber).Value;
            finalExamChapterFcmNumber = finalExamChapterFcmNumber.Remove(finalExamChapterFcmNumber.Length - 3);
            finalExamChapterNode.SetAttribute("display_name", "Final Mock Examination");
            finalExamChapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExamChapterNode " + finalExamChapterFcmNumber, CourseTypes.Topic));
            finalExamChapterNode.SetAttribute("cfa_type", "final_mock_exam");
            finalExamChapterNode = GetMockExamType(finalExamChapterNode);
            finalExamChapterNode.SetAttribute("cfa_short_name", "Final Mock Exam");
            finalExamChapterNode.SetAttribute("test_duration", "03:00");
            chapterNodes.Add(finalExamChapterNode);

            return chapterNodes;
        }

        private static XmlElement GetMockExamType(XmlElement chapterNode)
        {
            //check if mock exam is item_set or regular
            bool ifItemSet = true;
            foreach (XmlElement sequentialNode in chapterNode.ChildNodes)
            {
                foreach (XmlElement verticalNode in sequentialNode.ChildNodes)
                {
                    if (verticalNode.GetAttributeNode("vignette_title").Value == "" && verticalNode.GetAttributeNode("vignette_body").Value == "")
                    {
                        ifItemSet = false;
                    }
                }
            }
            chapterNode.SetAttribute("exam_type", ifItemSet ? "item_set" : "regular");

            return chapterNode;
        }

        private static XmlElement GetMockExamChapterNode(XmlDocument xml, IEnumerable<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {

            //create chapter node
            var chapterNode = xml.CreateElement("chapter");

            //divide between AM and PM
            var amRows = new List<List<IExcelColumn<MockExamExcelColumnType>>>();
            var pmRows = new List<List<IExcelColumn<MockExamExcelColumnType>>>();
            string amFcmNumber = "";
            string pmFcmNumber = "";
            foreach (var row in mockRows)
            {
                var fcmNumber = row.FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.FcmNumber).Value;
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

            //get sequential nodes for AM and PM sections
            var amSequentialNode = GetMockExamSequantialNode(xml, "AM", amFcmNumber, amRows);
            var pmSequentialNode = GetMockExamSequantialNode(xml, "PM", pmFcmNumber, pmRows);
            chapterNode.AppendChild(amSequentialNode);
            chapterNode.AppendChild(pmSequentialNode);

            return chapterNode;
        }

        private static XmlNode GetMockExamSequantialNode(XmlDocument xml, string displayName, string fcmNumber, List<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {

            //create sequential node
            var pdfAnswers = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.PdfAnswers).Value;
            var pdfQuestions = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.PdfQuestions).Value;
            var sequentialNode = xml.CreateElement("sequential");
            sequentialNode.SetAttribute("display_name", displayName);
            sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-sequential-" + fcmNumber), CourseTypes.Mock));
            sequentialNode.SetAttribute("taxon_id", fcmNumber);
            sequentialNode.SetAttribute("pdf_answers", pdfAnswers);
            sequentialNode.SetAttribute("pdf_questions", pdfQuestions);


            List<String> verticalContainerReferences = new List<String>();
            var verticalContainerReferencesValues = mockRows.GroupBy(r => r.First(tn => tn.Type == MockExamExcelColumnType.TopicAbbrevation).Value);
            foreach (var verticalContainerReferencesValue in verticalContainerReferencesValues)
            {
                verticalContainerReferences.Add(verticalContainerReferencesValue.Key);
            }
            foreach (var containerReference in verticalContainerReferences)
            {
                var topicRows = mockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.TopicAbbrevation && c.Value.Contains(containerReference)));

                //divide topic in item sets
                List<String> container2References = new List<String>();
                var container2ReferencesValues = topicRows.GroupBy(r => r.First(tn => tn.Type == MockExamExcelColumnType.Container2Ref).Value);
                foreach (var container2ReferencesValue in container2ReferencesValues)
                {
                    container2References.Add(container2ReferencesValue.Key);
                }

                foreach (var container2Reference in container2References)
                {
                    //work on vertical
                    var container2Rows = topicRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.Container2Ref && c.Value.Contains(container2Reference)));
                    var verticalNode = GetMockExamVerticalNode(xml, displayName, fcmNumber, container2Rows);
                    sequentialNode.AppendChild(verticalNode);
                }
            }

            return sequentialNode;
        }

        private static XmlNode GetMockExamVerticalNode(XmlDocument xml, string displayName, string fcmNumber, IEnumerable<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {
            string topicName = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.TopicName).Value;
            string topicTaxonId = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.TopicTaxonId).Value;
            string container2Title = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2Title) != null ? 
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2Title).Value : "";
            //if item set title empty leave old vertical display name, if not change it
            topicName = (container2Title == "") ? topicName : container2Title;
            string vignetteTitle = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteTitle) != null ? 
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteTitle).Value : "";
            string vignetteBody = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteBody) != null ? 
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteBody).Value : "";

            //create vertical node
            var verticalNode = xml.CreateElement("vertical");
            verticalNode.SetAttribute("display_name", topicName);
            verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-vertical-" + fcmNumber + "-" + topicName), CourseTypes.Mock));
            verticalNode.SetAttribute("study_session_test_id", "");
            verticalNode.SetAttribute("taxon_id", topicTaxonId);
            verticalNode.SetAttribute("vignette_title", vignetteTitle);
            verticalNode.SetAttribute("vignette_body", vignetteBody);

            //get which mock exam it is
            var letter = fcmNumber.Remove(fcmNumber.Length - 3).Last();

            //skip vignette row, if there is any
            var topicQuestions = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Question).HaveValue() &&
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Question) != null ? mockRows : mockRows.Skip(1);

            var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, topicQuestions, new ProblemBuilderNodeSettings
            {
                DisplayName = String.Format("Mock exam {0} - {1} questions", letter, displayName),
                UrlName = CourseConverterHelper.getGuid(String.Format("mock-progress-test-" + fcmNumber + "-" + topicName), CourseTypes.Mock),
                ProblemBuilderNodeElement = "problem-builder-mock-exam",
                PbMcqNodeElement = "pb-mcq-mock-exam",
                PbChoiceBlockElement = "pb-choice-mock-exam",
                PbTipBlockElement = "pb-tip-mock-exam"
            });

            verticalNode.AppendChild(problemBuilderNode);

            return verticalNode;
        }
    }
}
