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
            var practiceMockReference = "Practice";
            var finalMockReference = "Final";
            var practiceMockRows = mockExamsExcel.Rows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.MockType && c.Value.Contains(practiceMockReference)));
            var finalMockRows = mockExamsExcel.Rows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.MockType && c.Value.Contains(finalMockReference)));

            //work on practice mocks
            List<String> practiceMockExamPositionReferences = new List<String>();
            var practiceMockExamPositionReferencesValues = practiceMockRows.GroupBy(r => r.First(tn => tn.Type == MockExamExcelColumnType.PositionRef).Value);
            foreach (var practiceMockExamPositionReferencesValue in practiceMockExamPositionReferencesValues)
            {
                string containerReferenceValue = practiceMockExamPositionReferencesValue.Key;
                //remove _AM and _PM from keys
                practiceMockExamPositionReferences.Add(containerReferenceValue.Remove(containerReferenceValue.Length - 3));
            }
            practiceMockExamPositionReferences = practiceMockExamPositionReferences.Distinct().ToList();
            foreach (var positionReference in practiceMockExamPositionReferences)
            {
                var mockRows = practiceMockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.PositionRef && c.Value.Contains(positionReference)));
                if(mockRows.Any())
                {
                    char index = positionReference.Last();
                    var mockExamChapterContainerRef = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.Container1Ref).Value;
                    mockExamChapterContainerRef = mockExamChapterContainerRef.Remove(mockExamChapterContainerRef.Length - 3);
                    var mockExamChapterNode = GetMockExamChapterNode(xml, mockRows);
                    mockExamChapterNode.SetAttribute("display_name", "Mock Examination " + index);
                    mockExamChapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExamChapterNode " + mockExamChapterContainerRef, CourseTypes.Topic));
                    mockExamChapterNode.SetAttribute("cfa_type", "mock_exam");
                    mockExamChapterNode = GetMockExamType(mockExamChapterNode);
                    mockExamChapterNode.SetAttribute("cfa_short_name", "Mock Exam " + index);
                    mockExamChapterNode.SetAttribute("test_duration", "03:00");
                    chapterNodes.Add(mockExamChapterNode);
                }
            }


            //work on final mock exam
            var finalExamChapterNode = GetMockExamChapterNode(xml, finalMockRows);
            var finalExamChapterContainerRef = finalMockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.Container1Ref).Value;
            finalExamChapterContainerRef = finalExamChapterContainerRef.Remove(finalExamChapterContainerRef.Length - 3);
            finalExamChapterNode.SetAttribute("display_name", "Final Mock Examination");
            finalExamChapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExamChapterNode " + finalExamChapterContainerRef, CourseTypes.Topic));
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

            bool ifEssay = false;
            foreach (XmlElement sequentialNode in chapterNode.ChildNodes)
            {
                if (sequentialNode.GetAttributeNode("cfa_type").Value == "essay")
                {
                    ifEssay = true;
                }
            }

            if (ifEssay)
            {
                chapterNode.SetAttribute("exam_type", "essay");
            }

            return chapterNode;
        }

        private static XmlElement GetMockExamChapterNode(XmlDocument xml, IEnumerable<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {

            //create chapter node
            var chapterNode = xml.CreateElement("chapter");

            //divide between AM and PM
            var amRows = new List<List<IExcelColumn<MockExamExcelColumnType>>>();
            var pmRows = new List<List<IExcelColumn<MockExamExcelColumnType>>>();
            string amContainerRef = "";
            string pmContainerRef = "";
            foreach (var row in mockRows)
            {
                var seqContainerRef = row.FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.Container1Ref).Value;
                if (seqContainerRef.Contains("_AM"))
                {
                    amRows.Add(row);
                    amContainerRef = seqContainerRef;
                }
                else if (seqContainerRef.Contains("_PM"))
                {
                    pmRows.Add(row);
                    pmContainerRef = seqContainerRef;
                }
            }

            //get sequential nodes for AM and PM sections
            var amSequentialNode = GetMockExamSequantialNode(xml, "AM", amContainerRef, amRows);
            var pmSequentialNode = GetMockExamSequantialNode(xml, "PM", pmContainerRef, pmRows);
            chapterNode.AppendChild(amSequentialNode);
            chapterNode.AppendChild(pmSequentialNode);

            return chapterNode;
        }

        private static XmlNode GetMockExamSequantialNode(XmlDocument xml, string displayName, string seqContainerRef, List<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {

            //create sequential node
            var pdfAnswers = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.PdfAnswers).Value;
            var pdfQuestions = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.PdfQuestions).Value;
            var sequentialNode = xml.CreateElement("sequential");
            sequentialNode.SetAttribute("display_name", displayName);
            sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-sequential-{0}", seqContainerRef), CourseTypes.Mock));
            sequentialNode.SetAttribute("taxon_id", seqContainerRef);
            sequentialNode.SetAttribute("pdf_answers", pdfAnswers);
            sequentialNode.SetAttribute("pdf_questions", pdfQuestions);
            sequentialNode.SetAttribute("cfa_type", "");

            //get essays and questions
            var essayRows = mockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.Container2Type && c.Value.Contains("Essay")));
            var questionRows = mockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.Container2Type && c.Value.Contains("Item Set")));

            if (!questionRows.Any() && !essayRows.Any())
            {
                questionRows = mockRows;
            }

            //work on questions
            if (questionRows.Any())
            {
                //divide by item sets or topics
                List<String> container2References = new List<String>();
                List<String> verticalContainerReferences = new List<String>();
                //check if topic needs to be divided to item sets
                var container2RefKey = mockRows.First().FirstOrDefault(tn => tn.Type == MockExamExcelColumnType.Container2Ref).Value;
                if (container2RefKey != null)
                {
                    var container2ReferencesValues = mockRows.GroupBy(r => r.First(tn => tn.Type == MockExamExcelColumnType.Container2Ref).Value);
                    foreach (var container2ReferencesValue in container2ReferencesValues)
                    {
                        container2References.Add(container2ReferencesValue.Key);
                    }

                    foreach (var container2Reference in container2References)
                    {
                        //work on vertical
                        var container2Rows = mockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.Container2Ref && c.Value.Contains(container2Reference)));
                        var verticalNode = GetMockExamVerticalNode(xml, seqContainerRef, container2Rows);
                        sequentialNode.AppendChild(verticalNode);
                    }
                }
                else
                {
                    var verticalContainerReferencesValues = mockRows.GroupBy(r => r.First(tn => tn.Type == MockExamExcelColumnType.TopicRef).Value);
                    foreach (var verticalContainerReferencesValue in verticalContainerReferencesValues)
                    {
                        verticalContainerReferences.Add(verticalContainerReferencesValue.Key);
                    }
                    foreach (var containerReference in verticalContainerReferences)
                    {
                        //work on vertical
                        var topicRows = mockRows.Where(r => r.Any(c => c.Type == MockExamExcelColumnType.TopicRef && c.Value.Contains(containerReference)));
                        var verticalNode = GetMockExamVerticalNode(xml, seqContainerRef, topicRows);
                        sequentialNode.AppendChild(verticalNode);
                    }
                }
            }

            //work on essays
            if (essayRows.Any())
            {
                sequentialNode.SetAttribute("cfa_type", "essay");

                foreach (var row in essayRows)
                {
                    string topicTaxonId = row.FirstOrDefault(c => c.Type == MockExamExcelColumnType.TopicTaxonId).Value;
                    string container2Title = row.FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2Title).Value != null ?
                        row.FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2Title).Value : "";
                    string essayMaxPoints = row.FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2MaxPoints).Value;
                    string essayTopics = row.FirstOrDefault(c => c.Type == MockExamExcelColumnType.TopicRef).Value;

                    var verticalNode = xml.CreateElement("vertical");
                    verticalNode.SetAttribute("cfa_type", "essay");
                    verticalNode.SetAttribute("taxon_id", topicTaxonId);
                    verticalNode.SetAttribute("display_name", container2Title);
                    verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-vertical-{0}-{1}", seqContainerRef, container2Title), CourseTypes.Mock));
                    verticalNode.SetAttribute("essay_max_points", essayMaxPoints);
                    verticalNode.SetAttribute("study_session_test_id", "");
                    verticalNode.SetAttribute("vignette_title", "");
                    verticalNode.SetAttribute("vignette_body", "");
                    verticalNode.SetAttribute("item_set_sessions", essayTopics);

                    sequentialNode.AppendChild(verticalNode);

                }
            }

            return sequentialNode;
        }

        private static XmlNode GetMockExamVerticalNode(XmlDocument xml, string seqContainerRef, IEnumerable<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {
            string topicName = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.TopicName).Value;
            string topicTaxonId = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.TopicTaxonId).Value;
            string container2Title = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2Title).Value != null ? 
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Container2Title).Value : "";
            //if item set title empty leave old vertical display name, if not change it
            topicName = (container2Title == "") ? topicName : container2Title;
            string vignetteTitle = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteTitle).Value != null ? 
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteTitle).Value : "";
            string vignetteBody = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteBody).Value != null ? 
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.VignetteBody).Value : "";

            //create vertical node
            var verticalNode = xml.CreateElement("vertical");
            verticalNode.SetAttribute("display_name", topicName);
            verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-vertical-{0}-{1}", seqContainerRef, topicName), CourseTypes.Mock));
            verticalNode.SetAttribute("study_session_test_id", "");
            verticalNode.SetAttribute("taxon_id", topicTaxonId);
            verticalNode.SetAttribute("vignette_title", vignetteTitle);
            verticalNode.SetAttribute("vignette_body", vignetteBody);

            //skip vignette row, if there is any
            var topicQuestions = mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Question).HaveValue() &&
                mockRows.First().FirstOrDefault(c => c.Type == MockExamExcelColumnType.Question).Value != null ? mockRows : mockRows.Skip(1);

            var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, topicQuestions, new ProblemBuilderNodeSettings
            {
                DisplayName = String.Format("Mock exam questions - {0} - {1}", seqContainerRef, topicName),
                UrlName = CourseConverterHelper.getGuid(String.Format("mock-progress-test-{0}-{1}", seqContainerRef, topicName), CourseTypes.Mock),
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
