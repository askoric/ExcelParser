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
            var containerRefType = mockExamExcel.Rows.First().FirstOrDefault(tn => tn.Type == TestExcelColumnType.ContainerRef) != null &&
                mockExamExcel.Rows.First().FirstOrDefault(tn => tn.Type == TestExcelColumnType.ContainerRef).HaveValue() ? 
                TestExcelColumnType.ContainerRef : TestExcelColumnType.TopicWorkshopReference;
            var mockExamContainerReferencesValues = mockExamExcel.Rows.GroupBy(r => r.First(tn => tn.Type == containerRefType).Value);
            foreach (var mockExamContainerReferenceValue in mockExamContainerReferencesValues)
            {
                string containerReferenceValue = mockExamContainerReferenceValue.Key;
                mockExamContainerReferences.Add(containerReferenceValue.Remove(containerReferenceValue.Length - 3));
            }
            mockExamContainerReferences = mockExamContainerReferences.Distinct().ToList();
            foreach (var containerReference in mockExamContainerReferences)
            {
                char index = containerReference.Last();
                var excelRows = mockExamExcel.Rows.Where(r => r.Any(c => c.Type == containerRefType && c.Value.Contains(containerReference)));
                if (excelRows.Any())
                {
                    var chapterNode = xml.CreateElement("chapter");
                    chapterNode.SetAttribute("display_name", "Mock Examination " + index);
                    chapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExam" + index + "ChapterNode", CourseTypes.Topic));
                    chapterNode.SetAttribute("cfa_type", "mock_exam");
                    chapterNode.SetAttribute("cfa_short_name", "Mock Exam " + index);
                    chapterNode.SetAttribute("test_duration", "03:00");

                    var amRows = new List<List<IExcelColumn<TestExcelColumnType>>>();
                    var pmRows = new List<List<IExcelColumn<TestExcelColumnType>>>();
                    string amFcmNumber = "";
                    string pmFcmNumber = "";

                    foreach (var row in excelRows)
                    {
                        var fcmNumber = row.FirstOrDefault(tn => tn.Type == TestExcelColumnType.FcmNumber) != null &&
                            row.FirstOrDefault(tn => tn.Type == TestExcelColumnType.FcmNumber).HaveValue() ?
                            row.First(tn => tn.Type == TestExcelColumnType.FcmNumber).Value : row.First(tn => tn.Type == TestExcelColumnType.TopicWorkshopReference).Value;
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

            

            var itemSetRefType = rows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetReference) != null &&
                    rows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetReference).HaveValue() ? TestExcelColumnType.ItemSetReference : TestExcelColumnType.ContainerRef;

            var itemSetReferences = rows.GroupBy(r => r.First(tn => tn.Type == itemSetRefType).Value);

            foreach (var itemSetReference in itemSetReferences) 
            {
                string itemSetReferenceValue = itemSetReference.Key;
                var itemSetRows = rows.Where(r => r.Any(c => c.Type == itemSetRefType && c.Value.Contains(itemSetReferenceValue)));
                var topicNameGroup = itemSetRows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.TopicAbbrevation).Value);

                foreach (var topic in topicNameGroup)
                {
                    
                    string topicName = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicName).Value;
                    string topicTaxonId = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicTaxonId).Value;

                    string itemSetTitle = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetTitle) != null ?
                        topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetTitle).Value : "";
                    string vignetteTitle = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteTitle) != null ?
                        topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteTitle).Value : "";
                    string vignetteBody = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteBody) != null ?
                        topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteBody).Value : "";

                    var verticalNode = xml.CreateElement("vertical");
                    verticalNode.SetAttribute("display_name", topicName);
                    verticalNode.SetAttribute("study_session_test_id", "");
                    verticalNode.SetAttribute("taxon_id", topicTaxonId);

                    //if item set title empty leave old vertical display name, if not change it
                    topicName = (itemSetTitle == "") ? topicName : itemSetTitle;

                    verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("mock-{0}-vertical-{1}-{2}", index, displayName, topicName), CourseTypes.Mock));
                    verticalNode.SetAttribute("vignette_title", vignetteTitle);
                    verticalNode.SetAttribute("vignette_body", vignetteBody);

                    sequentialNode.AppendChild(verticalNode);

                    //skip vignette row. if there is any
                    var topicQuestions = topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.Question).HaveValue() &&
                        topic.First().FirstOrDefault(c => c.Type == TestExcelColumnType.Question) != null ? topic : topic.Skip(1);

                    var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, topicQuestions, new ProblemBuilderNodeSettings
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
            }

            return sequentialNode;
        }
    }
}
