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

            //divide practice mocks
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
                    var mockExamChapterNode = GetMockExamChapterNode(xml, mockRows);
                    mockExamChapterNode.SetAttribute("display_name", "Mock Examination " + index);
                    mockExamChapterNode.SetAttribute("cfa_type", "mock_exam");
                    mockExamChapterNode.SetAttribute("cfa_short_name", "Mock Exam " + index);
                    chapterNodes.Add(mockExamChapterNode);
                }
            }


            //work on final mock exam
            var finalExamChapterNode = GetMockExamChapterNode(xml, finalMockRows);
            finalExamChapterNode.SetAttribute("display_name", "Final Mock Examination");
            finalExamChapterNode.SetAttribute("cfa_type", "final_mock_exam");
            finalExamChapterNode.SetAttribute("cfa_short_name", "Final Mock Exam");
            chapterNodes.Add(finalExamChapterNode);

            return chapterNodes;
        }

        private static XmlElement GetMockExamChapterNode(XmlDocument xml, IEnumerable<List<IExcelColumn<MockExamExcelColumnType>>> mockRows)
        {

            //create chapter node
            var chapterNode = xml.CreateElement("chapter");
            chapterNode.SetAttribute("test_duration", "03:00");
            //chapterNode.SetAttribute("url_name", CourseConverterHelper.getGuid("MockExam" + index + "ChapterNode", CourseTypes.Topic)); TRIBA PROMINIT

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

            return chapterNode;
        }
    }
}
