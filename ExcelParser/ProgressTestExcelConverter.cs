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

            string progressTestId = progressTestExcel.Rows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicWorkshopReference) != null ? progressTestExcel.Rows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicWorkshopReference).Value : "";

            var essayRows = progressTestExcel.Rows.Where(r => r.Any(c => c.Type == TestExcelColumnType.ItemSetType && c.Value.Contains("Essay")));
            var questionRows = progressTestExcel.Rows.Where(r => r.Any(c => c.Type == TestExcelColumnType.ItemSetType && c.Value.Contains("Item Set")));

            if (!questionRows.Any())
            {
                questionRows = progressTestExcel.Rows;
            }

            string progressTestPdf = questionRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetPdf) != null ? questionRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetPdf).Value : "";

            chapterNode.SetAttribute("topic_pdf", progressTestPdf);

            var topicReferences = questionRows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.TopicAbbrevation).Value);

            foreach (var topicRef in topicReferences)
            {
                string topicRefValue = topicRef.Key;
                var topicRows = questionRows.Where(r => r.Any(c => c.Type == TestExcelColumnType.TopicAbbrevation && c.Value.Contains(topicRefValue)));

                var itemSetReferences = topicRows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.ItemSetReference).Value);

                foreach (var itemSetReference in itemSetReferences)
                {
                    string itemSetReferenceValue = itemSetReference.Key;
                    char index = itemSetReferenceValue.Last();
                    var itemSetRows = topicRows.Where(r => r.Any(c => c.Type == TestExcelColumnType.ItemSetReference && c.Value.Contains(itemSetReferenceValue)));

                    if (itemSetRows.Any())
                    {
                        string kStructure = String.Join("|", itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.KStructure).Value.Split('|').Take(2));
                        string topicName = itemSetRows.First().FirstOrDefault(tn => tn.Type == TestExcelColumnType.TopicName).Value;
                        string topicAbbrevation = itemSetRows.First().FirstOrDefault(tn => tn.Type == TestExcelColumnType.TopicAbbrevation).Value;
                        //if index is 1 leave old progress test topic id, else change it
                        string topicId = index == '1' ? String.Format("{0}-r-progressTest", topicAbbrevation) : topicId = String.Format("{0}-r-progressTest-{1}-itemSet", topicAbbrevation, index);

                        var sequentialNode = GetProgressTestSequantialNode(xml, topicName, topicId, kStructure, itemSetRows);
                        chapterNode.AppendChild(sequentialNode);
                    }
                }
            }

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
            chapterNode.SetAttribute("test_duration", ifItemSet ? "01:48" : "02:00");


            if (essayRows.Any())
            {
                chapterNode.SetAttribute("exam_type", "essay");

                string essaysPdfQuestions = essayRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.EssaysPdfQuestions).Value;
                string essaysPdfAnswers = essayRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.EssaysPdfAnswers).Value;

                var sequentialNode = xml.CreateElement("sequential");
                sequentialNode.SetAttribute("display_name", "Progress Test Essays");
                sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(String.Format("{0}--essays", progressTestId), CourseTypes.StudySession));
                sequentialNode.SetAttribute("cfa_type", "essay");
                sequentialNode.SetAttribute("pdf_answers", essaysPdfAnswers);
                sequentialNode.SetAttribute("pdf_questions", essaysPdfQuestions);

                foreach (var row in essayRows)
                {
                    string essayTitle = row.FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetTitle).Value;
                    string topicTaxonId = row.FirstOrDefault(c => c.Type == TestExcelColumnType.TopicTaxonId).Value;
                    string essayTopics = row.FirstOrDefault(c => c.Type == TestExcelColumnType.TopicAbbrevation).Value;
                    string essayMaxPoints = row.FirstOrDefault(c => c.Type == TestExcelColumnType.EssayMaxPoints).Value;
                    string essayReferenceValue = row.FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetReference).Value;

                    var verticalNode = xml.CreateElement("vertical");
                    verticalNode.SetAttribute("cfa_type", "essay");
                    verticalNode.SetAttribute("item_set_id", essayReferenceValue);
                    verticalNode.SetAttribute("display_name", essayTitle);
                    verticalNode.SetAttribute("taxon_id", topicTaxonId);
                    verticalNode.SetAttribute("item_set_sessions", essayTopics);
                    verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(essayReferenceValue, CourseTypes.Essay));
                    verticalNode.SetAttribute("essay_max_points", essayMaxPoints);

                    sequentialNode.AppendChild(verticalNode);
                }

                chapterNode.AppendChild(sequentialNode);
            }

            return chapterNode;
        }


        private static XmlNode GetProgressTestSequantialNode(XmlDocument xml, string topicName, string topicId, string kStructure, IEnumerable<List<IExcelColumn<TestExcelColumnType>>> topicGroup)
        {
            var sequentialNode = xml.CreateElement("sequential");
            sequentialNode.SetAttribute("display_name", topicName);
            sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(topicId, CourseTypes.StudySession));
            sequentialNode.SetAttribute("taxon_id", kStructure);

            string itemSetTitle = topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetTitle) != null ? topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetTitle).Value : "";
            string vignetteTitle = topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteTitle) != null ? topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteTitle).Value : "";
            string vignetteBody = topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteBody) != null ? topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteBody).Value : "";
            string topicTaxonId = topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicTaxonId) != null ? topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicTaxonId).Value : "";

            //if item set title empty leave old vertical display name, if not change it
            string displayName = (itemSetTitle == "") ? "Progress test - R" : itemSetTitle;

            var verticalNode = xml.CreateElement("vertical");
            verticalNode.SetAttribute("display_name", displayName);
            verticalNode.SetAttribute("study_session_test_id", "");
            verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(topicId, CourseTypes.Reading));
            verticalNode.SetAttribute("taxon_id", topicTaxonId);
            verticalNode.SetAttribute("vignette_title", vignetteTitle);
            verticalNode.SetAttribute("vignette_body", vignetteBody);

            sequentialNode.AppendChild(verticalNode);

            //skip vignette row. if there is any
            topicGroup = topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.Question).HaveValue() && topicGroup.First().FirstOrDefault(c => c.Type == TestExcelColumnType.Question) != null ? topicGroup : topicGroup.Skip(1);

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
