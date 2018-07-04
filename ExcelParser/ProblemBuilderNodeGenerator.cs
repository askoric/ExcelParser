using Newtonsoft.Json;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
    public class ProblemBuilderNodeSettings
    {
        public string DisplayName { get; set; }
        public string UrlName { get; set; }
        public string ProblemBuilderNodeElement { get; set; }
        public string PbMcqNodeElement { get; set; }
        public string PbChoiceBlockElement { get; set; }
        public string PbTipBlockElement { get; set; }
    }

    class ProblemBuilderNodeGenerator
    {
        public static XmlElement Generate(XmlDocument xml, IEnumerable<List<IExcelColumn<TestExcelColumnType>>> excelRows, ProblemBuilderNodeSettings settings)
        {
            var problemBuilderNode = xml.CreateElement(settings.ProblemBuilderNodeElement);
            problemBuilderNode.SetAttribute("display_name", settings.DisplayName);
            problemBuilderNode.SetAttribute("url_name", settings.UrlName);
            problemBuilderNode.SetAttribute("xblock-family", "xblock.v1");
            problemBuilderNode.SetAttribute("cfa_type", "question");

            foreach (var row in excelRows)
            {
                var questionDic = CourseConverterHelper.generateQuestionIds();
                var questionColumn = row.FirstOrDefault(c => c.Type == TestExcelColumnType.Question);
                var questionIdColumn = row.FirstOrDefault(c => c.Type == TestExcelColumnType.QuestionId);
                var questionImageUrlColumn = row.FirstOrDefault(c => c.Type == TestExcelColumnType.QuestionImageUrl);
                var answerImageUrlColumn = row.FirstOrDefault(c => c.Type == TestExcelColumnType.AnswerImageUrl);
                string questionValue = questionColumn.HaveValue() ? questionColumn.Value : questionImageUrlColumn.HaveValue() ? "" : "Question Missing";


                var pbMcqNode = xml.CreateElement(settings.PbMcqNodeElement);
                var correctColumn = row.FirstOrDefault(c => c.Type == TestExcelColumnType.Correct);

                var actualCorrectValues = new List<string>();

                if (correctColumn != null && correctColumn.HaveValue())
                {
                    var correctValues = correctColumn.Value.Split(' ');

                    foreach (var correctValue in correctValues)
                    {
                        actualCorrectValues.Add(questionDic[correctValue]);
                    }
                }

                pbMcqNode.SetAttribute("url_name", CourseConverterHelper.getGuid(questionIdColumn.Value, CourseTypes.Question));
                pbMcqNode.SetAttribute("xblock-family", "xblock.v1");
                pbMcqNode.SetAttribute("question", questionValue);
                pbMcqNode.SetAttribute("fitch_question_id", questionIdColumn.Value);
                pbMcqNode.SetAttribute("correct_choices", (correctColumn != null && correctColumn.Value != null) ? JsonConvert.SerializeObject(actualCorrectValues) : "");


                if (questionImageUrlColumn != null && questionImageUrlColumn.HaveValue())
                {
                    pbMcqNode.SetAttribute("image", questionImageUrlColumn.Value);
                }

                problemBuilderNode.AppendChild(pbMcqNode);

                var questionIds = new List<string>();

                var answer1Column = row.FirstOrDefault(c => c.Type == TestExcelColumnType.Answer1);
                var question1Id = questionDic["A"];
                var answer1Node = CourseConverterHelper.GetAnswerNode(xml, answer1Column, question1Id, false, settings.PbChoiceBlockElement);
                if (answer1Node != null)
                {
                    pbMcqNode.AppendChild(answer1Node);
                    questionIds.Add(question1Id);
                }

                var answer2Column = row.FirstOrDefault(c => c.Type == TestExcelColumnType.Answer2);
                var question2Id = questionDic["B"];
                var answer2Node = CourseConverterHelper.GetAnswerNode(xml, answer2Column, question2Id, true, settings.PbChoiceBlockElement);
                if (answer2Node != null)
                {
                    pbMcqNode.AppendChild(answer2Node);
                    questionIds.Add(question2Id);
                }

                var answer3Column = row.FirstOrDefault(c => c.Type == TestExcelColumnType.Answer3);
                var question3Id = questionDic["C"];
                var answer3Node = CourseConverterHelper.GetAnswerNode(xml, answer3Column, question3Id, true, settings.PbChoiceBlockElement);
                if (answer3Node != null)
                {
                    pbMcqNode.AppendChild(answer3Node);
                    questionIds.Add(question3Id);
                }


                //Harcoded answer 4 node
                var answer4Node = xml.CreateElement(settings.PbChoiceBlockElement);
                var question4Id = questionDic["D"];
                questionIds.Add(question4Id);
                answer4Node.SetAttribute("url_name", CourseConverterHelper.getNewGuid());
                answer4Node.SetAttribute("xblock-family", "xblock.v1");
                answer4Node.SetAttribute("value", question4Id);
                pbMcqNode.AppendChild(answer4Node);


                //tip  block
                var justificationCell = row.FirstOrDefault(c => c.Type == TestExcelColumnType.Justification);
                string justificationInnerText = "";
                if (answerImageUrlColumn != null && answerImageUrlColumn.HaveValue())
                {
                    justificationInnerText = String.Format("<div class='answer-image'><img src='{0}'></div>", answerImageUrlColumn.Value);
                }

                string justificationValue = (justificationCell != null && justificationCell.HaveValue())
                        ? justificationCell.Value
                        : "";

                var questionTipNode = xml.CreateElement(settings.PbTipBlockElement);
                questionTipNode.SetAttribute("url_name", CourseConverterHelper.getNewGuid());
                questionTipNode.SetAttribute("xblock-family", "xblock.v1");
                questionTipNode.SetAttribute("values", JsonConvert.SerializeObject(questionIds));
                questionTipNode.InnerText = String.Format("{0}{1}", justificationValue, justificationInnerText);
                pbMcqNode.AppendChild(questionTipNode);

            }

            return problemBuilderNode;
        }

        public static XmlElement Generate(XmlDocument xml, IEnumerable<List<IExcelColumn<ExamExcelColumnType>>> excelRows, ProblemBuilderNodeSettings settings)
        {
            var problemBuilderNode = xml.CreateElement(settings.ProblemBuilderNodeElement);
            problemBuilderNode.SetAttribute("display_name", settings.DisplayName);
            problemBuilderNode.SetAttribute("url_name", settings.UrlName);
            problemBuilderNode.SetAttribute("xblock-family", "xblock.v1");
            problemBuilderNode.SetAttribute("cfa_type", "question");

            foreach (var row in excelRows)
            {
                var questionDic = CourseConverterHelper.generateQuestionIds();
                var questionColumn = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.Question);
                var questionIdColumn = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.QuestionId);
                string questionValue = questionColumn.HaveValue() ? questionColumn.Value : "Question Missing";


                var pbMcqNode = xml.CreateElement(settings.PbMcqNodeElement);
                var correctColumn = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.Correct);

                var actualCorrectValues = new List<string>();

                if (correctColumn != null && correctColumn.HaveValue())
                {
                    var correctValues = correctColumn.Value.Split(' ');

                    foreach (var correctValue in correctValues)
                    {
                        actualCorrectValues.Add(questionDic[correctValue]);
                    }
                }

                pbMcqNode.SetAttribute("url_name", CourseConverterHelper.getGuid(questionIdColumn.Value, CourseTypes.Question));
                pbMcqNode.SetAttribute("xblock-family", "xblock.v1");
                pbMcqNode.SetAttribute("question", questionValue);
                pbMcqNode.SetAttribute("fitch_question_id", questionIdColumn.Value);
                pbMcqNode.SetAttribute("correct_choices", (correctColumn != null && correctColumn.Value != null) ? JsonConvert.SerializeObject(actualCorrectValues) : "");

                problemBuilderNode.AppendChild(pbMcqNode);

                var questionIds = new List<string>();

                var answer1Column = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.Answer1);
                var question1Id = questionDic["A"];
                var answer1Node = CourseConverterHelper.GetAnswerNode(xml, answer1Column, question1Id, false, settings.PbChoiceBlockElement);
                if (answer1Node != null)
                {
                    pbMcqNode.AppendChild(answer1Node);
                    questionIds.Add(question1Id);
                }

                var answer2Column = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.Answer2);
                var question2Id = questionDic["B"];
                var answer2Node = CourseConverterHelper.GetAnswerNode(xml, answer2Column, question2Id, true, settings.PbChoiceBlockElement);
                if (answer2Node != null)
                {
                    pbMcqNode.AppendChild(answer2Node);
                    questionIds.Add(question2Id);
                }

                var answer3Column = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.Answer3);
                var question3Id = questionDic["C"];
                var answer3Node = CourseConverterHelper.GetAnswerNode(xml, answer3Column, question3Id, true, settings.PbChoiceBlockElement);
                if (answer3Node != null)
                {
                    pbMcqNode.AppendChild(answer3Node);
                    questionIds.Add(question3Id);
                }


                //Harcoded answer 4 node
                var answer4Node = xml.CreateElement(settings.PbChoiceBlockElement);
                var question4Id = questionDic["D"];
                questionIds.Add(question4Id);
                answer4Node.SetAttribute("url_name", CourseConverterHelper.getNewGuid());
                answer4Node.SetAttribute("xblock-family", "xblock.v1");
                answer4Node.SetAttribute("value", question4Id);
                pbMcqNode.AppendChild(answer4Node);


                //tip  block
                var justificationCell = row.FirstOrDefault(c => c.Type == ExamExcelColumnType.Justification);
                string justificationInnerText = "";

                string justificationValue = (justificationCell != null && justificationCell.HaveValue())
                        ? justificationCell.Value
                        : "";

                var questionTipNode = xml.CreateElement(settings.PbTipBlockElement);
                questionTipNode.SetAttribute("url_name", CourseConverterHelper.getNewGuid());
                questionTipNode.SetAttribute("xblock-family", "xblock.v1");
                questionTipNode.SetAttribute("values", JsonConvert.SerializeObject(questionIds));
                questionTipNode.InnerText = String.Format("{0}{1}", justificationValue, justificationInnerText);
                pbMcqNode.AppendChild(questionTipNode);

            }

            return problemBuilderNode;
        }
    }
}
