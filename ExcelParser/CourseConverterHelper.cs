using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
    class CourseConverterHelper
    {
        public static Dictionary<string, string> generatedQuestionIds;
        public static Dictionary<string, guidRequest> _generatedGuids;

        public static Dictionary<string, string> generateQuestionIds()
        {
            var dic = new Dictionary<string, string>();

            while (!dic.ContainsKey("D"))
            {

                var questionId = generateQuestionId();
                if (!generatedQuestionIds.ContainsKey(questionId))
                {
                    generatedQuestionIds[questionId] = questionId;

                    if (!dic.ContainsKey("A"))
                    {
                        dic["A"] = questionId;
                    }
                    else if (!dic.ContainsKey("B"))
                    {
                        dic["B"] = questionId;
                    }
                    else if (!dic.ContainsKey("C"))
                    {
                        dic["C"] = questionId;
                    }
                    else {
                        dic["D"] = questionId;
                    }
                }
            }

            return dic;

        }

        public static string getGuid(string elementId, CourseTypes elementType)
        {
            string key = Database.Instance.GetKey(elementId, elementType);
            if (String.IsNullOrEmpty(key))
            {
                key = getNewGuid();
                Database.Instance.AddKey(elementId, key, elementType);
                Program.Log.Info(String.Format("New Key generated elementType: {0}; element_id: {1}; generatedKey: {2}", elementType.ToString(), elementId, key));
            }

            if (_generatedGuids.ContainsKey(key))
            {
                var existing = _generatedGuids[key];
                Program.Log.Warn(
                    String.Format(
                        "DUPLICATE ID DETECTED >>GeneratedId = '{0}' ; existingReferenceId = {1}  existingType = {2}; newRefrenceId = {3} newType = {4} ",
                        key, existing.ElementId, existing.elementType, elementId, elementType));
            }
            else {
                _generatedGuids[key] = new guidRequest
                {
                    ElementId = elementId,
                    elementType = elementType
                };
            }

            return key;
        }


        public static XmlElement GetAnswerNode(XmlDocument xml, IExcelColumn<QuestionExcelColumnType> answerColumn, string questionId, bool addMissingValue = false)
        {
            return GetAnswerNode(xml, answerColumn != null && answerColumn.HaveValue() ? answerColumn.Value : "", questionId, addMissingValue);
        }

        public static XmlElement GetAnswerNode(XmlDocument xml, IExcelColumn<TestExcelColumnType> answerColumn, string questionId, bool addMissingValue = false)
        {
            return GetAnswerNode(xml, answerColumn != null && answerColumn.HaveValue() ? answerColumn.Value : "", questionId, addMissingValue);
        }

        public static XmlElement GetAnswerNode(XmlDocument xml, string answer, string questionId, bool addMissingValue = false)
        {
            if (!String.IsNullOrWhiteSpace(answer) || addMissingValue)
            {
                var answerNode = xml.CreateElement("pb-choice-block");
                answerNode.SetAttribute("url_name", getNewGuid());
                answerNode.SetAttribute("xblock-family", "xblock.v1");
                answerNode.SetAttribute("value", questionId);
                answerNode.InnerText = String.IsNullOrWhiteSpace(answer) ? "Answer Missing" : answer.Replace("/n", "");
                return answerNode;
            }

            return null;
        }

        public static string getNewGuid()
        {
            return Guid.NewGuid().ToString().Replace("-", "");
        }

        public static string generateQuestionId()
        {
            string guid = getNewGuid();
            return guid.Substring(guid.Length - 7);
        }
    }
}
