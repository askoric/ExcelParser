using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
	public class XmlCourseParser
	{
		public static void FillDbIdsFromCourseXml(XmlDocument courseXml)
		{
			XmlElement root = courseXml.DocumentElement;
			//TOPIC
			XmlNodeList chapterNodes = courseXml.GetElementsByTagName("chapter");
			foreach ( XmlNode node in chapterNodes )
			{
                string generated_id = node.Attributes["url_name"].Value;

                if (generated_id != "")
                {
                    var sequentialNode = node.FirstChild;
                    var fcmNumber = sequentialNode.Attributes["taxon_id"].Value;
                    fcmNumber = fcmNumber.Remove(fcmNumber.Length - 3);
                    var element_id = "MockExamChapterNode " + fcmNumber;
                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Topic);
                }
				
			}

            //MOCK EXAMS SEQUENTIALS
            XmlNodeList mockExamSequentials = courseXml.GetElementsByTagName("sequential");
            foreach (XmlNode node in mockExamSequentials)
            {
                string generated_id = node.Attributes["url_name"].Value;

                if (generated_id != "")
                {
                    var fcmNumber = node.Attributes["taxon_id"].Value;
                    var element_id = "mock-sequential-" + fcmNumber;
                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Mock);
                }
            }

            //MOCK EXAMS VERTICALS
            XmlNodeList mockExamVerticals = courseXml.GetElementsByTagName("vertical");
            foreach (XmlNode node in mockExamVerticals)
            {
                string generated_id = node.Attributes["url_name"].Value;
                
                if (generated_id != "")
                {
                    string topicName = node.Attributes["display_name"].Value;
                    var seqNode = node.SelectSingleNode("..");
                    var fcmNumber = seqNode.Attributes["taxon_id"].Value;
                    var element_id = String.Format("mock-vertical-{0}-{1}", fcmNumber, topicName);
                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Mock);
                }
            }

            //MOCK EXAMS PROBLEM BUILDERS
            XmlNodeList mockExamProblemBuilders = courseXml.GetElementsByTagName("problem-builder-mock-exam");
            foreach (XmlNode node in mockExamProblemBuilders)
            {
                string generated_id = node.Attributes["url_name"].Value;

                if (generated_id != "")
                {
                    var verNode = node.SelectSingleNode("..");
                    string topicName = verNode.Attributes["display_name"].Value;
                    var seqNode = verNode.SelectSingleNode("..");
                    var fcmNumber = seqNode.Attributes["taxon_id"].Value;
                    var element_id = String.Format("mock-progress-test-{0}-{1}", fcmNumber, topicName);
                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Mock);
                }
            }

            //MOCK EXAMS QUESTIONS
            XmlNodeList mockExamsQuestions = courseXml.GetElementsByTagName("pb-mcq-mock-exam");
            foreach (XmlNode node in mockExamsQuestions)
            {
                string element_id = node.Attributes["fitch_question_id"].Value;
                string generated_id = node.Attributes["url_name"].Value;

                Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Question);
            }

            ////STUDY SESSION
            //XmlNodeList studySessionNodes = courseXml.GetElementsByTagName("container");
            //foreach (XmlNode node in studySessionNodes)
            //{
            //    var type = CourseTypes.Concept;
            //    if (node.Attributes["acceptance_criteria"] != null)
            //    {
            //        type = CourseTypes.Band;
            //    }

            //    string element_id = node.Attributes["learning_objective_id"].Value;
            //    string generated_id = node.Attributes["url_name"].Value;

            //    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, type);
            //}

            ////QUESTION
            //XmlNodeList questionNodes = courseXml.GetElementsByTagName("problem-builder-block");
            //foreach (XmlNode node in questionNodes)
            //{

            //    string element_id = node.Attributes["atom_id"] != null ? node.Attributes["atom_id"].Value : "";
            //    string generated_id = node.Attributes["url_name"].Value;

            //    if (element_id != "")
            //    {
            //        Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Question);
            //    }
            //}

            ////VIDEO
            //XmlNodeList videoNodes = courseXml.GetElementsByTagName("brightcove-video");
            //foreach (XmlNode node in videoNodes)
            //{
            //    string element_id = node.Attributes["atom_id"].Value;
            //    string generated_id = node.Attributes["url_name"].Value;

            //    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Video);
            //}
        }

	}
}
