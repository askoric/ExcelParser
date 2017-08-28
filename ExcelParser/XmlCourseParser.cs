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
                string element_id = node.Attributes["cfa_short_name"] != null ? node.Attributes["cfa_short_name"].Value : "";
                string generated_id = node.Attributes["url_name"].Value;

                if (element_id != "")
                {
                    string type = node.Attributes["cfa_type"].Value;
                    if (type == "mock_exam")
                    {
                        char index = element_id.Last();
                        element_id = "MockExam" + index + "ChapterNode";
                    }
                    else if (type == "final_mock_exam")
                    {
                        element_id = "FinalMockExamChapterNode";
                    }
                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Topic);
                }
				
			}

            //MOCK EXAMS SEQUENTIALS
            XmlNodeList mockExamSequentials = courseXml.GetElementsByTagName("sequential");
            foreach (XmlNode node in mockExamSequentials)
            {
                string displayName = node.Attributes["display_name"].Value;
                string fcmNumber = node.Attributes["taxon_id"] != null ? node.Attributes["taxon_id"].Value : "";
                string element_id = "";
                if (fcmNumber != "")
                {
                    var chapter = node.SelectSingleNode("..");
                    string type = chapter.Attributes["cfa_type"].Value;
                    char index = chapter.Attributes["cfa_short_name"].Value.Last();
                    if (type == "mock_exam")
                    {
                        element_id = String.Format("mock-{0}-sequential-{1}-{2}", index, displayName, fcmNumber);
                    }
                    else if (type == "final_mock_exam")
                    {
                        element_id = String.Format("final-mock-sequential-{0}-{1}", displayName, fcmNumber);
                    }
                    
                    string generated_id = node.Attributes["url_name"].Value;

                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Mock);
                }
            }

            //MOCK EXAMS VERTICALS
            XmlNodeList mockExamVerticals = courseXml.GetElementsByTagName("vertical");
            foreach (XmlNode node in mockExamVerticals)
            {
                string element_id = "";
                string topicName = node.Attributes["display_name"].Value;
                var sequential = node.SelectSingleNode("..");
                string displayName = sequential.Attributes["display_name"].Value;
                var chapter = sequential.SelectSingleNode("..");
                char index = chapter.Attributes["cfa_short_name"].Value.Last();
                string type = chapter.Attributes["cfa_type"].Value;
                if (type == "mock_exam")
                {
                    element_id = String.Format("mock-{0}-vertical-{1}-{2}", index, displayName, topicName);
                }
                else if (type == "final_mock_exam")
                {
                    element_id = String.Format("final-mock-vertical-{0}-{1}", displayName, topicName);
                }
                string generated_id = node.Attributes["url_name"].Value;

                Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Mock);
            }

            //MOCK EXAMS PROBLEM BUILDERS
            XmlNodeList mockExamProblemBuilders = courseXml.GetElementsByTagName("problem-builder-mock-exam");
            foreach (XmlNode node in mockExamProblemBuilders)
            {
                string element_id = "";
                var vertical = node.SelectSingleNode("..");
                string topicName = vertical.Attributes["display_name"].Value;
                var sequential = vertical.SelectSingleNode("..");
                string displayName = sequential.Attributes["display_name"].Value;
                var chapter = sequential.SelectSingleNode("..");
                char index = chapter.Attributes["cfa_short_name"].Value.Last();
                string type = chapter.Attributes["cfa_type"].Value;
                if (type == "mock_exam")
                {
                    element_id = String.Format("mock-{0}-progress-test-{1}-{2}", index, displayName, topicName);
                }
                else if (type == "final_mock_exam")
                {
                    element_id = String.Format("final-mock-progress-test-{0}-{1}", displayName, topicName);
                }
                string generated_id = node.Attributes["url_name"].Value;

                Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Mock);
            }

            //MOCK EXAMS QUESTIONS
            XmlNodeList mockExamsQuestions = courseXml.GetElementsByTagName("pb-mcq-mock-exam");
            foreach (XmlNode node in mockExamsQuestions)
            {
                string element_id = node.Attributes["fitch_question_id"].Value;
                string generated_id = node.Attributes["url_name"].Value;

                Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Question);
            }

            //STUDY SESSION
            XmlNodeList studySessionNodes = courseXml.GetElementsByTagName( "container" );
			foreach ( XmlNode node in studySessionNodes )
			{
				var type = CourseTypes.Concept;
				if (node.Attributes["acceptance_criteria"] != null)
				{
					type = CourseTypes.Band;
				}

				string element_id = node.Attributes["learning_objective_id"].Value;
				string generated_id = node.Attributes["url_name"].Value;

				Database.Instance.AddKeyIfDoesntExists( element_id, generated_id, type );
			}

			//QUESTION
			XmlNodeList questionNodes = courseXml.GetElementsByTagName( "problem-builder-block" );
			foreach ( XmlNode node in questionNodes ) {

                string element_id = node.Attributes["atom_id"] != null ? node.Attributes["atom_id"].Value : "" ;
				string generated_id = node.Attributes["url_name"].Value;

                if (element_id != "")
                {
                    Database.Instance.AddKeyIfDoesntExists(element_id, generated_id, CourseTypes.Question);
                }
			}

			//VIDEO
			XmlNodeList videoNodes = courseXml.GetElementsByTagName( "brightcove-video" );
			foreach ( XmlNode node in videoNodes ) {
				string element_id = node.Attributes["atom_id"].Value;
				string generated_id = node.Attributes["url_name"].Value;

				Database.Instance.AddKeyIfDoesntExists( element_id, generated_id, CourseTypes.Video );
			}
        }

	}
}
