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
				string element_id = node.Attributes["cfa_short_name"].Value;
				string generated_id = node.Attributes["url_name"].Value;


				Database.Instance.AddKeyIfDoesntExists( element_id, generated_id, CourseTypes.Topic );
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
				string element_id = node.Attributes["atom_id"].Value;
				string generated_id = node.Attributes["url_name"].Value;

				Database.Instance.AddKeyIfDoesntExists( element_id, generated_id, CourseTypes.Question );
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
