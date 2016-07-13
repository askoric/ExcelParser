using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Newtonsoft.Json;

namespace ExcelParser
{
	public class XmlCourseConverter
	{
		private Dictionary<string, string> generatedQuestionIds;

		public XmlDocument ConvertExcelToXml( Excel excel )
		{
			generatedQuestionIds = new Dictionary<string, string>();
			XmlDocument xml = new XmlDocument();

			XmlElement rootNode = xml.CreateElement( "xbundle" );
			xml.AppendChild( rootNode );

			XmlElement courseNode = xml.CreateElement( "course" );
			courseNode.SetAttribute( "advanced_modules", "[&quot;annotatable&quot;, &quot;videoalpha&quot;, &quot;openassessment&quot;, &quot;container&quot;, &quot;problem-builder-block&quot;]" );
			courseNode.SetAttribute( "display_name", "CFA sample 1" );
			courseNode.SetAttribute( "language", "en" );
			courseNode.SetAttribute( "start", "&quot;2030-01-01T00:00:00+00:00&quot;" );
			courseNode.SetAttribute( "org", "s" );
			courseNode.SetAttribute( "course", "s" );
			courseNode.SetAttribute( "url_name_orig", "course" );
			courseNode.SetAttribute( "semester", "course" );
			rootNode.AppendChild( courseNode );

			XmlElement chapterNode = null;
			XmlElement sequentialNode = null;
			XmlElement verticalNode = null;
			XmlElement bandContainerNode = null;
			XmlElement conceptNameContainerNode = null;

			bool skip = false;

			foreach ( var row in excel.Rows ) {

				//Band need to be legal value else skip this row
				var bandColumn = row.FirstOrDefault( c => c.Type == ColumnType.Band );
				if ( bandColumn == null || bandColumn.Value == null || bandColumn.Value[0].ToString().ToLower() != "b" ) {
					continue;
				}

				var topicNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.TopicName );
				skip = (chapterNode != null && chapterNode.GetAttribute( "display_name" ) != topicNameColumn.Value) ? false : skip;
				if ( !skip ) {
					chapterNode = xml.CreateElement( "chapter" );
					chapterNode.SetAttribute( "display_name", topicNameColumn != null && topicNameColumn.HaveValue() ? topicNameColumn.Value.Replace( "/n", "<br>" ) : "" );
					chapterNode.SetAttribute( "url_name", getGuid() );
					courseNode.AppendChild( chapterNode );
				}

				var sessionNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.SessionName );
				skip = (sequentialNode != null && sequentialNode.GetAttribute( "display_name" ) != sessionNameColumn.Value) ? false : skip;
				if ( !skip ) {
					sequentialNode = xml.CreateElement( "sequential" );
					sequentialNode.SetAttribute( "display_name", sessionNameColumn != null && sessionNameColumn.HaveValue() ? sessionNameColumn.Value.Replace( "/n", "<br>" ) : "" );
					sequentialNode.SetAttribute( "url_name", getGuid() );
					chapterNode.AppendChild( sequentialNode );
				}

				var readingNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.ReadingName );
				skip = (verticalNode != null && verticalNode.GetAttribute( "display_name" ) != readingNameColumn.Value) ? false : skip;
				if ( !skip ) {
					verticalNode = xml.CreateElement( "vertical" );
					verticalNode.SetAttribute( "display_name", readingNameColumn != null && readingNameColumn.HaveValue() ? readingNameColumn.Value.Replace( "/n", "<br>" ) : "" );
					verticalNode.SetAttribute( "url_name", getGuid() );
					sequentialNode.AppendChild( verticalNode );
				}

				skip = (bandContainerNode != null && bandContainerNode.GetAttribute( "display_name" ) != bandColumn.Value) ? false : skip;
				if ( !skip ) {
					bandContainerNode = xml.CreateElement( "container" );
					bandContainerNode.SetAttribute( "url_name", getGuid() );
					bandContainerNode.SetAttribute( "display_name", bandColumn != null && bandColumn.HaveValue() ? bandColumn.Value.Replace( "/n", "<br>" ) : "" );
					bandContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					bandContainerNode.SetAttribute( "container_description", "" );
					verticalNode.AppendChild( bandContainerNode );
				}

				var conceptNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.ConceptName );
				skip = (conceptNameContainerNode != null && conceptNameContainerNode.GetAttribute( "display_name" ) != conceptNameColumn.Value) ? false : skip;
				if ( !skip ) {
					conceptNameContainerNode = xml.CreateElement( "container" );
					conceptNameContainerNode.SetAttribute( "url_name", getGuid() );
					conceptNameContainerNode.SetAttribute( "display_name", conceptNameColumn != null && conceptNameColumn.HaveValue() ? conceptNameColumn.Value.Replace( "/n", "<br>" ) : "" );
					conceptNameContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					conceptNameContainerNode.SetAttribute( "container_description", "" );
					bandContainerNode.AppendChild( conceptNameContainerNode );
				}

				//QUESTION
				var questionDic = generateQuestionIds();
				var problemBuilderNode = xml.CreateElement( "problem-builder-block" );
				var questionIdColumn = row.FirstOrDefault( c => c.Type == ColumnType.QuestionId );
				var questionColumn = row.FirstOrDefault( c => c.Type == ColumnType.Question );
				string questionValue = questionColumn.HaveValue() ? questionColumn.Value : "Question Missing";

				problemBuilderNode.SetAttribute( "display_name", questionValue.Replace( "/n", "" ) );
				problemBuilderNode.SetAttribute( "url_name", getGuid() );
				problemBuilderNode.SetAttribute( "xblock-family", "xblock.v1" );
				conceptNameContainerNode.AppendChild( problemBuilderNode );

				var pbMcqNode = xml.CreateElement( "pb-mcq-block" );
				var correctColumn = row.FirstOrDefault( c => c.Type == ColumnType.Correct );

				var actualCorrectValues = new List<string>();

				if ( correctColumn != null && correctColumn.HaveValue() ) {
					var correctValues = correctColumn.Value.Split( ' ' );

					foreach ( var correctValue in correctValues ) {
						actualCorrectValues.Add( questionDic[correctValue] );
					}
				}

				pbMcqNode.SetAttribute( "url_name", getGuid() );
				pbMcqNode.SetAttribute( "xblock-family", "xblock.v1" );
				pbMcqNode.SetAttribute( "question", questionValue.Replace( "/n", "<br>" ) );
				pbMcqNode.SetAttribute( "fitch_question_id", questionIdColumn.Value );
				pbMcqNode.SetAttribute( "correct_choices", (correctColumn != null && correctColumn.Value != null) ? JsonConvert.SerializeObject( actualCorrectValues ) : "" );

				var questionImageUrlColumn = row.FirstOrDefault( c => c.Type == ColumnType.QuestionImageUrl );
				var answerImageUrlColumn = row.FirstOrDefault( c => c.Type == ColumnType.AnswerImageUrl );

				if ( questionImageUrlColumn != null && questionImageUrlColumn.HaveValue() ) {
					pbMcqNode.SetAttribute( "image", questionImageUrlColumn.Value );
				}

				if ( answerImageUrlColumn != null && answerImageUrlColumn.HaveValue() ) {
					pbMcqNode.SetAttribute( "answer_image", answerImageUrlColumn.Value );
				}

				problemBuilderNode.AppendChild( pbMcqNode );

				var questionIds = new List<string>();
				string guid = Guid.NewGuid().ToString();

				var answer1Column = row.FirstOrDefault( c => c.Type == ColumnType.Answer1 );
				var question1Id = questionDic["A"];
				var answer1Node = GetAnswerNode( xml, answer1Column, question1Id, false );
				if ( answer1Node != null ) {
					pbMcqNode.AppendChild( answer1Node );
					questionIds.Add( question1Id );
				}

				var answer2Column = row.FirstOrDefault( c => c.Type == ColumnType.Answer2 );
				var question2Id = questionDic["B"];
				var answer2Node = GetAnswerNode( xml, answer2Column, question2Id, true );
				if ( answer2Node != null ) {
					pbMcqNode.AppendChild( answer2Node );
					questionIds.Add( question2Id );
				}

				var answer3Column = row.FirstOrDefault( c => c.Type == ColumnType.Answer3 );
				var question3Id = questionDic["C"];
				var answer3Node = GetAnswerNode( xml, answer3Column, question3Id, true );
				if ( answer3Node != null ) {
					pbMcqNode.AppendChild( answer3Node );
					questionIds.Add( question3Id );
				}

				var answer4Column = row.FirstOrDefault( c => c.Type == ColumnType.Answer4 );
				var question4Id = questionDic["D"];
				var answer4Node = GetAnswerNode( xml, answer4Column, question4Id );
				if ( answer4Node != null ) {
					pbMcqNode.AppendChild( answer4Node );
					questionIds.Add( question4Id );
				}


				//tip  block
				var justificationCell = row.FirstOrDefault( c => c.Type == ColumnType.Justification );
				if ( justificationCell != null && justificationCell.HaveValue() ) {
					var questionTipNode = xml.CreateElement( "pb-tip-block" );
					questionTipNode.SetAttribute( "url_name", getGuid() );
					questionTipNode.SetAttribute( "xblock-family", "xblock.v1" );
					questionTipNode.SetAttribute( "values", JsonConvert.SerializeObject( questionIds ) );
					questionTipNode.InnerText = justificationCell.Value.Replace( "/n", "<br>" );
					pbMcqNode.AppendChild( questionTipNode );
				}

				skip = true;

			}

			XmlElement wikiNode = xml.CreateElement( "wiki" );
			wikiNode.SetAttribute("slug", "test.1.2015");
			courseNode.AppendChild(wikiNode);

			return xml;
		}

		private string getGuid()
		{
			return Guid.NewGuid().ToString().Replace( "-", "" );
		}

		private Dictionary<string, string> generateQuestionIds()
		{
			var dic = new Dictionary<string, string>();

			while ( !dic.ContainsKey( "D" ) ) {
				string guid = getGuid();
				string questionId = guid.Substring( guid.Length - 7 );
				if ( !generatedQuestionIds.ContainsKey( questionId ) ) {
					generatedQuestionIds[questionId] = questionId;

					if ( !dic.ContainsKey( "A" ) ) {
						dic["A"] = questionId;
					}
					else if ( !dic.ContainsKey( "B" ) ) {
						dic["B"] = questionId;
					}
					else if ( !dic.ContainsKey( "C" ) ) {
						dic["C"] = questionId;
					}
					else {
						dic["D"] = questionId;
					}
				}
			}

			return dic;

		}

		private XmlElement GetAnswerNode( XmlDocument xml, ExcelColumn answerColumn, string questionId, bool addMissingValue = false )
		{
			if ( answerColumn.HaveValue() || addMissingValue ) {
				var answerNode = xml.CreateElement( "pb-choice-block" );
				answerNode.SetAttribute( "url_name", getGuid() );
				answerNode.SetAttribute( "xblock-family", "xblock.v1" );
				answerNode.SetAttribute( "value", questionId );
				answerNode.InnerText = answerColumn.HaveValue() ? answerColumn.Value.Replace( "/n", "" ) : "Answer Missing";
				return answerNode;
			}

			return null;
		}
	}
}
