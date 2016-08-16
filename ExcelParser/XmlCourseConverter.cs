using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Excel.Log;
using Newtonsoft.Json;

namespace ExcelParser
{

	public class ExcelParser
	{
		private Dictionary<string, string> generatedQuestionIds;

		public List<string> GetVideoReferenceIds( Excel mainStructureExcel )
		{
			var referenceIds = new List<string>();
			var rowIndex = 0;
			foreach ( var row in mainStructureExcel.Rows ) {
				rowIndex++;
				var atomType = row.FirstOrDefault( c => c.Type == ColumnType.AtomType );
				if ( atomType == null || !atomType.HaveValue() ) {
					Program.Log.Info( String.Format( "Missing atom type for row {0}", rowIndex ) );
					continue;
				}

				if ( atomType.Value == "IN" ) {
					var atomIdColumn = row.FirstOrDefault( c => c.Type == ColumnType.AtomId );
					if ( atomIdColumn != null && atomIdColumn.HaveValue() ) {
						referenceIds.Add( atomIdColumn.Value );
					}
					else {
						Program.Log.Info( String.Format( "Missing atom Id for row {0}", rowIndex ) );
					}
				}
			}
			return referenceIds;
		}

		public XmlDocument ConvertExcelToCourseXml( Excel mainStructureExcel, Excel questionExcel, Excel losExcel, Excel acceptanceCriteriaExcel, bool setTranscripts )
		{
			generatedQuestionIds = new Dictionary<string, string>();
			XmlDocument xml = new XmlDocument();
			var xmlTranscriptAccessor = new XmlTranscriptAccessor();

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

			foreach ( var row in mainStructureExcel.Rows ) {

				//Band need to be legal value else skip this row
				var bandColumn = row.FirstOrDefault( c => c.Type == ColumnType.Band );
				if ( bandColumn == null || bandColumn.Value == null || bandColumn.Value[0].ToString().ToLower() != "b" ) {
					continue;
				}

				var atomType = row.FirstOrDefault( c => c.Type == ColumnType.AtomType );
				if ( atomType == null || !atomType.HaveValue() || !(atomType.Value == "IN" || atomType.Value == "Q") ) {
					continue;
				}

				var atomIdColumn = row.FirstOrDefault( c => c.Type == ColumnType.AtomId );
				List<ExcelColumn> questionRow = null;

				if ( atomType.Value == "Q" ) {
					if ( atomIdColumn == null || !atomIdColumn.HaveValue() ) {
						continue;
					}

					questionRow = questionExcel.Rows.FirstOrDefault( r => r.Any( c => c.Type == ColumnType.QuestionId && c.Value == atomIdColumn.Value ) );
					if ( questionRow == null ) {
						continue;
					}
				}


				var topicNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.TopicName );
				skip = (chapterNode != null && chapterNode.GetAttribute( "display_name" ) != topicNameColumn.Value) ? false : skip;
				if ( !skip ) {
					var topicShortName = row.FirstOrDefault( c => c.Type == ColumnType.TopicShortName );
					var examPercantage = row.FirstOrDefault( c => c.Type == ColumnType.ExamPercentage );
					var description = row.FirstOrDefault( c => c.Type == ColumnType.Description );
					chapterNode = xml.CreateElement( "chapter" );
					chapterNode.SetAttribute( "display_name", topicNameColumn != null && topicNameColumn.HaveValue() ? topicNameColumn.Value : "" );
					chapterNode.SetAttribute( "url_name", getGuid() );
					chapterNode.SetAttribute( "cfa_short_name", topicShortName != null && topicShortName.HaveValue() ? topicShortName.Value : "" );
					chapterNode.SetAttribute( "exam_percentage", examPercantage != null && examPercantage.HaveValue() ? examPercantage.Value : "" );
					chapterNode.SetAttribute( "description", description != null && description.HaveValue() ? description.Value : "" );
					courseNode.AppendChild( chapterNode );
				}

				var sessionNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.SessionName );
				skip = (sequentialNode != null && sequentialNode.GetAttribute( "display_name" ) != sessionNameColumn.Value) ? false : skip;
				if ( !skip ) {
					var studySession = row.FirstOrDefault( c => c.Type == ColumnType.StudySession );
					sequentialNode = xml.CreateElement( "sequential" );
					sequentialNode.SetAttribute( "display_name", sessionNameColumn != null && sessionNameColumn.HaveValue() ? sessionNameColumn.Value : "" );
					sequentialNode.SetAttribute( "url_name", getGuid() );
					sequentialNode.SetAttribute( "cfa_short_name", studySession != null && studySession.HaveValue() ? studySession.Value : "" );
					chapterNode.AppendChild( sequentialNode );
				}

				var readingNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.ReadingName );
				var downloadsColumn = row.FirstOrDefault( c => c.Type == ColumnType.Downloads );
				var downloads2Column = row.FirstOrDefault( c => c.Type == ColumnType.Downloads2 );
				var downloads = new List<string>();
				if ( downloadsColumn != null && downloadsColumn.HaveValue() )
					downloads.Add( downloadsColumn.Value );
				if ( downloads2Column != null && downloads2Column.HaveValue() )
					downloads.Add( downloads2Column.Value );

				skip = (verticalNode != null && verticalNode.GetAttribute( "display_name" ) != readingNameColumn.Value) ? false : skip;
				if ( !skip ) {
					var readingColumn = row.FirstOrDefault( c => c.Type == ColumnType.Reading );
					verticalNode = xml.CreateElement( "vertical" );
					verticalNode.SetAttribute( "display_name", readingNameColumn != null && readingNameColumn.HaveValue() ? readingNameColumn.Value : "" );
					verticalNode.SetAttribute( "url_name", getGuid() );
					verticalNode.SetAttribute( "cfa_short_name", readingColumn != null && readingColumn.HaveValue() ? readingColumn.Value : "" );
					verticalNode.SetAttribute( "downloads", JsonConvert.SerializeObject( downloads ) );

					if ( losExcel != null ) {
						var outcomes = new List<object>();

						var losRows = losExcel.Rows.Where( r =>
							r.Any( c => c.Type == ColumnType.ReadingName && c.Value == readingNameColumn.Value ) &&
							r.Any( c => c.Type == ColumnType.TopicName && c.Value == topicNameColumn.Value ) &&
							r.Any( c => c.Type == ColumnType.SessionName && c.Value == sessionNameColumn.Value )
							);

						foreach ( var losRow in losRows ) {
							var cfaAlpfaColumn = losRow.Find( c => c.Type == ColumnType.CfaAlpha );
							var losTextColumn = losRow.Find( c => c.Type == ColumnType.LosText );

							outcomes.Add( new {
								text = losTextColumn != null ? losTextColumn.Value : "",
								letter = cfaAlpfaColumn != null ? cfaAlpfaColumn.Value : ""
							} );
						}

						verticalNode.SetAttribute( "outcome_statements", JsonConvert.SerializeObject( outcomes ) );

					}

					sequentialNode.AppendChild( verticalNode );
				}

				var conceptNameColumn = row.FirstOrDefault( c => c.Type == ColumnType.ConceptName );
				var conceptIdColumn = row.FirstOrDefault( c => c.Type == ColumnType.ConceptId );

				skip = (bandContainerNode != null && bandContainerNode.GetAttribute( "display_name" ) != bandColumn.Value) ? false : skip;
				if ( !skip ) {
					bandContainerNode = xml.CreateElement( "container" );
					bandContainerNode.SetAttribute( "url_name", getGuid() );
					bandContainerNode.SetAttribute( "display_name", bandColumn != null && bandColumn.HaveValue() ? bandColumn.Value : "" );
					bandContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					bandContainerNode.SetAttribute( "container_description", "" );
					bandContainerNode.SetAttribute( "learning_objective_id", "" );

					string targetScore = "";
					if ( acceptanceCriteriaExcel != null ) {
						var acceptanceCriteriaRow =
							acceptanceCriteriaExcel.Rows.FirstOrDefault(
								r => r.Any( c => c.Type == ColumnType.Lo1 && c.Value == conceptIdColumn.Value ) );

						if ( acceptanceCriteriaRow != null ) {
							var scoreColumn = acceptanceCriteriaRow.FirstOrDefault( c => c.Type == ColumnType.TargetScore );
							if ( scoreColumn != null && scoreColumn.HaveValue() ) {
								targetScore = scoreColumn.Value.Replace( "0.", "" );
							}
						}
					}

					bandContainerNode.SetAttribute( "acceptance_criteria", targetScore );

					verticalNode.AppendChild( bandContainerNode );
				}

				skip = (conceptNameContainerNode != null && conceptNameContainerNode.GetAttribute( "display_name" ) != conceptNameColumn.Value) ? false : skip;
				if ( !skip ) {

					conceptNameContainerNode = xml.CreateElement( "container" );
					conceptNameContainerNode.SetAttribute( "url_name", getGuid() );
					conceptNameContainerNode.SetAttribute( "display_name", conceptNameColumn != null && conceptNameColumn.HaveValue() ? conceptNameColumn.Value : "" );
					conceptNameContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					conceptNameContainerNode.SetAttribute( "container_description", "" );
					conceptNameContainerNode.SetAttribute( "learning_objective_id", conceptIdColumn != null && conceptIdColumn.HaveValue() ? conceptIdColumn.Value : "" );
					bandContainerNode.AppendChild( conceptNameContainerNode );
				}

				//VIDEO
				if ( atomType.Value == "IN" ) {
					var atomTitleColumn = row.FirstOrDefault( c => c.Type == ColumnType.AtomTitle );
					var itemIdCoclumn = row.FirstOrDefault( c => c.Type == ColumnType.ItemId );
					var videoNode = xml.CreateElement( "brightcove-video" );
					videoNode.SetAttribute( "url_name", getGuid() );
					videoNode.SetAttribute( "xblock-family", "xblock.v1" );
					videoNode.SetAttribute( "api_bckey", "AQ~~,AAAELMh4AWE~,vVFFDlX6sNOap1Tww7YwaMvqbQ8TtDoh" );
					videoNode.SetAttribute( "display_name", atomTitleColumn != null && atomTitleColumn.HaveValue() ? atomTitleColumn.Value : "" );
					videoNode.SetAttribute( "api_key", "JqnRdhYvLWNtVJllXkMzGGGTh66uLLmz8JB8YlcZQlC8OX94H4ZXXw.." );
					videoNode.SetAttribute( "text_values", "[]" );
					videoNode.SetAttribute( "api_bctid", itemIdCoclumn.Value );
					videoNode.SetAttribute( "begin_values", "[]" );
					videoNode.SetAttribute( "api_bcpid", "4830051907001" );
					videoNode.SetAttribute( "cfa_type", "video" );
					videoNode.SetAttribute( "atom_id", atomIdColumn != null && atomIdColumn.HaveValue() ? atomIdColumn.Value : "" );

					if ( setTranscripts ) {
						string xmlTranscriptString = "";
						if ( atomIdColumn != null && atomIdColumn.HaveValue() ) {
							XmlDocument xmlTranscript = xmlTranscriptAccessor.FindVideoTranscript( atomIdColumn.Value );
							if ( xmlTranscript != null ) {
								xmlTranscriptString = xmlTranscript.InnerXml.Replace( "<br />", "" ).Replace( "<br/>", "" );
							}

						}

						videoNode.SetAttribute( "xml_string", xmlTranscriptString );
					}
					else {
						videoNode.SetAttribute( "xml_string", "&#10;" );
					}

					conceptNameContainerNode.AppendChild( videoNode );
				}
				//QUESTION
				else {
					var questionDic = generateQuestionIds();
					var problemBuilderNode = xml.CreateElement( "problem-builder-block" );
					var questionIdColumn = questionRow.FirstOrDefault( c => c.Type == ColumnType.QuestionId );
					var questionColumn = questionRow.FirstOrDefault( c => c.Type == ColumnType.Question );
					var kkEeColumn = questionRow.FirstOrDefault( c => c.Type == ColumnType.KKEE );
					var answerImageUrlColumn = questionRow.FirstOrDefault( c => c.Type == ColumnType.AnswerImageUrl );
					string questionValue = questionColumn.HaveValue() ? questionColumn.Value : "Question Missing";

					problemBuilderNode.SetAttribute( "display_name", questionValue.Replace( "<br>", "" ) );
					problemBuilderNode.SetAttribute( "url_name", getGuid() );
					problemBuilderNode.SetAttribute( "xblock-family", "xblock.v1" );
					problemBuilderNode.SetAttribute( "cfa_type", "question" );
					problemBuilderNode.SetAttribute( "atom_id", atomIdColumn != null && atomIdColumn.HaveValue() ? atomIdColumn.Value : "" );
					problemBuilderNode.SetAttribute( "instruct_assessment", kkEeColumn != null && kkEeColumn.HaveValue() ? kkEeColumn.Value : "" );

					if ( answerImageUrlColumn != null && answerImageUrlColumn.HaveValue() ) {
						problemBuilderNode.SetAttribute( "answer_image", answerImageUrlColumn.Value );
					}

					conceptNameContainerNode.AppendChild( problemBuilderNode );

					var pbMcqNode = xml.CreateElement( "pb-mcq-block" );
					var correctColumn = questionRow.FirstOrDefault( c => c.Type == ColumnType.Correct );

					var actualCorrectValues = new List<string>();

					if ( correctColumn != null && correctColumn.HaveValue() ) {
						var correctValues = correctColumn.Value.Split( ' ' );

						foreach ( var correctValue in correctValues ) {
							actualCorrectValues.Add( questionDic[correctValue] );
						}
					}

					pbMcqNode.SetAttribute( "url_name", getGuid() );
					pbMcqNode.SetAttribute( "xblock-family", "xblock.v1" );
					pbMcqNode.SetAttribute( "question", questionValue );
					pbMcqNode.SetAttribute( "fitch_question_id", questionIdColumn.Value );
					pbMcqNode.SetAttribute( "correct_choices", (correctColumn != null && correctColumn.Value != null) ? JsonConvert.SerializeObject( actualCorrectValues ) : "" );

					var questionImageUrlColumn = questionRow.FirstOrDefault( c => c.Type == ColumnType.QuestionImageUrl );

					if ( questionImageUrlColumn != null && questionImageUrlColumn.HaveValue() ) {
						pbMcqNode.SetAttribute( "image", questionImageUrlColumn.Value );
					}

					problemBuilderNode.AppendChild( pbMcqNode );

					var questionIds = new List<string>();

					var answer1Column = questionRow.FirstOrDefault( c => c.Type == ColumnType.Answer1 );
					var question1Id = questionDic["A"];
					var answer1Node = GetAnswerNode( xml, answer1Column, question1Id, false );
					if ( answer1Node != null ) {
						pbMcqNode.AppendChild( answer1Node );
						questionIds.Add( question1Id );
					}

					var answer2Column = questionRow.FirstOrDefault( c => c.Type == ColumnType.Answer2 );
					var question2Id = questionDic["B"];
					var answer2Node = GetAnswerNode( xml, answer2Column, question2Id, true );
					if ( answer2Node != null ) {
						pbMcqNode.AppendChild( answer2Node );
						questionIds.Add( question2Id );
					}

					var answer3Column = questionRow.FirstOrDefault( c => c.Type == ColumnType.Answer3 );
					var question3Id = questionDic["C"];
					var answer3Node = GetAnswerNode( xml, answer3Column, question3Id, true );
					if ( answer3Node != null ) {
						pbMcqNode.AppendChild( answer3Node );
						questionIds.Add( question3Id );
					}

					var answer4Column = questionRow.FirstOrDefault( c => c.Type == ColumnType.Answer4 );
					var question4Id = questionDic["D"];
					var answer4Node = GetAnswerNode( xml, answer4Column, question4Id );
					if ( answer4Node != null ) {
						pbMcqNode.AppendChild( answer4Node );
						questionIds.Add( question4Id );
					}


					//tip  block
					var justificationCell = questionRow.FirstOrDefault( c => c.Type == ColumnType.Justification );
					if ( justificationCell != null && justificationCell.HaveValue() ) {
						var questionTipNode = xml.CreateElement( "pb-tip-block" );
						questionTipNode.SetAttribute( "url_name", getGuid() );
						questionTipNode.SetAttribute( "xblock-family", "xblock.v1" );
						questionTipNode.SetAttribute( "values", JsonConvert.SerializeObject( questionIds ) );
						questionTipNode.InnerText = justificationCell.Value;
						pbMcqNode.AppendChild( questionTipNode );
					}

				}

				skip = true;

			}

			XmlElement wikiNode = xml.CreateElement( "wiki" );
			wikiNode.SetAttribute( "slug", "test.1.2015" );
			courseNode.AppendChild( wikiNode );

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
