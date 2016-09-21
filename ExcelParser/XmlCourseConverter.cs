﻿using System;
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

		public List<string> GetVideoReferenceIds( Excel<MainStructureExcelColumn, MainStructureColumnType> mainStructureExcel )
		{
			var referenceIds = new List<string>();
			var rowIndex = 0;
			foreach ( var row in mainStructureExcel.Rows ) {
				rowIndex++;
				var atomType = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomType );
				if ( atomType == null || !atomType.HaveValue() ) {
					Program.Log.Info( String.Format( "Missing atom type for row {0}", rowIndex ) );
					continue;
				}

				if ( atomType.Value == "IN" ) {
					var atomIdColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomId );
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

		public XmlDocument ConvertExcelToCourseXml( Excel<MainStructureExcelColumn, MainStructureColumnType> mainStructureExcel, Excel<QuestionExcelColumn, QuestionExcelColumnType> questionExcel, Excel<LosExcelColumn, LosExcelColumnType> losExcel, Excel<AcceptanceCriteriaExcelColumn, AcceptanceCriteriaColumnType> acceptanceCriteriaExcel, Excel<SsTestExcelColumn, SsTestExcelColumnType> ssTestExcel, bool setTranscripts )
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


			XmlElement PreviousSequentialNode = null;
			IExcelColumn<MainStructureColumnType> previousStudySessionId = null;
			string locked = "";

			foreach ( var row in mainStructureExcel.Rows ) {

				//Band need to be legal value else skip this row
				var bandColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Band );
				if ( bandColumn == null || bandColumn.Value == null || bandColumn.Value[0].ToString().ToLower() != "b" ) {
					continue;
				}

				var atomType = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomType );
				if ( atomType == null || !atomType.HaveValue() || !(atomType.Value == "IN" || atomType.Value == "Q") ) {
					continue;
				}

				var atomIdColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomId );
				List<IExcelColumn<QuestionExcelColumnType>> questionRow = null;

				if ( atomType.Value == "Q" ) {
					if ( atomIdColumn == null || !atomIdColumn.HaveValue() ) {
						continue;
					}

					questionRow = questionExcel.Rows.FirstOrDefault( r => r.Any( c => c.Type == QuestionExcelColumnType.QuestionId && c.Value == atomIdColumn.Value ) );
					if ( questionRow == null ) {
						continue;
					}
				}


				var topicNameColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.TopicName );
				skip = (chapterNode != null && chapterNode.GetAttribute( "display_name" ) != topicNameColumn.Value) ? false : skip;
				if ( !skip ) {
					var topicShortName = row.FirstOrDefault( c => c.Type == MainStructureColumnType.TopicShortName );
					var examPercantage = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ExamPercentage );
					var lockedColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Locked );
					var colorColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Color );
					var cfaTypeColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.CfaType );
					locked = lockedColumn != null && lockedColumn.HaveValue() ? lockedColumn.Value : "";

					var description = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Description );
					chapterNode = xml.CreateElement( "chapter" );
					chapterNode.SetAttribute( "display_name", topicNameColumn != null && topicNameColumn.HaveValue() ? topicNameColumn.Value : "" );
					chapterNode.SetAttribute( "url_name", getGuid( topicShortName.Value, CourseTypes.Topic ) );
					chapterNode.SetAttribute( "cfa_short_name", topicShortName != null && topicShortName.HaveValue() ? topicShortName.Value : "" );
					chapterNode.SetAttribute( "exam_percentage", examPercantage != null && examPercantage.HaveValue() ? examPercantage.Value : "" );
					chapterNode.SetAttribute( "description", description != null && description.HaveValue() ? description.Value : "" );
					chapterNode.SetAttribute( "locked", locked );
					chapterNode.SetAttribute( "topic_color", colorColumn != null && colorColumn.HaveValue() ? colorColumn.Value : "" );
					chapterNode.SetAttribute( "cfa_type", cfaTypeColumn != null && cfaTypeColumn.HaveValue() ? cfaTypeColumn.Value : "" );
					courseNode.AppendChild( chapterNode );
				}

				var sessionNameColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.SessionName );
				skip = (sequentialNode != null && sequentialNode.GetAttribute( "display_name" ) != sessionNameColumn.Value) ? false : skip;
				var studySessionId = row.FirstOrDefault( c => c.Type == MainStructureColumnType.StudySessionId );
				if ( !skip ) {
					var studySession = row.FirstOrDefault( c => c.Type == MainStructureColumnType.StudySession );
					sequentialNode = xml.CreateElement( "sequential" );
					sequentialNode.SetAttribute( "display_name",
						sessionNameColumn != null && sessionNameColumn.HaveValue() ? sessionNameColumn.Value : "" );
					sequentialNode.SetAttribute( "url_name", getGuid( studySessionId.Value, CourseTypes.StudySession ) );
					sequentialNode.SetAttribute( "cfa_short_name",
						studySession != null && studySession.HaveValue() ? studySession.Value : "" );
					chapterNode.AppendChild( sequentialNode );

					//ADD TEST TO THE BOTTOM OF LAST SESSION NAME NODE
					if ( PreviousSequentialNode != null && previousStudySessionId != null && locked != "yes" ) {
						AppendStudySessionTestQuestions( xml, PreviousSequentialNode, previousStudySessionId.Value, ssTestExcel );
					}

					PreviousSequentialNode = sequentialNode;
					previousStudySessionId = studySessionId;
				}

				var readingNameColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ReadingName );
				var downloadsColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Downloads );
				var downloads2Column = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Downloads2 );
				var downloads = new List<string>();
				if ( downloadsColumn != null && downloadsColumn.HaveValue() )
					downloads.Add( downloadsColumn.Value );
				if ( downloads2Column != null && downloads2Column.HaveValue() )
					downloads.Add( downloads2Column.Value );

				skip = (verticalNode != null && verticalNode.GetAttribute( "display_name" ) != readingNameColumn.Value) ? false : skip;
				if ( !skip ) {
					var readingColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Reading );
					var readingIdColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ReadingId );
					verticalNode = xml.CreateElement( "vertical" );
					verticalNode.SetAttribute( "display_name", readingNameColumn != null && readingNameColumn.HaveValue() ? readingNameColumn.Value : "" );
					verticalNode.SetAttribute( "url_name", getGuid( readingIdColumn.Value, CourseTypes.Reading ) );
					verticalNode.SetAttribute( "cfa_short_name", readingColumn != null && readingColumn.HaveValue() ? readingColumn.Value : "" );
					verticalNode.SetAttribute( "downloads", JsonConvert.SerializeObject( downloads ) );

					if ( losExcel != null ) {
						var outcomes = new List<object>();

						var losRows = losExcel.Rows.Where( r =>
							r.Any( c => c.Type == LosExcelColumnType.ReadingTitle && c.Value == readingNameColumn.Value ) &&
							r.Any( c => c.Type == LosExcelColumnType.TopicTitle && c.Value == topicNameColumn.Value ) &&
							r.Any( c => c.Type == LosExcelColumnType.SessionTitle && c.Value == sessionNameColumn.Value )
							);

						foreach ( var losRow in losRows ) {
							var cfaAlpfaColumn = losRow.Find( c => c.Type == LosExcelColumnType.CfaAlpha );
							var losTextColumn = losRow.Find( c => c.Type == LosExcelColumnType.LosText );

							outcomes.Add( new {
								text = losTextColumn != null ? losTextColumn.Value : "",
								letter = cfaAlpfaColumn != null ? cfaAlpfaColumn.Value : ""
							} );
						}

						verticalNode.SetAttribute( "outcome_statements", JsonConvert.SerializeObject( outcomes ) );

					}

					sequentialNode.AppendChild( verticalNode );
				}


				var conceptNameColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ConceptName );
				var conceptIdColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ConceptId );

				skip = (bandContainerNode != null && bandContainerNode.GetAttribute( "display_name" ) != bandColumn.Value) ? false : skip;
				if ( !skip ) {
					var bandIdColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.BandId );
					bandContainerNode = xml.CreateElement( "container" );
					bandContainerNode.SetAttribute( "url_name", getGuid( bandIdColumn.Value, CourseTypes.Band ) );
					bandContainerNode.SetAttribute( "display_name", bandColumn != null && bandColumn.HaveValue() ? bandColumn.Value : "" );
					bandContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					bandContainerNode.SetAttribute( "container_description", "" );
					bandContainerNode.SetAttribute( "learning_objective_id", "" );

					string targetScore = "";
					if ( acceptanceCriteriaExcel != null ) {
						var acceptanceCriteriaRow =
							acceptanceCriteriaExcel.Rows.FirstOrDefault(
								r => r.Any( c => c.Type == AcceptanceCriteriaColumnType.Lo1 && c.Value == conceptIdColumn.Value ) );

						if ( acceptanceCriteriaRow != null ) {
							var scoreColumn = acceptanceCriteriaRow.FirstOrDefault( c => c.Type == AcceptanceCriteriaColumnType.TargetScore );
							if ( scoreColumn != null && scoreColumn.HaveValue() ) {
								targetScore = scoreColumn.Value;
							}
						}
					}

					bandContainerNode.SetAttribute( "acceptance_criteria", targetScore );

					verticalNode.AppendChild( bandContainerNode );
				}

				skip = (conceptNameContainerNode != null && conceptNameContainerNode.GetAttribute( "display_name" ) != conceptNameColumn.Value) ? false : skip;
				if ( !skip ) {
					conceptNameContainerNode = xml.CreateElement( "container" );
					conceptNameContainerNode.SetAttribute( "url_name", getGuid( conceptIdColumn.Value, CourseTypes.Concept ) );
					conceptNameContainerNode.SetAttribute( "display_name", conceptNameColumn != null && conceptNameColumn.HaveValue() ? conceptNameColumn.Value : "" );
					conceptNameContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					conceptNameContainerNode.SetAttribute( "container_description", "" );
					conceptNameContainerNode.SetAttribute( "learning_objective_id", conceptIdColumn != null && conceptIdColumn.HaveValue() ? conceptIdColumn.Value : "" );
					bandContainerNode.AppendChild( conceptNameContainerNode );
				}

				//VIDEO
				if ( atomType.Value == "IN" ) {
					var atomTitleColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomTitle );
					var itemIdCoclumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ItemId );
					var videoNode = xml.CreateElement( "brightcove-video" );
					videoNode.SetAttribute( "url_name", getGuid( atomIdColumn.Value, CourseTypes.Video ) );
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
					var questionIdColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.QuestionId );
					var questionColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Question );
					var kkEeColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.KKEE );
					var answerImageUrlColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.AnswerImageUrl );
					var questionImageUrlColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.QuestionImageUrl );
					string questionValue = questionColumn.HaveValue() ? questionColumn.Value : questionImageUrlColumn.HaveValue() ? "" : "Question Missing";

					problemBuilderNode.SetAttribute( "display_name", questionValue.Replace( "<br>", "" ) );
					problemBuilderNode.SetAttribute( "url_name", getGuid( atomIdColumn.Value, CourseTypes.Question ) );
					problemBuilderNode.SetAttribute( "xblock-family", "xblock.v1" );
					problemBuilderNode.SetAttribute( "cfa_type", "question" );
					problemBuilderNode.SetAttribute( "atom_id", atomIdColumn != null && atomIdColumn.HaveValue() ? atomIdColumn.Value : "" );
					problemBuilderNode.SetAttribute( "instruct_assessment", kkEeColumn != null && kkEeColumn.HaveValue() ? kkEeColumn.Value : "" );

					if ( answerImageUrlColumn != null && answerImageUrlColumn.HaveValue() ) {
						problemBuilderNode.SetAttribute( "answer_image", answerImageUrlColumn.Value );
					}

					conceptNameContainerNode.AppendChild( problemBuilderNode );

					var pbMcqNode = xml.CreateElement( "pb-mcq-block" );
					var correctColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Correct );

					var actualCorrectValues = new List<string>();

					if ( correctColumn != null && correctColumn.HaveValue() ) {
						var correctValues = correctColumn.Value.Split( ' ' );

						foreach ( var correctValue in correctValues ) {
							actualCorrectValues.Add( questionDic[correctValue] );
						}
					}

					pbMcqNode.SetAttribute( "url_name", getNewGuid() );
					pbMcqNode.SetAttribute( "xblock-family", "xblock.v1" );
					pbMcqNode.SetAttribute( "question", questionValue );
					pbMcqNode.SetAttribute( "fitch_question_id", questionIdColumn.Value );
					pbMcqNode.SetAttribute( "correct_choices", (correctColumn != null && correctColumn.Value != null) ? JsonConvert.SerializeObject( actualCorrectValues ) : "" );


					if ( questionImageUrlColumn != null && questionImageUrlColumn.HaveValue() ) {
						pbMcqNode.SetAttribute( "image", questionImageUrlColumn.Value );
					}

					problemBuilderNode.AppendChild( pbMcqNode );

					var questionIds = new List<string>();

					var answer1Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer1 );
					var question1Id = questionDic["A"];
					var answer1Node = GetAnswerNode( xml, answer1Column, question1Id, false );
					if ( answer1Node != null ) {
						pbMcqNode.AppendChild( answer1Node );
						questionIds.Add( question1Id );
					}

					var answer2Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer2 );
					var question2Id = questionDic["B"];
					var answer2Node = GetAnswerNode( xml, answer2Column, question2Id, true );
					if ( answer2Node != null ) {
						pbMcqNode.AppendChild( answer2Node );
						questionIds.Add( question2Id );
					}

					var answer3Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer3 );
					var question3Id = questionDic["C"];
					var answer3Node = GetAnswerNode( xml, answer3Column, question3Id, true );
					if ( answer3Node != null ) {
						pbMcqNode.AppendChild( answer3Node );
						questionIds.Add( question3Id );
					}

					var answer4Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer4 );
					var question4Id = questionDic["D"];
					var answer4Node = GetAnswerNode( xml, answer4Column, question4Id );
					if ( answer4Node != null ) {
						pbMcqNode.AppendChild( answer4Node );
						questionIds.Add( question4Id );
					}


					//tip  block
					var justificationCell = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Justification );
					if ( justificationCell != null && justificationCell.HaveValue() ) {
						var questionTipNode = xml.CreateElement( "pb-tip-block" );
						questionTipNode.SetAttribute( "url_name", getNewGuid() );
						questionTipNode.SetAttribute( "xblock-family", "xblock.v1" );
						questionTipNode.SetAttribute( "values", JsonConvert.SerializeObject( questionIds ) );
						questionTipNode.InnerText = justificationCell.Value;
						pbMcqNode.AppendChild( questionTipNode );
					}

				}

				skip = true;

			}

			if ( locked != "yes" ) {
				AppendStudySessionTestQuestions( xml, PreviousSequentialNode, previousStudySessionId.Value, ssTestExcel );
			}

			XmlElement wikiNode = xml.CreateElement( "wiki" );
			wikiNode.SetAttribute( "slug", "test.1.2015" );
			courseNode.AppendChild( wikiNode );

			return xml;
		}

		private string getGuid( string elementId, CourseTypes elementType )
		{
			string key = Database.Instance.GetKey( elementId, elementType );
			if ( !String.IsNullOrEmpty( key ) ) {
				return key;
			}

			key = getNewGuid();
			Database.Instance.AddKey( elementId, key, elementType );
			Program.Log.Info( String.Format( "New Key generated elementType: {0}; element_id: {1}; generatedKey: {2}", elementType.ToString(), elementId, key ) );
			return key;
		}

		private string getNewGuid()
		{
			return Guid.NewGuid().ToString().Replace( "-", "" );
		}


		private Dictionary<string, string> generateQuestionIds()
		{
			var dic = new Dictionary<string, string>();

			while ( !dic.ContainsKey( "D" ) ) {
				string guid = getNewGuid();
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

		private XmlElement GetAnswerNode( XmlDocument xml, IExcelColumn<QuestionExcelColumnType> answerColumn, string questionId, bool addMissingValue = false )
		{
			return GetAnswerNode( xml, answerColumn != null && answerColumn.HaveValue() ? answerColumn.Value : "", questionId, addMissingValue );
		}

		private XmlElement GetAnswerNode( XmlDocument xml, IExcelColumn<SsTestExcelColumnType> answerColumn, string questionId, bool addMissingValue = false )
		{
			return GetAnswerNode( xml, answerColumn != null && answerColumn.HaveValue() ? answerColumn.Value : "", questionId, addMissingValue );
		}

		private XmlElement GetAnswerNode( XmlDocument xml, string answer, string questionId, bool addMissingValue = false )
		{
			if ( !String.IsNullOrWhiteSpace( answer ) || addMissingValue ) {
				var answerNode = xml.CreateElement( "pb-choice-block" );
				answerNode.SetAttribute( "url_name", getNewGuid() );
				answerNode.SetAttribute( "xblock-family", "xblock.v1" );
				answerNode.SetAttribute( "value", questionId );
				answerNode.InnerText = String.IsNullOrWhiteSpace( answer ) ? "Answer Missing" : answer.Replace( "/n", "" );
				return answerNode;
			}

			return null;
		}


		private void AppendStudySessionTestQuestions( XmlDocument xml, XmlElement sequentialNode, string studySessionId, Excel<SsTestExcelColumn, SsTestExcelColumnType> ssTestExcel )
		{
			var ssRows = ssTestExcel.Rows.Where( r => r.Any( c => c.Type == SsTestExcelColumnType.StudySessionId && c.Value == studySessionId ) );

			if ( ssRows.Any() ) {

				string verticalTestId = ssRows.First().FirstOrDefault( c => c.Type == SsTestExcelColumnType.KStructure ).Value;

				var verticalTestNode = xml.CreateElement( "vertical" );
				verticalTestNode.SetAttribute( "display_name", "Study Session Test" );
				verticalTestNode.SetAttribute( "cfa_type", "test" );
				verticalTestNode.SetAttribute( "cfa_short_name", "SST" );
				verticalTestNode.SetAttribute( "study_session_test_id", verticalTestId );
				verticalTestNode.SetAttribute( "url_name", getGuid( verticalTestId, CourseTypes.Reading ) );

				foreach ( var ssRow in ssRows ) {
					var questionDic = generateQuestionIds();
					var problemBuilderNode = xml.CreateElement( "problem-builder-block" );
					string questionId = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.QuestionId ).Value;
					var questionType = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.QuestionType );
					var questionColumn = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Question );
					var questionIdColumn = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.QuestionId );
					var questionImageUrlColumn = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.QuestionImageUrl );
					string questionValue = questionColumn.HaveValue() ? questionColumn.Value : questionImageUrlColumn.HaveValue() ? "" : "Question Missing";

					problemBuilderNode.SetAttribute( "display_name", questionValue.Replace( "<br>", "" ) );
					problemBuilderNode.SetAttribute( "url_name", getGuid( questionId, CourseTypes.Question ) );
					problemBuilderNode.SetAttribute( "xblock-family", "xblock.v1" );
					problemBuilderNode.SetAttribute( "cfa_type", "question" );
					problemBuilderNode.SetAttribute( "atom_id", questionId );
					problemBuilderNode.SetAttribute( "instruct_assessment", questionType != null && questionType.HaveValue() ? questionType.Value : "" );

					var answerImageUrlColumn = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.AnswerImageUrl );
					if ( answerImageUrlColumn != null && answerImageUrlColumn.HaveValue() ) {
						problemBuilderNode.SetAttribute( "answer_image", answerImageUrlColumn.Value );
					}

					var pbMcqNode = xml.CreateElement( "pb-mcq-block" );
					var correctColumn = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Correct );

					var actualCorrectValues = new List<string>();

					if ( correctColumn != null && correctColumn.HaveValue() ) {
						var correctValues = correctColumn.Value.Split( ' ' );

						foreach ( var correctValue in correctValues ) {
							actualCorrectValues.Add( questionDic[correctValue] );
						}
					}

					pbMcqNode.SetAttribute( "url_name", getNewGuid() );
					pbMcqNode.SetAttribute( "xblock-family", "xblock.v1" );
					pbMcqNode.SetAttribute( "question", questionValue );
					pbMcqNode.SetAttribute( "fitch_question_id", questionIdColumn.Value );
					pbMcqNode.SetAttribute( "correct_choices", (correctColumn != null && correctColumn.Value != null) ? JsonConvert.SerializeObject( actualCorrectValues ) : "" );


					if ( questionImageUrlColumn != null && questionImageUrlColumn.HaveValue() ) {
						pbMcqNode.SetAttribute( "image", questionImageUrlColumn.Value );
					}

					problemBuilderNode.AppendChild( pbMcqNode );

					var questionIds = new List<string>();

					var answer1Column = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Answer1 );
					var question1Id = questionDic["A"];
					var answer1Node = GetAnswerNode( xml, answer1Column, question1Id, false );
					if ( answer1Node != null ) {
						pbMcqNode.AppendChild( answer1Node );
						questionIds.Add( question1Id );
					}

					var answer2Column = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Answer2 );
					var question2Id = questionDic["B"];
					var answer2Node = GetAnswerNode( xml, answer2Column, question2Id, true );
					if ( answer2Node != null ) {
						pbMcqNode.AppendChild( answer2Node );
						questionIds.Add( question2Id );
					}

					var answer3Column = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Answer3 );
					var question3Id = questionDic["C"];
					var answer3Node = GetAnswerNode( xml, answer3Column, question3Id, true );
					if ( answer3Node != null ) {
						pbMcqNode.AppendChild( answer3Node );
						questionIds.Add( question3Id );
					}

					var answer4Column = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Answer4 );
					var question4Id = questionDic["D"];
					var answer4Node = GetAnswerNode( xml, answer4Column, question4Id );
					if ( answer4Node != null ) {
						pbMcqNode.AppendChild( answer4Node );
						questionIds.Add( question4Id );
					}


					//tip  block
					var justificationCell = ssRow.FirstOrDefault( c => c.Type == SsTestExcelColumnType.Justification );
					if ( justificationCell != null && justificationCell.HaveValue() ) {
						var questionTipNode = xml.CreateElement( "pb-tip-block" );
						questionTipNode.SetAttribute( "url_name", getNewGuid() );
						questionTipNode.SetAttribute( "xblock-family", "xblock.v1" );
						questionTipNode.SetAttribute( "values", JsonConvert.SerializeObject( questionIds ) );
						questionTipNode.InnerText = justificationCell.Value;
						pbMcqNode.AppendChild( questionTipNode );
					}

					verticalTestNode.AppendChild( problemBuilderNode );
				}

				sequentialNode.AppendChild( verticalTestNode );
			}


		}
	}
}
