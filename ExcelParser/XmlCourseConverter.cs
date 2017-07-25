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
	public class guidRequest
	{
		public string ElementId { get; set; }
		public CourseTypes elementType { get; set; }
	}

	public class ExcelParser
	{
		public XmlDocument ConvertExcelToCourseXml( Excel<MainStructureExcelColumn, MainStructureColumnType> mainStructureExcel, Excel<QuestionExcelColumn, QuestionExcelColumnType> questionExcel, Excel<LosExcelColumn, LosExcelColumnType> losExcel, Excel<AcceptanceCriteriaExcelColumn, AcceptanceCriteriaColumnType> acceptanceCriteriaExcel, Excel<TestExcelColumn, TestExcelColumnType> ssTestExcel, Excel<TestExcelColumn, TestExcelColumnType> progressTestExcel, Excel<TestExcelColumn, TestExcelColumnType> MockExamExcel, Excel<TestExcelColumn, TestExcelColumnType> FinalMockExamExcel, Excel<TestExcelColumn, TestExcelColumnType> TopicWorkshopExcel, bool setTranscripts )
		{
            CourseConverterHelper.generatedQuestionIds = new Dictionary<string, string>();
            CourseConverterHelper._generatedGuids = new Dictionary<string, guidRequest>();

			XmlDocument xml = new XmlDocument();
			var xmlTranscriptAccessor = new XmlTranscriptAccessor();

			XmlElement rootNode = xml.CreateElement( "xbundle" );
			xml.AppendChild( rootNode );

            // Getting Metadata from xml file
            XmlDocument doc = new XmlDocument();
            System.Reflection.Assembly a = System.Reflection.Assembly.GetExecutingAssembly();
            doc.Load(a.GetManifestResourceStream("ExcelParser.MetadataXml.xml"));
            XmlNode importNode = doc.DocumentElement.SelectSingleNode("/metadata");
            XmlNode metadataNode = rootNode.OwnerDocument.ImportNode(importNode, true);
            rootNode.AppendChild(metadataNode);

            XmlElement courseNode = xml.CreateElement( "course" );
			courseNode.SetAttribute( "advanced_modules", "[&quot;annotatable&quot;, &quot;videoalpha&quot;, &quot;openassessment&quot;, &quot;container&quot;, &quot;problem-builder-block&quot;, &quot;problem-builder-progress-test&quot;, &quot;problem-builder-mock-exam&quot;, &quot;textualatom&quot;]");
			courseNode.SetAttribute( "display_name", "CFA level 2 - version 81");
			courseNode.SetAttribute( "language", "en" );
			courseNode.SetAttribute( "start", "&quot;2016-01-01T00:00:00+00:00&quot;" );
			courseNode.SetAttribute( "org", "s" );
			courseNode.SetAttribute( "course", "s" );
			courseNode.SetAttribute( "url_name_orig", "course" );
			courseNode.SetAttribute( "semester", "course" );
            courseNode.SetAttribute( "days_before_review_unlock", "28");
            rootNode.AppendChild( courseNode );

            // Add Introduction Node
            courseNode.AppendChild(AddIntroductionTopic(xml));

            XmlElement chapterNode = null;
			XmlElement sequentialNode = null;
			XmlElement previousSequentialNode = null;
            XmlElement previousChapterNode = null;
            XmlElement verticalNode = null;
			XmlElement bandContainerNode = null;
			XmlElement conceptNameContainerNode = null;

            bool skip = false;

			IExcelColumn<MainStructureColumnType> previousStudySessionId = null;
            IExcelColumn<MainStructureColumnType> previousTopicId = null;
            string previousStudySessionLocked = "";
			string locked = "";

			foreach ( var row in mainStructureExcel.Rows ) {

				var structureTokens = row.First( c => c.Type == MainStructureColumnType.Structure ).Value.Split( '|' );

				//Band need to be legal value else skip this row
				var bandColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Band );
				if ( bandColumn == null || bandColumn.Value == null || bandColumn.Value[0].ToString().ToLower() != "b" ) {
					continue;
				}

				var atomType = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomType );
				if ( atomType == null || !atomType.HaveValue() || !(atomType.Value == "IN" || atomType.Value == "Q" || atomType.Value == "TXT") ) {
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
                var topicId = row.FirstOrDefault(c => c.Type == MainStructureColumnType.TopicShortName);
                if ( !skip ) {
					var topicShortName = row.FirstOrDefault( c => c.Type == MainStructureColumnType.TopicShortName );
					var examPercantage = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ExamPercentage );
                    var cfatopicweight = row.FirstOrDefault(c => c.Type == MainStructureColumnType.CfaTopicWeight);
                    var lockedColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Locked );
					var colorColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Color );
					var cfaTypeColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.CfaType );
					locked = lockedColumn != null && lockedColumn.HaveValue() ? lockedColumn.Value : "";

					var description = row.FirstOrDefault( c => c.Type == MainStructureColumnType.Description );
					chapterNode = xml.CreateElement( "chapter" );
					chapterNode.SetAttribute( "display_name", topicNameColumn != null && topicNameColumn.HaveValue() ? topicNameColumn.Value : "" );
					chapterNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( topicShortName.Value, CourseTypes.Topic ) );
					chapterNode.SetAttribute( "cfa_short_name", topicShortName != null && topicShortName.HaveValue() ? topicShortName.Value : "" );
					chapterNode.SetAttribute( "exam_percentage", examPercantage != null && examPercantage.HaveValue() ? examPercantage.Value : "" );
                    chapterNode.SetAttribute( "cfa_topic_weight", cfatopicweight != null && cfatopicweight.HaveValue() ? cfatopicweight.Value : "");
                    chapterNode.SetAttribute( "description", description != null && description.HaveValue() ? description.Value : "" );
					chapterNode.SetAttribute( "locked", locked );
					chapterNode.SetAttribute( "topic_color", colorColumn != null && colorColumn.HaveValue() ? colorColumn.Value : "" );
					chapterNode.SetAttribute( "cfa_type", cfaTypeColumn != null && cfaTypeColumn.HaveValue() ? cfaTypeColumn.Value : "topic" );
					chapterNode.SetAttribute( "taxon_id", String.Join( "|", structureTokens.Take( 2 ) ) );
					courseNode.AppendChild( chapterNode );

                    //ADD TOPIC WORKSHOP
                    if (TopicWorkshopExcel != null && previousChapterNode != null && previousTopicId != null)
                    {
                        AppendTopicWorkshop(xml, previousChapterNode, previousTopicId.Value, TopicWorkshopExcel);
                    }

                    previousChapterNode = chapterNode;
                    previousTopicId = topicId;

                }


				var sessionNameColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.SessionName );
				skip = (sequentialNode != null && sequentialNode.GetAttribute( "display_name" ) != sessionNameColumn.Value) ? false : skip;
				var studySessionId = row.FirstOrDefault( c => c.Type == MainStructureColumnType.StudySessionId );
				if ( !skip ) {
					var studySession = row.FirstOrDefault( c => c.Type == MainStructureColumnType.StudySession );
					sequentialNode = xml.CreateElement( "sequential" );
					sequentialNode.SetAttribute( "display_name",
						sessionNameColumn != null && sessionNameColumn.HaveValue() ? sessionNameColumn.Value : "" );
					sequentialNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( studySessionId.Value, CourseTypes.StudySession ) );
					sequentialNode.SetAttribute( "cfa_short_name",
						studySession != null && studySession.HaveValue() ? studySession.Value : "" );
					sequentialNode.SetAttribute( "taxon_id", String.Join( "|", structureTokens.Take( 3 ) ) );
					sequentialNode.SetAttribute( "proficiency_target", "70" );

					chapterNode.AppendChild( sequentialNode );

					//ADD TEST TO THE BOTTOM OF LAST SESSION NAME NODE
					if ( ssTestExcel != null && previousSequentialNode != null && previousStudySessionId != null && previousStudySessionLocked != "yes" ) {
                        AppendStudySessionTestQuestions( xml, previousSequentialNode, previousStudySessionId.Value, ssTestExcel );
					}

					previousSequentialNode = sequentialNode;
					previousStudySessionId = studySessionId;
					previousStudySessionLocked = locked;
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
					verticalNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( readingIdColumn.Value, CourseTypes.Reading ) );
					verticalNode.SetAttribute( "cfa_short_name", readingColumn != null && readingColumn.HaveValue() ? readingColumn.Value : "" );
					verticalNode.SetAttribute( "downloads", JsonConvert.SerializeObject( downloads ) );
					verticalNode.SetAttribute( "taxon_id", String.Join( "|", structureTokens.Take( 4 ) ) );
					verticalNode.SetAttribute( "proficiency_target", "70" );

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
					bandContainerNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( bandIdColumn.Value, CourseTypes.Band ) );
					bandContainerNode.SetAttribute( "display_name", bandColumn != null && bandColumn.HaveValue() ? bandColumn.Value : "" );
					bandContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					bandContainerNode.SetAttribute( "container_description", "" );
					bandContainerNode.SetAttribute( "learning_objective_id", "" );
					bandContainerNode.SetAttribute( "taxon_id", String.Join( "|", structureTokens.Take( 5 ) ) );

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

				skip = (conceptNameContainerNode != null && conceptNameContainerNode.GetAttribute( "learning_objective_id" ) != conceptIdColumn.Value) ? false : skip;
				if ( !skip ) {
					conceptNameContainerNode = xml.CreateElement( "container" );
					conceptNameContainerNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( conceptIdColumn.Value, CourseTypes.Concept ) );
					conceptNameContainerNode.SetAttribute( "display_name", conceptNameColumn != null && conceptNameColumn.HaveValue() ? conceptNameColumn.Value : "" );
					conceptNameContainerNode.SetAttribute( "xblock-family", "xblock.v1" );
					conceptNameContainerNode.SetAttribute( "container_description", "" );
					conceptNameContainerNode.SetAttribute( "learning_objective_id", conceptIdColumn != null && conceptIdColumn.HaveValue() ? conceptIdColumn.Value : "" );
					conceptNameContainerNode.SetAttribute( "taxon_id", String.Join( "|", structureTokens.Take( 6 ) ) );
					bandContainerNode.AppendChild( conceptNameContainerNode );
				}

				string atomTaxonId = String.Format( "{0}|{1}", String.Join( "|", structureTokens.Take( 6 ) ), atomIdColumn.Value );

				//VIDEO
				if ( atomType.Value == "IN" ) {
					var atomTitleColumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.AtomTitle );
					var itemIdCoclumn = row.FirstOrDefault( c => c.Type == MainStructureColumnType.ItemId );
					var videoNode = xml.CreateElement( "brightcove-video" );

					videoNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( atomIdColumn.Value, CourseTypes.Video ) );
					videoNode.SetAttribute( "xblock-family", "xblock.v1" );
					videoNode.SetAttribute( "api_bckey", "AQ~~,AAAELMh4AWE~,vVFFDlX6sNOap1Tww7YwaMvqbQ8TtDoh" );
					videoNode.SetAttribute( "display_name", atomTitleColumn != null && atomTitleColumn.HaveValue() ? atomTitleColumn.Value : "" );
					videoNode.SetAttribute( "api_key", "JqnRdhYvLWNtVJllXkMzGGGTh66uLLmz8JB8YlcZQlC8OX94H4ZXXw.." );
					videoNode.SetAttribute( "text_values", "[]" );
					videoNode.SetAttribute( "api_bctid", itemIdCoclumn.Value );
					videoNode.SetAttribute( "begin_values", "[]" );
					videoNode.SetAttribute( "api_bcpid", "4830051907001" );
					videoNode.SetAttribute( "cfa_type", "video" );
					videoNode.SetAttribute( "atom_id", atomIdColumn.Value );
					videoNode.SetAttribute( "taxon_id", atomTaxonId );

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

                //TEXT
                else if(atomType.Value == "TXT")
                {
                    var atomTitleColumn = row.FirstOrDefault(c => c.Type == MainStructureColumnType.AtomTitle);
                    var itemIdCoclumn = row.FirstOrDefault(c => c.Type == MainStructureColumnType.ItemId);
                    var atomText = row.FirstOrDefault(c => c.Type == MainStructureColumnType.AtomBody);
                    var textualNode = xml.CreateElement("textualatom");

                    textualNode.SetAttribute("url_name", CourseConverterHelper.getGuid(atomIdColumn.Value, CourseTypes.TextAtom));
                    textualNode.SetAttribute("xblock-family", "xblock.v1");
                    textualNode.SetAttribute("atom_text", atomText != null && atomText.HaveValue() ? atomText.Value : "");
                    textualNode.SetAttribute("display_name", atomTitleColumn != null && atomTitleColumn.HaveValue() ? atomTitleColumn.Value : "");
                    textualNode.SetAttribute("cfa_type", "text");
                    textualNode.SetAttribute("atom_id", atomIdColumn.Value);
                    textualNode.SetAttribute("taxon_id", atomTaxonId);

                    conceptNameContainerNode.AppendChild(textualNode);
                }

				//QUESTION
				else {
					var questionDic = CourseConverterHelper.generateQuestionIds();
					var problemBuilderNode = xml.CreateElement( "problem-builder-block" );
					var questionIdColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.QuestionId );
					var questionColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Question );
					var kkEeColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.KKEE );
					var answerImageUrlColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.AnswerImageUrl );
					var questionImageUrlColumn = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.QuestionImageUrl );
					string questionValue = questionColumn.HaveValue() ? questionColumn.Value : questionImageUrlColumn.HaveValue() ? "" : "Question Missing";

					problemBuilderNode.SetAttribute( "display_name", questionValue.Replace( "<br>", "" ) );
					problemBuilderNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( atomIdColumn.Value, CourseTypes.Question ) );
					problemBuilderNode.SetAttribute( "xblock-family", "xblock.v1" );
					problemBuilderNode.SetAttribute( "cfa_type", "question" );
					problemBuilderNode.SetAttribute( "atom_id", atomIdColumn != null && atomIdColumn.HaveValue() ? atomIdColumn.Value : "" );
					problemBuilderNode.SetAttribute( "instruct_assessment", kkEeColumn != null && kkEeColumn.HaveValue() ? kkEeColumn.Value : "" );
					problemBuilderNode.SetAttribute( "taxon_id", atomTaxonId );

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

					pbMcqNode.SetAttribute( "url_name", CourseConverterHelper.getNewGuid() );
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
					var answer1Node = CourseConverterHelper.GetAnswerNode( xml, answer1Column, question1Id, false );
					if ( answer1Node != null ) {
						pbMcqNode.AppendChild( answer1Node );
						questionIds.Add( question1Id );
					}

					var answer2Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer2 );
					var question2Id = questionDic["B"];
					var answer2Node = CourseConverterHelper.GetAnswerNode( xml, answer2Column, question2Id, true );
					if ( answer2Node != null ) {
						pbMcqNode.AppendChild( answer2Node );
						questionIds.Add( question2Id );
					}

					var answer3Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer3 );
					var question3Id = questionDic["C"];
					var answer3Node = CourseConverterHelper.GetAnswerNode( xml, answer3Column, question3Id, true );
					if ( answer3Node != null ) {
						pbMcqNode.AppendChild( answer3Node );
						questionIds.Add( question3Id );
					}

					var answer4Column = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Answer4 );
					var question4Id = questionDic["D"];
					var answer4Node = CourseConverterHelper.GetAnswerNode( xml, answer4Column, question4Id );
					if ( answer4Node != null ) {
						pbMcqNode.AppendChild( answer4Node );
						questionIds.Add( question4Id );
					}

					//tip  block
					var justificationCell = questionRow.FirstOrDefault( c => c.Type == QuestionExcelColumnType.Justification );
					string justificationInnerText = "";
					if ( answerImageUrlColumn != null && answerImageUrlColumn.HaveValue() ) {
						justificationInnerText = String.Format( "<div class='answer-image'><img src='{0}'></div>", answerImageUrlColumn.Value );
					}

					string justificationValue = (justificationCell != null && justificationCell.HaveValue())
						? justificationCell.Value
						: "";

					var questionTipNode = xml.CreateElement( "pb-tip-block" );
					questionTipNode.SetAttribute( "url_name", CourseConverterHelper.getNewGuid() );
					questionTipNode.SetAttribute( "xblock-family", "xblock.v1" );
					questionTipNode.SetAttribute( "values", JsonConvert.SerializeObject( questionIds ) );
					questionTipNode.InnerText = String.Format( "{0}{1}", justificationValue, justificationInnerText );
					pbMcqNode.AppendChild( questionTipNode );
				}

				skip = true;
			}

			if ( ssTestExcel != null && previousStudySessionLocked != "yes" ) {
                AppendStudySessionTestQuestions( xml, previousSequentialNode, previousStudySessionId.Value, ssTestExcel );
			}

            if ( progressTestExcel != null ) {
				var progressTestChapterNode = ProgressTestExcelConverter.Convert( xml, progressTestExcel );
                // Insert Progress Test after Equity
                var equityTopicNode = courseNode.SelectSingleNode("chapter[@display_name = 'Equity']");
                courseNode.InsertAfter(progressTestChapterNode, equityTopicNode);
            }

            if (MockExamExcel != null) {
                var mockExamChapterNodes = MockExamExcelConverter.Convert( xml, MockExamExcel );
                foreach(var mockExamChapterNode in mockExamChapterNodes)
                {
                    courseNode.AppendChild(mockExamChapterNode);
                }
            }

            if (FinalMockExamExcel != null)
            {
                var finalMockExamChapterNode = FinalMockExamExcelConverter.Convert(xml, FinalMockExamExcel);
                courseNode.AppendChild(finalMockExamChapterNode);
            }

            if (TopicWorkshopExcel != null)
            {
                AppendTopicWorkshop(xml, previousChapterNode, previousTopicId.Value, TopicWorkshopExcel);
            }

            // Add Revision and FinalExam Node
            courseNode.AppendChild(AddRevisionTopic(xml));
            courseNode.AppendChild(AddFinalExamTopic(xml));

            XmlElement wikiNode = xml.CreateElement( "wiki" );
			wikiNode.SetAttribute( "slug", "sf1.sf1.sf1");
			courseNode.AppendChild( wikiNode );

			return xml;
		}



		private void AppendStudySessionTestQuestions( XmlDocument xml, XmlElement sequentialNode, string studySessionId, Excel<TestExcelColumn, TestExcelColumnType> ssTestExcel )
		{
			var excelRows = ssTestExcel.Rows.Where( r => r.Any( c => c.Type == TestExcelColumnType.SessionAbbrevation && c.Value == studySessionId ) );

			if ( excelRows.Any() ) {

				string verticalTestId = excelRows.First().FirstOrDefault( c => c.Type == TestExcelColumnType.KStructure ).Value;
				verticalTestId = String.Join( "|", verticalTestId.Split( '|' ).Take( 3 ) );

				var verticalNode = xml.CreateElement( "vertical" );
				verticalNode.SetAttribute( "display_name", "Study Session Test" );
				verticalNode.SetAttribute( "cfa_type", "test" );
				verticalNode.SetAttribute( "cfa_short_name", "SST" );
				verticalNode.SetAttribute( "study_session_test_id", verticalTestId );
				verticalNode.SetAttribute( "url_name", CourseConverterHelper.getGuid( verticalTestId, CourseTypes.Reading ) );

                var containerReferences = excelRows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.ContainerRef).Value);
                foreach (var containerReference in containerReferences)
                {
                    string containerReferenceValue = containerReference.Key;
                    var ssRows = excelRows.Where(r => r.Any(c => c.Type == TestExcelColumnType.ContainerRef && c.Value.Contains(containerReferenceValue)));

                    if (ssRows.Any())
                    {
                        char index = containerReferenceValue.Last();
                        var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, ssRows, new ProblemBuilderNodeSettings
                        {
                            DisplayName = "Study Session Test",
                            UrlName = CourseConverterHelper.getGuid(verticalTestId + '_' + index, CourseTypes.Question),
                            ProblemBuilderNodeElement = "problem-builder-block",
                            PbMcqNodeElement = "pb-mcq-block",
                            PbChoiceBlockElement = "pb-choice-block",
                            PbTipBlockElement = "pb-tip-block"
                        });

                        verticalNode.AppendChild(problemBuilderNode);
                    }
                }

				sequentialNode.AppendChild( verticalNode );
			}
		}


        private void AppendTopicWorkshop(XmlDocument xml, XmlElement chapterNode, string topicId, Excel<TestExcelColumn, TestExcelColumnType> TopicWorkshopExcel)
        {
            var topicWorkshopRows = TopicWorkshopExcel.Rows.Where(r => r.Any(c => c.Type == TestExcelColumnType.TopicAbbrevation && c.Value == topicId));
            if (topicWorkshopRows.Any())
            {
                var workshopReferences = topicWorkshopRows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.TopicWorkshopReference).Value);
                foreach (var workshopReference in workshopReferences)
                {
                    string workshopReferenceValue = workshopReference.Key;
                    var workshopRows = topicWorkshopRows.Where(r => r.Any(c => c.Type == TestExcelColumnType.TopicWorkshopReference && c.Value.Contains(workshopReferenceValue)));

                    if (workshopRows.Any())
                    {
                        string topicWorkshopTitle = workshopRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicWorkshopTitle).Value;
                        string topicWorkshopType = workshopRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicWorkshopType).Value;

                        if (topicWorkshopType == "Topic Workshop")
                        {
                            topicWorkshopType = "topic_workshop";
                        }
                        else if (topicWorkshopType == "Review Course Workshop")
                        {
                            topicWorkshopType = "course_workshop";
                        }

                        var sequentialNode = xml.CreateElement("sequential");
                        sequentialNode.SetAttribute("display_name", topicWorkshopTitle);
                        sequentialNode.SetAttribute("url_name", CourseConverterHelper.getGuid(workshopReferenceValue, CourseTypes.Workshop));
                        sequentialNode.SetAttribute("workshop_id", workshopReferenceValue);
                        sequentialNode.SetAttribute("cfa_type", topicWorkshopType);

                        chapterNode.AppendChild(sequentialNode);

                        var itemSetReferences = workshopRows.GroupBy(r => r.First(tn => tn.Type == TestExcelColumnType.ItemSetReference).Value);
                        foreach (var itemSetReference in itemSetReferences)
                        {
                            string itemSetReferenceValue = itemSetReference.Key;
                            char index = itemSetReferenceValue.Last();
                            var itemSetRows = workshopRows.Where(r => r.Any(c => c.Type == TestExcelColumnType.ItemSetReference && c.Value.Contains(itemSetReferenceValue)));

                            if (itemSetRows.Any())
                            {
                                string itemSetTitle = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetTitle).Value;
                                string itemSetPdf = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.ItemSetPdf).Value;
                                string itemSetStudySessions = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.Session).Value;
                                string itemSetAnswerVideo = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.AnswerVideo).Value;
                                string vignetteTitle = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteTitle).Value;
                                string vignetteBody = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.VignetteBody).Value;
                                string topicTaxonId = itemSetRows.First().FirstOrDefault(c => c.Type == TestExcelColumnType.TopicTaxonId).Value;

                                var verticalNode = xml.CreateElement("vertical");
                                verticalNode.SetAttribute("cfa_type", "item_set");
                                verticalNode.SetAttribute("item_set_id", itemSetReferenceValue);
                                verticalNode.SetAttribute("url_name", CourseConverterHelper.getGuid(itemSetReferenceValue, CourseTypes.ItemSet));
                                verticalNode.SetAttribute("display_name", itemSetTitle);
                                verticalNode.SetAttribute("taxon_id", topicTaxonId);
                                verticalNode.SetAttribute("item_set_pdf", itemSetPdf);
                                verticalNode.SetAttribute("item_set_sessions", itemSetStudySessions);
                                verticalNode.SetAttribute("item_set_video", itemSetAnswerVideo);
                                verticalNode.SetAttribute("vignette_title", vignetteTitle);
                                verticalNode.SetAttribute("vignette_body", vignetteBody);

                                sequentialNode.AppendChild(verticalNode);

                                //skip first row(vignette row)
                                itemSetRows = itemSetRows.Skip(1);

                                var problemBuilderNode = ProblemBuilderNodeGenerator.Generate(xml, itemSetRows, new ProblemBuilderNodeSettings
                                {
                                    DisplayName = "Item Set " + index,
                                    UrlName = CourseConverterHelper.getGuid(itemSetReferenceValue, CourseTypes.Question),
                                    ProblemBuilderNodeElement = "problem-builder-block",
                                    PbMcqNodeElement = "pb-mcq-block",
                                    PbChoiceBlockElement = "pb-choice-block",
                                    PbTipBlockElement = "pb-tip-block"
                                });

                                verticalNode.AppendChild(problemBuilderNode);
                            }
                        }
                    }
                }
            }
        }

        public List<string> GetVideoReferenceIds(Excel<MainStructureExcelColumn, MainStructureColumnType> mainStructureExcel)
        {
            var referenceIds = new List<string>();
            var rowIndex = 0;
            foreach (var row in mainStructureExcel.Rows)
            {
                rowIndex++;
                var atomType = row.FirstOrDefault(c => c.Type == MainStructureColumnType.AtomType);
                if (atomType == null || !atomType.HaveValue())
                {
                    Program.Log.Info(String.Format("Missing atom type for row {0}", rowIndex));
                    continue;
                }

                if (atomType.Value == "IN")
                {
                    var atomIdColumn = row.FirstOrDefault(c => c.Type == MainStructureColumnType.AtomId);
                    if (atomIdColumn != null && atomIdColumn.HaveValue())
                    {
                        referenceIds.Add(atomIdColumn.Value);
                    }
                    else {
                        Program.Log.Info(String.Format("Missing atom Id for row {0}", rowIndex));
                    }
                }
            }
            return referenceIds;
        }

        public XmlElement AddIntroductionTopic(XmlDocument xml)
        {
            var introTopicNode = ChapterNodeGenerator.Generate(xml, new ChapterNodeGeneratorSettings
            {
                DisplayName = "Introduction",
                UrlName = "mrnwbbsgvab1y5faqfq7vv28e29yhsfk",
                CfaType = "intro",
                Description = "This is the introduction video.",
                Locked = "no"
            });
            return introTopicNode;
        }

        public XmlElement AddRevisionTopic(XmlDocument xml)
        {
            var revisionTopicNode = ChapterNodeGenerator.Generate(xml, new ChapterNodeGeneratorSettings
            {
                DisplayName = "Review",
                UrlName = "jqb8xtxn9ub55a7pnlys9e1zmyl5xf7u",
                CfaType = "revision",
                Locked = "no"
            });
            return revisionTopicNode;
        }

        public XmlElement AddFinalExamTopic(XmlDocument xml)
        {
            var finalExamTopicNode = ChapterNodeGenerator.Generate(xml, new ChapterNodeGeneratorSettings
            {
                DisplayName = "Final Exam",
                UrlName = "lysuv2kibu68r387yhboz5mfh5sxg3ve",
                CfaType = "final_exam",
                Locked = "no"
            });
            return finalExamTopicNode;
        }
    }
}
