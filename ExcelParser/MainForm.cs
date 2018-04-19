using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Forms;
using System.Windows.Forms.VisualStyles;
using System.Xml;
using Excel;

namespace ExcelParser
{
	public partial class MainForm : Form
	{
		private Excel<MainStructureExcelColumn, MainStructureColumnType> MainStructureExcel { get; set; }
		private Excel<QuestionExcelColumn, QuestionExcelColumnType> QuestionsExcel { get; set; }
		private Excel<LosExcelColumn, LosExcelColumnType> LosExcel { get; set; }
		private Excel<AcceptanceCriteriaExcelColumn, AcceptanceCriteriaColumnType> AcceptanceCriteriaExcel { get; set; }
		private Excel<TestExcelColumn, TestExcelColumnType> SsTestExcel { get; set; }
		private Excel<TestExcelColumn, TestExcelColumnType> ProgressTestExcel { get; set; }
		private Excel<MockExamExcelColumn, MockExamExcelColumnType> MockExamsExcel { get; set; }
        private Excel<TestExcelColumn, TestExcelColumnType> TopicWorkshopExcel { get; set; }

        private OpenFileDialog OpenFileDialog { get; set; }

		public MainForm()
		{
			OpenFileDialog = new OpenFileDialog();
			InitializeComponent();
			StatusLabel.Text = "Waiting for user interaction.";

		}

		private void openFileDialog1_FileOk( object sender, CancelEventArgs e )
		{
			var a = new LosExcelColumn();
		}

		private void button1_Click( object sender, EventArgs e )
		{

		}

		/// <summary>
		/// Fill DB from existing course XMLS
		/// </summary>
		/// <param name="sender"></param>
		/// <param name="e"></param>
		private void FillDbFromExistingCourseXml_Click( object sender, EventArgs e )
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			XmlDocument courseXml = new XmlDocument();
			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				StatusLabel.Text = "Importing DB from course XML";
				courseXml.Load( openFileDialog.FileName );
				XmlCourseParser.FillDbIdsFromCourseXml( courseXml );
				StatusLabel.Text = "Done importing DB.";
			}
		}

		private void UploadMainStructureExcelBtn_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				var excel = new Excel<MainStructureExcelColumn, MainStructureColumnType>();
				MainStructureExcel = excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				if ( MainStructureExcel.Header.Count() == Enum.GetNames( typeof( MainStructureColumnType ) ).Length - 1 ) {
					MainStructureExcelCheckImg.Visible = true;
				}
				else {
					MainStructureExcelCheckImg.Visible = false;
					MessageBox.Show( "Invalid excel. Excel does not have all required columns: TopicTitle, Topic, SessionTitle, ReadingTitle, Band, BandId, lo description, lo /concept id, item id, atom type, atom id, atom title, StudySession, StudySession abb, reading, exam percentage, description, downloads, downloads2, locked, color, type" );
					MainStructureExcel = null;
				}
			}
		}
		private void UploadQuestionsExcelBtn_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				var excel = new Excel<QuestionExcelColumn, QuestionExcelColumnType>();
				QuestionsExcel = excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				if ( QuestionsExcel.Header.Count() == Enum.GetNames( typeof( QuestionExcelColumnType ) ).Length - 1 ) {
					QuestionExcelCheckImg.Visible = true;
				}
				else {
					QuestionExcelCheckImg.Visible = false;
					MessageBox.Show( "Invalid excel. Excel does not have all required columns: id (original & new placeholder), Question, Question_Image_UR, Answer_Image_Url, Correct, Answer_1, Answer_2, Answer_3, Answer_4, Juristification, (kk/ee) instruct (k/e) asseess" );
					QuestionsExcel = null;
				}
			}
		}

		private void UploadLOSExcelBtn_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				var excel = new Excel<LosExcelColumn, LosExcelColumnType>();
				LosExcel = excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				if ( LosExcel.Header.Count() == Enum.GetNames( typeof( LosExcelColumnType ) ).Length - 1 ) {
					LosExcelCheckImg.Visible = true;
				}
				else {
					LosExcelCheckImg.Visible = false;
					MessageBox.Show( "Invalid excel. Excel does not have all required columns: TopicTitle, SessionTitle, ReadingTitle Cfa_Alpha, Los Text" );
					LosExcel = null;
				}

			}
		}


		private void UploadAcceptanceCriteriaExcel_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				var excel = new Excel<AcceptanceCriteriaExcelColumn, AcceptanceCriteriaColumnType>();
				AcceptanceCriteriaExcel = excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				if ( AcceptanceCriteriaExcel.Header.Count() == Enum.GetNames( typeof( AcceptanceCriteriaColumnType ) ).Length - 1 ) {
					AcceptanceCriteriaCheckImg.Visible = true;
				}
				else {
					AcceptanceCriteriaCheckImg.Visible = false;
					MessageBox.Show( "Invalid excel. Excel does not have all required columns: Lo1, Target Score" );
					AcceptanceCriteriaExcel = null;
				}
			}
		}

		private void UploadSSTestsExcel_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				var excel = new Excel<TestExcelColumn, TestExcelColumnType>();
				SsTestExcel = excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
                // -14 there is no FcmNumber and TopicTaxonId and PdfAnswers/Questions and workshop stuff(9) and essays(4)
                if ( SsTestExcel.Header.Count() == Enum.GetNames( typeof( TestExcelColumnType ) ).Length - 18 ) {
					UploadSsTestCheckImage.Visible = true;
				}
				else {
					UploadSsTestCheckImage.Visible = false;
					MessageBox.Show( "Invalid excel. Excel does not have all required columns: StudySessionId, KStructure, QuestionId, Question, QuestionType, Answer1, Answer2, Answer3, Answer4, AnswerImageUrl, Correct, QuestionImageUrl, Justification" );
					SsTestExcel = null;
				}
			}
		}

		private void UploadProgressTestExcell_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				var excel = new Excel<TestExcelColumn, TestExcelColumnType>();
				ProgressTestExcel = excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
                //-5 fcmNumber and ContainerRef and PdfAnswers/Questions and essays(2)
                if ((ProgressTestExcel.Header.Count() == Enum.GetNames(typeof(TestExcelColumnType)).Length - 5) || (ProgressTestExcel.Header.Count() == Enum.GetNames( typeof( TestExcelColumnType ) ).Length - 8) || ProgressTestExcel.Header.Count() == Enum.GetNames(typeof(TestExcelColumnType)).Length - 19 ) {
					uploadProgressTestCheckIcon.Visible = true;
				}
				else {
					uploadProgressTestCheckIcon.Visible = false;
					MessageBox.Show( "Invalid excel. Excel does not have all required columns: StudySessionId, KStructure, QuestionId, Question, Answer1, Answer2, Answer3, Answer4, AnswerImageUrl, Correct, QuestionImageUrl, Justification" );
					ProgressTestExcel = null;
				}
			}
		}

        private void MockExamBtn_Click(object sender, EventArgs e)
        {
            if (OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                var excel = new Excel<MockExamExcelColumn, MockExamExcelColumnType>();
                MockExamsExcel = excel.ReadExcell(OpenFileDialog.FileName, XmlValueParser.Instance);
                if ((MockExamsExcel.Header.Count() == Enum.GetNames(typeof(MockExamExcelColumnType)).Length - 1))
                {
                    uploadMockExamCheckIcon.Visible = true;
                }
                else {
                    uploadMockExamCheckIcon.Visible = false;
                    MessageBox.Show("Invalid excel. Excel does not have all required columns");
                    MockExamsExcel = null;
                }
            }
        }

        private void TopicWorkshopBtn_Click(object sender, EventArgs e)
        {
            if (OpenFileDialog.ShowDialog() == DialogResult.OK)
            {
                var excel = new Excel<TestExcelColumn, TestExcelColumnType>();
                TopicWorkshopExcel = excel.ReadExcell(OpenFileDialog.FileName, XmlValueParser.Instance);
                //-5 there is no fcmNumber and ContainerRef and PdfAnswers/ Questions and essays(3)
                if ((TopicWorkshopExcel.Header.Count() == Enum.GetNames(typeof(TestExcelColumnType)).Length - 7) || TopicWorkshopExcel.Header.Count() == Enum.GetNames(typeof(TestExcelColumnType)).Length - 8)
                {
                    uploadTopicWorkshopCheckIcon.Visible = true;
                }
                else
                {
                    uploadTopicWorkshopCheckIcon.Visible = false;
                    MessageBox.Show("Invalid excel. Excel does not have all required columns: StudySessionId, KStructure, QuestionId, Question, Answer1, Answer2, Answer3, Answer4, AnswerImageUrl, Correct, QuestionImageUrl, Justification");
                    TopicWorkshopExcel = null;
                }
            }
        }

        private void GenerateCourseXmlBtn_Click( object sender, EventArgs e )
		{
			string missingExcels = String.Format( "{0} {1} {2} {3}",
				MainStructureExcel == null ? "Main Structure Excel ," : "",
				AcceptanceCriteriaExcel == null ? "Acceptance Criteria Excel ," : "",
				LosExcel == null ? "Los Excel ," : "",
				QuestionsExcel == null ? "Question Excel ," : "" );

			if ( !String.IsNullOrWhiteSpace( missingExcels ) ) {
				MessageBox.Show( String.Format( "You need to upload {0} in order to generate course XML", missingExcels.Remove( missingExcels.Length - 2 ) ) );
				return;
			}

			var excelParser = new ExcelParser();

			if ( SetTranscript.Checked ) {
				StatusLabel.Text = "Getting video id's from Brightcove API";
				var xmlTranscriptAccesor = new XmlTranscriptAccessor();
				var videoReferenceIds = excelParser.GetVideoReferenceIds( MainStructureExcel );
				var videoReferencesWithoutTranscript = videoReferenceIds.Where( vr => !xmlTranscriptAccesor.TranscriptXmlExists( vr ) );
				var brightcoveResponse = BrightcoveService.GetVideoIdFromReferenceId( videoReferencesWithoutTranscript );

				StatusLabel.Text = "Getting video Transcript XML from 3PlayMedia API";
				_3playmediaService.GetTranscriptsXmlForVideo( brightcoveResponse.items );
			}

			StatusLabel.Text = "Generating output XML";
			XmlDocument xml = excelParser.ConvertExcelToCourseXml( MainStructureExcel, QuestionsExcel, LosExcel, AcceptanceCriteriaExcel, SsTestExcel, ProgressTestExcel, MockExamsExcel, TopicWorkshopExcel, SetTranscript.Checked );
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = "output.xml";
			if ( saveFileDialog.ShowDialog() == DialogResult.OK ) {
				xml.Save( saveFileDialog.FileName );
			}

			StatusLabel.Text = "All Done";
		}
    }
}
