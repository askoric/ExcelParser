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
		private Excel MainStructureExcel { get; set; }
		private Excel QuestionsExcel { get; set; }
		private Excel LosExcel { get; set; }
		private Excel AcceptanceCriteriaExcel { get; set; }

		private OpenFileDialog OpenFileDialog { get; set; }

		public MainForm()
		{
			OpenFileDialog = new OpenFileDialog();
			InitializeComponent();
			StatusLabel.Text = "Waiting for user interaction.";

		}

		private void openFileDialog1_FileOk( object sender, CancelEventArgs e )
		{

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
				MainStructureExcel = Excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				MainStructureExcelCheckImg.Visible = true;
			}
		}
		private void UploadQuestionsExcelBtn_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				QuestionsExcel = Excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				QuestionExcelCheckImg.Visible = true;
			}
		}

		private void UploadLOSExcelBtn_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				LosExcel = Excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				LosExcelCheckImg.Visible = true;
			}
		}


		private void UploadAcceptanceCriteriaExcel_Click( object sender, EventArgs e )
		{
			if ( OpenFileDialog.ShowDialog() == DialogResult.OK ) {
				AcceptanceCriteriaExcel = Excel.ReadExcell( OpenFileDialog.FileName, XmlValueParser.Instance );
				AcceptanceCriteriaCheckImg.Visible = true;
			}
		}

		private void GenerateCourseXmlBtn_Click( object sender, EventArgs e )
		{
			string missingExcels = String.Format("{0} {1} {2} {3}", MainStructureExcel == null ? "Main Structure Excel ," : "",
				AcceptanceCriteriaExcel == null ? "Acceptance Criteria Excel ," : "", LosExcel == null ? "Los Excel ," : "", QuestionsExcel == null ? "Question Excel ," : "" );

			if ( !String.IsNullOrWhiteSpace( missingExcels ) )
			{
				MessageBox.Show( String.Format( "You need to upload {0} in order to generate course XML", missingExcels.Remove( missingExcels.Length - 2 ) ) );
				return;
			}

				
			var excelParser = new ExcelParser();

			if (SetTranscript.Checked)
			{
				StatusLabel.Text = "Getting video id's from Brightcove API";
				var xmlTranscriptAccesor = new XmlTranscriptAccessor();
				var videoReferenceIds = excelParser.GetVideoReferenceIds( MainStructureExcel );
				var videoReferencesWithoutTranscript = videoReferenceIds.Where(vr => !xmlTranscriptAccesor.TranscriptXmlExists(vr));
				var brightcoveResponse = BrightcoveService.GetVideoIdFromReferenceId( videoReferencesWithoutTranscript );

				StatusLabel.Text = "Getting video Transcript XML from 3PlayMedia API";
				_3playmediaService.GetTranscriptsXmlForVideo( brightcoveResponse.items );
			}

			StatusLabel.Text = "Generating output XML";
			XmlDocument xml = excelParser.ConvertExcelToCourseXml( MainStructureExcel, QuestionsExcel, LosExcel, AcceptanceCriteriaExcel, SetTranscript.Checked );
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = "output.xml";
			if ( saveFileDialog.ShowDialog() == DialogResult.OK ) {
				xml.Save( saveFileDialog.FileName );
			}

			StatusLabel.Text = "All Done";
		}

	}
}
