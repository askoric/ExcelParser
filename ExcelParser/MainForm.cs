using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Data;
using System.Drawing;
using System.IO;
using System.Linq;
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
		public MainForm()
		{
			InitializeComponent();
			StatusLabel.Text = "Waiting for user interaction.";

		}

		private void openFileDialog1_FileOk( object sender, CancelEventArgs e )
		{

		}

		private void button1_Click( object sender, EventArgs e )
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			Excel mainStructureExcel = null;
			Excel questionsExcel = null;
			Excel losExcel = null;
			Excel acceptanceCriteriaExcel = null;

			if ( openFileDialog.ShowDialog() == DialogResult.OK )
			{
				StatusLabel.Text = "reading excel";
				mainStructureExcel = Excel.ReadExcell( openFileDialog.FileName, XmlValueParser.Instance );
			}

			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				questionsExcel = Excel.ReadExcell( openFileDialog.FileName, XmlValueParser.Instance );
			}

			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				losExcel = Excel.ReadExcell( openFileDialog.FileName, XmlValueParser.Instance );
			}

			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				acceptanceCriteriaExcel = Excel.ReadExcell( openFileDialog.FileName, XmlValueParser.Instance );
			}

			var excelParser = new ExcelParser();

			if (SetTranscript.Checked)
			{
				StatusLabel.Text = "Getting video id's from Brightcove API";
				var xmlTranscriptAccesor = new XmlTranscriptAccessor();
				var videoReferenceIds = excelParser.GetVideoReferenceIds( mainStructureExcel );
				var videoReferencesWithoutTranscript = videoReferenceIds.Where(vr => !xmlTranscriptAccesor.TranscriptXmlExists(vr));
				var brightcoveResponse = BrightcoveService.GetVideoIdFromReferenceId( videoReferencesWithoutTranscript );

				StatusLabel.Text = "Getting video Transcript XML from 3PlayMedia API";
				_3playmediaService.GetTranscriptsXmlForVideo( brightcoveResponse.items );
			}

			StatusLabel.Text = "Generating output XML";
			XmlDocument xml = excelParser.ConvertExcelToCourseXml( mainStructureExcel, questionsExcel, losExcel, acceptanceCriteriaExcel, SetTranscript.Checked );
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = "output.xml";
			if ( saveFileDialog.ShowDialog() == DialogResult.OK ) {
				xml.Save( saveFileDialog.FileName );
			}

			StatusLabel.Text = "All Done";

		}

	}
}
