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
using System.Xml;
using Excel;

namespace ExcelParser
{
	public partial class MainForm : Form
	{
		public MainForm()
		{
			InitializeComponent();
		}

		private void openFileDialog1_FileOk( object sender, CancelEventArgs e )
		{

		}

		private void button1_Click( object sender, EventArgs e )
		{
			OpenFileDialog openFileDialog = new OpenFileDialog();
			Excel mainStructureExcel = null;
			Excel questionsExcel = null;
			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				mainStructureExcel = Excel.ReadExcell( openFileDialog.FileName, XmlValueParser.Instance );
			}

			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				questionsExcel = Excel.ReadExcell( openFileDialog.FileName, XmlValueParser.Instance );
			}


			var xmlCourseConverter = new XmlCourseConverter();
			XmlDocument xml = xmlCourseConverter.ConvertExcelToXml( mainStructureExcel, questionsExcel );
			SaveFileDialog saveFileDialog = new SaveFileDialog();
			saveFileDialog.FileName = "output.xml";
			if ( saveFileDialog.ShowDialog() == DialogResult.OK ) {
				xml.Save( saveFileDialog.FileName );
			}


		}

		private void MainForm_Load( object sender, EventArgs e )
		{

		}
	}
}
