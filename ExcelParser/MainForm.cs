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
			if ( openFileDialog.ShowDialog() == DialogResult.OK ) {
				var fileName = openFileDialog.FileName;
				var excel = Excel.ReadExcell( fileName );

                var xmlCourseConverter = new XmlCourseConverter();
                XmlDocument xml = xmlCourseConverter.ConvertExcelToXml( excel );
				SaveFileDialog saveFileDialog = new SaveFileDialog();
				saveFileDialog.FileName = "output.xml";
				if ( saveFileDialog.ShowDialog() == DialogResult.OK ) {
					xml.Save( saveFileDialog.FileName );
				}
			}

		}
	}
}
