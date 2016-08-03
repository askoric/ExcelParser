using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
	class XmlTranscriptAccessor
	{
		public string TranscriptFolderPath = String.Format( "{0}/VideoTranscripts", AppDomain.CurrentDomain.BaseDirectory );

		public XmlTranscriptAccessor()
		{
			Directory.CreateDirectory( TranscriptFolderPath );
		}

		public XmlDocument FindVideoTranscript( string videoReferenceId )
		{
			string filePath = String.Format("{0}/{1}.xml", TranscriptFolderPath, videoReferenceId);
			if (this.TranscriptXmlExists(videoReferenceId))
			{
				XmlDocument xmlDocument = new XmlDocument();
				xmlDocument.Load(filePath);
				return xmlDocument;
			}
			else
			{
				Program.Log.Info( String.Format( "Missing trenascript file: {0}", filePath ) );
			}

			return null;
		}

		public bool TranscriptXmlExists(string videoReferenceId)
		{
			return File.Exists( String.Format( "{0}/{1}.xml", TranscriptFolderPath, videoReferenceId ) );
		}

		public void SaveVideTranscript( XmlDocument transcriptXml, string videoReferenceId )
		{
			string xmlPath = String.Format( "{0}/{1}.xml", TranscriptFolderPath, videoReferenceId );
			transcriptXml.Save( xmlPath );
		}
	}
}
