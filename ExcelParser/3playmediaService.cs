using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using System.Xml;

namespace ExcelParser
{
	public class _3playmediaService
	{
		private const string _3PLAY_TOKEN = "crKWdxlepc_c6eDdw9TV1XGwGTAZRL5c";

		public static void GetTranscriptsXmlForVideo( IEnumerable<VideoIndentification> videoIndentifications )
		{
			using ( var client = new HttpClient() ) {
				client.BaseAddress = new Uri( "http://static.3playmedia.com/files/" );
				client.DefaultRequestHeaders.Accept.Clear();
				client.DefaultRequestHeaders.Accept.Add( new MediaTypeWithQualityHeaderValue( "application/json" ) );
				var xmlTranscriptAccessor = new XmlTranscriptAccessor();

				foreach ( var videoIndentification in videoIndentifications ) {

					if (videoIndentification == null || String.IsNullOrEmpty(videoIndentification.id))
					{
						continue;
					}

					string url = String.Format( "http://static.3playmedia.com/files/{0}/transcript.pptxml?apikey={1}&usevideoid=1", videoIndentification.id, _3PLAY_TOKEN );
					XmlDocument xmlDoc = new XmlDocument();
					xmlDoc.Load( url );

					xmlTranscriptAccessor.SaveVideTranscript( xmlDoc, videoIndentification.referenceId );

					Thread.Sleep( 50 );

				}

			}
		}

	}
}
