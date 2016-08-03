using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Http;
using System.Net.Http.Headers;
using System.Net.Mail;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading;
using System.Threading.Tasks;
using Excel.Log;

namespace ExcelParser
{
	public class BrightcoveResponse
	{
		public BrightcoveResponse()
		{
			items = new List<VideoIndentification>();
		}

		public List<VideoIndentification> items;
	}

	public class VideoIndentification
	{
		public string id { get; set; }
		public string referenceId { get; set; }
	}

	public class BrightcoveService
	{
		private const string BRIGHTCOVE_TOKEN = "wTnYeX2K01GUtv9UWBzT-kLWGEENreKujaHbRmFwAkluh9GKDEd6Gg..";
		public static BrightcoveResponse GetVideoIdFromReferenceId( IEnumerable<string> referenceIds )
		{
			var brightcoveResponse = new BrightcoveResponse();
			if (referenceIds == null || !referenceIds.Any())
			{
				Program.Log.Info( "No reference ids to get video id from brightcove." );
			}

			int take = 100;
			int skip = 0;
			while (referenceIds.Count() >= (take + skip))
			{
				var batchBrightcoveResponse = GetVideoIdFromReferenceIdBatch( referenceIds.Take( take ).Skip( skip ) );
				brightcoveResponse.items.AddRange( batchBrightcoveResponse.items );
				skip += 100;
				take += 100;

				//Access frequency should be less than 10 queries per second
				Thread.Sleep(200);
			}

			return brightcoveResponse;
		}

		private static BrightcoveResponse GetVideoIdFromReferenceIdBatch( IEnumerable<string> referenceIds )
		{
			var brightcoveResponse = new BrightcoveResponse();
			string referenceIdsString = String.Join(",", referenceIds);
			using (var client = new HttpClient())
			{
				client.BaseAddress = new Uri( "http://api.brightcove.com/services/" );
				client.DefaultRequestHeaders.Accept.Clear();
				client.DefaultRequestHeaders.Accept.Add( new MediaTypeWithQualityHeaderValue( "application/json" ) );

				// HTTP GET
				HttpResponseMessage response = Task.Run( () => client.GetAsync( String.Format( "library?command=find_videos_by_reference_ids&reference_ids={0}&video_fields=id,referenceId&token={1}", referenceIdsString, BRIGHTCOVE_TOKEN ) ) ).Result;
				if (response.IsSuccessStatusCode)
				{
					brightcoveResponse = response.Content.ReadAsAsync<BrightcoveResponse>().Result;
				}
				else
				{
					Program.Log.Info(String.Format("bad response from Brightcove API: status: {0}, reason phrase: {1}", response.StatusCode, response.ReasonPhrase));
				}


			}

			return brightcoveResponse;
		}
	}
}
