using System;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Threading.Tasks;
using System.Windows.Forms;

namespace ExcelParser
{
	static class Program
	{
		public static log4net.ILog Log = log4net.LogManager.GetLogger( System.Reflection.MethodBase.GetCurrentMethod().DeclaringType );


		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main()
		{
			Log.Info("-------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------------");
			Log.Info( String.Format( "Program have starderd {0}", DateTime.Now ) );

			Application.EnableVisualStyles();
			Application.SetCompatibleTextRenderingDefault( false );
			Application.Run( new MainForm() );
		}
	}
}
