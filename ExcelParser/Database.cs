using System;
using System.CodeDom;
using System.Collections.Generic;
using System.Data.SQLite;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
	public class Database
	{
		private SQLiteConnection conn;
		private static Database instance;

		private Database()
		{
			conn = new SQLiteConnection( "Data Source=Database.sqlite;Version=3;" );
			conn.Open();

			string sql = "CREATE TABLE IF NOT EXISTS GeneratedIds (element_id VARCHAR(150), type VARCHAR(20), generated_id VARCHAR(20))";

			SQLiteCommand command = new SQLiteCommand( sql, conn );
			command.ExecuteNonQuery();
		}

		public static Database Instance
		{
			get
			{
				if ( instance == null ) {
					instance = new Database();
				}
				return instance;
			}
		}


		public string GetKey( string element_id, CourseTypes type )
		{
			if ( String.IsNullOrEmpty( element_id ) ) {
				return null;
			}

			string sql = String.Format( @"SELECT generated_id FROM GeneratedIds WHERE element_id = '{0}' AND type = '{1}'", element_id, type.ToString() );

			var command = new SQLiteCommand( sql, conn );
			using ( var dr = command.ExecuteReader() ) {
				if ( dr.Read() && !dr.IsDBNull( 0 ) ) {
					return dr.GetString( 0 );
				}
			}

			return null;
		}

		public void AddKeyIfDoesntExists( string element_id, string generatedId, CourseTypes elementType )
		{
			if ( String.IsNullOrEmpty( element_id ) || String.IsNullOrEmpty( generatedId ) ) {
				return;
			}

			if ( GetKey( element_id, elementType ) == null ) {
				AddKey( element_id, generatedId, elementType );
			}
		}

		public void AddKey( string element_id, string generatedId, CourseTypes elementType )
		{
			if ( String.IsNullOrEmpty( element_id ) || String.IsNullOrEmpty( generatedId ) ) {
				return;
			}

			string sql = String.Format( @"INSERT INTO GeneratedIds(element_id, type, generated_id) VALUES('{0}', '{1}', '{2}')", element_id, elementType.ToString(), generatedId );
			var command = new SQLiteCommand( sql, conn );
			command.ExecuteNonQuery();
		}


	}
}
