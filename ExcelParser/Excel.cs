using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Xml;
using Excel;

namespace ExcelParser
{
	public class Excel
	{
		private Excel()
		{
			Header = new List<ExcelColumn>();
			Rows = new List<List<ExcelColumn>>();
		}

		public int NoRows { get; set; }
		public List<ExcelColumn> Header { get; set; }
		public List<List<ExcelColumn>> Rows { get; set; }

		public static Excel ReadExcell( string filePath )
		{
			var excel = new Excel();

			FileStream stream = File.Open( filePath, FileMode.Open, FileAccess.Read );

			IExcelDataReader excelReader = ExcelReaderFactory.CreateBinaryReader( stream );

			excelReader.IsFirstRowAsColumnNames = true;

			while ( excelReader.Read() ) {

				bool isHeader = !excel.Header.Any();
				if ( isHeader ) {
					bool haveColumns = true;
					int columnIndex = 0;
					while ( haveColumns ) {
						try {
							string columnName = excelReader.GetString( columnIndex );
							ExcelColumn column = ExcelColumn.GetColumn( columnName, columnIndex );

							if ( column.Type != ColumnType.Undefined ) {
								excel.Header.Add( column );
							}

							columnIndex++;
						}
						catch ( IndexOutOfRangeException exc ) {
							haveColumns = false;
						}
					}
				}
				else
				{
					var rowValues = new List<ExcelColumn>(); 
					foreach (var headerColumn in excel.Header)
					{
						string rowValue = excelReader.GetString( headerColumn.ColumnIndex );
						rowValues.Add(new ExcelColumn(rowValue, headerColumn.Type, headerColumn.ColumnIndex));
					}

					excel.Rows.Add(rowValues);
				}


			}

			//6. Free resources (IExcelDataReader is IDisposable)
			excelReader.Close();

			return excel;
		}

	}
}
