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

	public class Excel<T, TCtype> where T : IExcelColumn<TCtype>, new()
	{
		public Excel()
		{
			Header = new List<IExcelColumn<TCtype>>();
			Rows = new List<List<IExcelColumn<TCtype>>>();
		}

		public int NoRows { get; set; }
		public List<IExcelColumn<TCtype>> Header { get; set; }
		public List<List<IExcelColumn<TCtype>>> Rows { get; set; }

		public Excel<T, TCtype> ReadExcell( string filePath, IValueParser valueParser )
		{
			var excel = new Excel<T, TCtype>();

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

							var column = new T().GetColumn( columnName, columnIndex );

							if ( column.IsRecognizableColumn() ) {
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
					var rowValues = new List<IExcelColumn<TCtype>>(); 
					foreach (var headerColumn in excel.Header)
					{
						string rowValue = valueParser.ParseValue(excelReader.GetString( headerColumn.ColumnIndex ));
						rowValues.Add( new T
						{
							Value = rowValue,
							Type = headerColumn.Type,
							ColumnIndex = headerColumn.ColumnIndex
						});
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
