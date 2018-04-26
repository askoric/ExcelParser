using System;
using System.Collections.Generic;
using System.ComponentModel;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
	public interface IValueParser
	{
		string ParseValue( string value );
	}

	class XmlValueParser : IValueParser
	{
		public static XmlValueParser Instance = new XmlValueParser();

		public string ParseValue( string value )
		{
			if ( String.IsNullOrEmpty( value ) ) {
				return value;
			}

			return value.Replace( "", "<br>" );
		}
	}
}
