using System;
using System.Collections.Generic;
using System.Linq;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
	public class ExcelColumn
	{

		public ExcelColumn( string value, ColumnType type, int index )
		{
			Value = value;
			Type = type;
			ColumnIndex = index;
		}

		public string Value { get; set; }
		public int ColumnIndex { get; set; }
		public ColumnType Type { get; set; }

		public bool HaveValue()
		{
			return Value != null && Value.Trim() != "";
		}

		public static ExcelColumn GetColumn( string columnName, int index )
		{
			switch ( columnName.ToLower().Trim() ) {
				case "topicname":
				case "topictitle":
					return new ExcelColumn( columnName, ColumnType.TopicName, index );
				case "sessionname":
				case "sessiontitle":
					return new ExcelColumn( columnName, ColumnType.SessionName, index );
				case "readingname":
				case "readingtitle":
					return new ExcelColumn( columnName, ColumnType.ReadingName, index );
				case "band":
					return new ExcelColumn( columnName, ColumnType.Band, index );
				case "conceptname":
				case "concept":
					return new ExcelColumn( columnName, ColumnType.ConceptName, index );
				case "lo /concept id":
					return new ExcelColumn( columnName, ColumnType.ConceptId, index );
				case "id (original & new place holders)":
					return new ExcelColumn( columnName, ColumnType.QuestionId, index );
				case "question":
					return new ExcelColumn( columnName, ColumnType.Question, index );
				case "correct":
					return new ExcelColumn( columnName, ColumnType.Correct, index );
				case "answer_1":
					return new ExcelColumn( columnName, ColumnType.Answer1, index );
				case "answer_2":
					return new ExcelColumn( columnName, ColumnType.Answer2, index );
				case "answer_3":
					return new ExcelColumn( columnName, ColumnType.Answer3, index );
				case "answer_4":
					return new ExcelColumn( columnName, ColumnType.Answer4, index );
				case "justification":
					return new ExcelColumn( columnName, ColumnType.Justification, index );
				case "answer_image_url":
					return new ExcelColumn( columnName, ColumnType.AnswerImageUrl, index );
				case "question_image_url":
					return new ExcelColumn( columnName, ColumnType.QuestionImageUrl, index );
				case "item id":
					return new ExcelColumn( columnName, ColumnType.ItemId, index );
				case "atom type":
					return new ExcelColumn( columnName, ColumnType.AtomType, index );
				case "atom id":
					return new ExcelColumn( columnName, ColumnType.AtomId, index );
				default:
					return new ExcelColumn( columnName, ColumnType.Undefined, index );
			}
		}


	}

	public enum ColumnType
	{
		Undefined,
		TopicName,
		SessionName,
		ReadingName,
		Band,
		ConceptName,
		ConceptId,
		QuestionId,
		Question,
		Correct,
		Answer1,
		Answer2,
		Answer3,
		Answer4,
		Justification,
		AnswerImageUrl,
		QuestionImageUrl,
		ItemId,
		AtomType,
		AtomId
	}
}
