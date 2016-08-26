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
				case "topic":
					return new ExcelColumn( columnName, ColumnType.TopicShortName, index );
				case "sessionname":
				case "sessiontitle":
					return new ExcelColumn( columnName, ColumnType.SessionName, index );
				case "readingname":
				case "readingtitle":
					return new ExcelColumn( columnName, ColumnType.ReadingName, index );
				case "reading abb":
					return new ExcelColumn( columnName, ColumnType.ReadingId, index );
				case "band":
					return new ExcelColumn( columnName, ColumnType.Band, index );
				case "bandid":
					return new ExcelColumn( columnName, ColumnType.BandId, index );
				case "conceptname":
				case "lo description":
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
				case "atom title":
					return new ExcelColumn( columnName, ColumnType.AtomTitle, index );
				case "studysession":
					return new ExcelColumn( columnName, ColumnType.StudySession, index );
				case "studysession abb":
					return new ExcelColumn( columnName, ColumnType.StudySessionId, index );
				case "reading":
					return new ExcelColumn( columnName, ColumnType.Reading, index );
				case "exam percentage":
					return new ExcelColumn( columnName, ColumnType.ExamPercentage, index );
				case "(kk/ee) instruct (k/e) asseess":
					return new ExcelColumn( columnName, ColumnType.KKEE, index );
				case "description":
					return new ExcelColumn( columnName, ColumnType.Description, index );
				case "downloads":
					return new ExcelColumn( columnName, ColumnType.Downloads, index );
				case "downloads2":
					return new ExcelColumn( columnName, ColumnType.Downloads2, index );
				case "cfa_alpha":
					return new ExcelColumn( columnName, ColumnType.CfaAlpha, index );
				case "los text":
					return new ExcelColumn( columnName, ColumnType.LosText, index );
				case "lo1":
					return new ExcelColumn( columnName, ColumnType.Lo1, index );
				case "target score":
					return new ExcelColumn( columnName, ColumnType.TargetScore, index );
				case "locked":
					return new ExcelColumn( columnName, ColumnType.Locked, index );
				case "color":
					return new ExcelColumn( columnName, ColumnType.Color, index );
				case "type":
					return new ExcelColumn( columnName, ColumnType.CfaType, index );
				default:
					return new ExcelColumn( columnName, ColumnType.Undefined, index );
			}
		}


	}

	public enum ColumnType
	{
		Undefined,
		TopicName,
		TopicShortName,
		SessionName,
		ReadingName,
		ReadingId,
		Band,
		BandId,
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
		AtomId,
		AtomTitle,
		StudySession,
		StudySessionId,
		Reading,
		KKEE,
		ExamPercentage,
		Description,
		Downloads,
		Downloads2,
		CfaAlpha,
		LosText,
		Lo1,
		TargetScore,
		Locked,
		Color,
		CfaType
	}
}
