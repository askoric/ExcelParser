using System;
using System.Collections.Generic;
using System.Linq;
using System.Net.Mail;
using System.Runtime.Remoting.Messaging;
using System.Text;
using System.Threading.Tasks;

namespace ExcelParser
{
	public interface IExcelColumn<T>
	{
		string Value { get; set; }
		int ColumnIndex { get; set; }
		T Type { get; set; }

		bool HaveValue();

		IExcelColumn<T> GetColumn( string columnName, int index );

		bool IsRecognizableColumn();

	}


	public abstract class ExcelColumn<T> : IExcelColumn<T>
	{
		protected ExcelColumn()
		{

		}

		protected ExcelColumn( string value, T type, int index )
		{
			Value = value;
			Type = type;
			ColumnIndex = index;
		}

		public string Value { get; set; }
		public int ColumnIndex { get; set; }
		public T Type { get; set; }


		public bool HaveValue()
		{
			return Value != null && Value.Trim() != "";
		}

		public abstract IExcelColumn<T> GetColumn( string columnName, int index );
		public abstract bool IsRecognizableColumn();

	}

	public class LosExcelColumn : ExcelColumn<LosExcelColumnType> 
	{

		public LosExcelColumn()
		{
		}

		public LosExcelColumn(string value, LosExcelColumnType type, int index) : base(value, type, index) { }

		public override bool IsRecognizableColumn()
		{
			return this.Type != LosExcelColumnType.Undefined;
		}

		public override IExcelColumn<LosExcelColumnType> GetColumn( string columnName, int index )
		{
			if ( columnName != null )
			{
				switch ( columnName.ToLower().Trim() ) {
					case "topictitle":
						return new LosExcelColumn( columnName, LosExcelColumnType.TopicTitle, index );
					case "sessiontitle":
						return new LosExcelColumn( columnName, LosExcelColumnType.SessionTitle, index );
					case "readingtitle":
						return new LosExcelColumn( columnName, LosExcelColumnType.ReadingTitle, index );
					case "cfa_alpha":
						return new LosExcelColumn( columnName, LosExcelColumnType.CfaAlpha, index );
					case "los text":
						return new LosExcelColumn( columnName, LosExcelColumnType.LosText, index );
				}
			}
			return new LosExcelColumn( columnName, LosExcelColumnType.Undefined, index );
		}
	}

	public enum LosExcelColumnType
	{
		Undefined,
		TopicTitle,
		SessionTitle,
		ReadingTitle,
		CfaAlpha,
		LosText,
	}


	public class QuestionExcelColumn : ExcelColumn<QuestionExcelColumnType>
	{
		public QuestionExcelColumn()
		{
		}

		public QuestionExcelColumn( string value, QuestionExcelColumnType type, int index ) : base( value, type, index ) { }

		public override bool IsRecognizableColumn()
		{
			return this.Type != QuestionExcelColumnType.Undefined;
		}

		public override IExcelColumn<QuestionExcelColumnType> GetColumn( string columnName, int index )
		{
			if ( columnName != null ) {
				switch ( columnName.ToLower().Trim() ) {
					case "id (original & new place holders)":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.QuestionId, index );
					case "question":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Question, index );
					case "question_image_url":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.QuestionImageUrl, index );
					case "answer_image_url":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.AnswerImageUrl, index );
					case "correct":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Correct, index );
					case "answer_1":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Answer1, index );
					case "answer_2":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Answer2, index );
					case "answer_3":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Answer3, index );
					case "answer_4":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Answer4, index );
					case "justification":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Justification, index );
					case "(kk/ee) instruct (k/e) asseess":
						return new QuestionExcelColumn( columnName, QuestionExcelColumnType.KKEE, index );
				}
			}

			return new QuestionExcelColumn( columnName, QuestionExcelColumnType.Undefined, index );
		}
	}


	public enum QuestionExcelColumnType
	{
		Undefined,
		QuestionId,
		Question,
		QuestionImageUrl,
		AnswerImageUrl,
		Correct,
		Answer1,
		Answer2,
		Answer3,
		Answer4,
		Justification,
		KKEE
	}


	public class AcceptanceCriteriaExcelColumn : ExcelColumn<AcceptanceCriteriaColumnType>
	{
		public AcceptanceCriteriaExcelColumn()
		{
		}

		public AcceptanceCriteriaExcelColumn( string value, AcceptanceCriteriaColumnType type, int index ) : base( value, type, index ) { }

		public override bool IsRecognizableColumn()
		{
			return this.Type != AcceptanceCriteriaColumnType.Undefined;
		}

		public override IExcelColumn<AcceptanceCriteriaColumnType> GetColumn( string columnName, int index )
		{
			if ( columnName != null ) {
				switch ( columnName.ToLower().Trim() ) {
					case "lo1":
						return new AcceptanceCriteriaExcelColumn( columnName, AcceptanceCriteriaColumnType.Lo1, index );
					case "target score":
						return new AcceptanceCriteriaExcelColumn( columnName, AcceptanceCriteriaColumnType.TargetScore, index );
				}
			}

			return new AcceptanceCriteriaExcelColumn( columnName, AcceptanceCriteriaColumnType.Undefined, index );
		}
	}


	public enum AcceptanceCriteriaColumnType
	{
		Undefined,
		Lo1,
		TargetScore
	}

	public class TestExcelColumn : ExcelColumn<TestExcelColumnType>
	{
		public TestExcelColumn()
		{
		}

		public TestExcelColumn( string value, TestExcelColumnType type, int index ) : base( value, type, index ) { }

		public override bool IsRecognizableColumn()
		{
			return this.Type != TestExcelColumnType.Undefined;
		}

		public override IExcelColumn<TestExcelColumnType> GetColumn( string columnName, int index )
		{
			if ( columnName != null ) {
				switch ( columnName.ToLower().Trim() ) {
					case "session abbreviation":
						return new TestExcelColumn( columnName, TestExcelColumnType.SessionAbbrevation, index );
					case "topic abbreviation":
						return new TestExcelColumn( columnName, TestExcelColumnType.TopicAbbrevation, index );
					case "topicname":
						return new TestExcelColumn( columnName, TestExcelColumnType.TopicName, index );
					case "k_structure":
						return new TestExcelColumn( columnName, TestExcelColumnType.KStructure, index );
					case "q_id":
					case "id (original & new place holders)":
						return new TestExcelColumn( columnName, TestExcelColumnType.QuestionId, index );
					case "q_type":
						return new TestExcelColumn( columnName, TestExcelColumnType.QuestionType, index );
					case "question":
						return new TestExcelColumn( columnName, TestExcelColumnType.Question, index );
					case "answer_1":
						return new TestExcelColumn( columnName, TestExcelColumnType.Answer1, index );
					case "answer_2":
						return new TestExcelColumn( columnName, TestExcelColumnType.Answer2, index );
					case "answer_3":
						return new TestExcelColumn( columnName, TestExcelColumnType.Answer3, index );
					case "answer_4":
						return new TestExcelColumn( columnName, TestExcelColumnType.Answer4, index );
					case "answer_image_url":
						return new TestExcelColumn( columnName, TestExcelColumnType.AnswerImageUrl, index );
					case "correct":
						return new TestExcelColumn( columnName, TestExcelColumnType.Correct, index );
					case "question_image_url":
						return new TestExcelColumn( columnName, TestExcelColumnType.QuestionImageUrl, index );
					case "justification":
						return new TestExcelColumn( columnName, TestExcelColumnType.Justification, index );
                    case "fcm_number":
                        return new TestExcelColumn(columnName, TestExcelColumnType.FcmNumber, index);
                }
			}

			return new TestExcelColumn( columnName, TestExcelColumnType.Undefined, index );
		}
	}


	public enum TestExcelColumnType
	{
		Undefined,
		SessionAbbrevation,
		TopicAbbrevation,
		TopicName,
		KStructure,
		QuestionId,
		Question,
		QuestionType,
		Answer1, 
		Answer2, 
		Answer3,
		Answer4,
		AnswerImageUrl, 
		Correct,
		QuestionImageUrl, 
		Justification,
        FcmNumber //MOckTestExcelOnly
	}


	public class MainStructureExcelColumn : ExcelColumn<MainStructureColumnType>
	{
		public MainStructureExcelColumn()
		{
		}

		public MainStructureExcelColumn(string value, MainStructureColumnType type, int index) : base(value, type, index) { }

		public override bool IsRecognizableColumn()
		{
			return this.Type != MainStructureColumnType.Undefined;
		}

		public override IExcelColumn<MainStructureColumnType> GetColumn( string columnName, int index )
		{
			if (columnName != null)
			{
				switch ( columnName.ToLower().Trim() ) {
					case "topictitle":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.TopicName, index );
					case "topic":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.TopicShortName, index );
					case "sessiontitle":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.SessionName, index );
					case "readingtitle":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.ReadingName, index );
					case "reading abb":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.ReadingId, index );
					case "band":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Band, index );
					case "bandid":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.BandId, index );
					case "lo description":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.ConceptName, index );
					case "lo /concept id":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.ConceptId, index );
					case "item id":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.ItemId, index );
					case "atom type":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.AtomType, index );
					case "atom id":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.AtomId, index );
					case "atom title":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.AtomTitle, index );
					case "studysession":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.StudySession, index );
					case "studysession abb":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.StudySessionId, index );
					case "structure":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Structure, index );
					case "reading":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Reading, index );
					case "exam percentage":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.ExamPercentage, index );
					case "description":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Description, index );
					case "downloads":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Downloads, index );
					case "downloads2":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Downloads2, index );
					case "locked":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Locked, index );
					case "color":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.Color, index );
					case "type":
						return new MainStructureExcelColumn( columnName, MainStructureColumnType.CfaType, index );
				}

			}

			return new MainStructureExcelColumn( columnName, MainStructureColumnType.Undefined, index );
		}


	}

	public enum MainStructureColumnType
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
		ItemId, 
		AtomType, 
		AtomId, 
		AtomTitle, 
		StudySession,
		StudySessionId, 
		Structure,
		Reading, 
		ExamPercentage, 
		Description, 
		Downloads, 
		Downloads2, 
		Locked, 
		Color, 
		CfaType 
	}
}
