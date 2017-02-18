namespace ExcelParser
{
	partial class MainForm
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose( bool disposing )
		{
			if ( disposing && (components != null) ) {
				components.Dispose();
			}
			base.Dispose( disposing );
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
            System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
            this.label6 = new System.Windows.Forms.Label();
            this.label7 = new System.Windows.Forms.Label();
            this.StatusLabel = new System.Windows.Forms.Label();
            this.SetTranscript = new System.Windows.Forms.CheckBox();
            this.FillDbFromExistingCourseXml = new System.Windows.Forms.Button();
            this.UploadMainStructureExcelBtn = new System.Windows.Forms.Button();
            this.UploadQuestionsExcelBtn = new System.Windows.Forms.Button();
            this.UploadLOSExcelBtn = new System.Windows.Forms.Button();
            this.UploadAcceptanceCriteriaExcel = new System.Windows.Forms.Button();
            this.GenerateCourseXmlBtn = new System.Windows.Forms.Button();
            this.MainStructureExcelCheckImg = new System.Windows.Forms.PictureBox();
            this.LosExcelCheckImg = new System.Windows.Forms.PictureBox();
            this.QuestionExcelCheckImg = new System.Windows.Forms.PictureBox();
            this.AcceptanceCriteriaCheckImg = new System.Windows.Forms.PictureBox();
            this.UploadSSTestsExcel = new System.Windows.Forms.Button();
            this.UploadSsTestCheckImage = new System.Windows.Forms.PictureBox();
            this.uploadProgressTestCheckIcon = new System.Windows.Forms.PictureBox();
            this.UploadProgressTestExcell = new System.Windows.Forms.Button();
            ((System.ComponentModel.ISupportInitialize)(this.MainStructureExcelCheckImg)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.LosExcelCheckImg)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.QuestionExcelCheckImg)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.AcceptanceCriteriaCheckImg)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.UploadSsTestCheckImage)).BeginInit();
            ((System.ComponentModel.ISupportInitialize)(this.uploadProgressTestCheckIcon)).BeginInit();
            this.SuspendLayout();
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(223, 206);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(227, 13);
            this.label6.TabIndex = 6;
            this.label6.Text = "Excel files need to be in 97-2003 excel format. ";
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(22, 243);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(43, 13);
            this.label7.TabIndex = 7;
            this.label7.Text = "Status: ";
            // 
            // StatusLabel
            // 
            this.StatusLabel.AutoSize = true;
            this.StatusLabel.Location = new System.Drawing.Point(71, 243);
            this.StatusLabel.Name = "StatusLabel";
            this.StatusLabel.Size = new System.Drawing.Size(0, 13);
            this.StatusLabel.TabIndex = 8;
            // 
            // SetTranscript
            // 
            this.SetTranscript.AutoSize = true;
            this.SetTranscript.Enabled = false;
            this.SetTranscript.Location = new System.Drawing.Point(25, 206);
            this.SetTranscript.Name = "SetTranscript";
            this.SetTranscript.Size = new System.Drawing.Size(89, 17);
            this.SetTranscript.TabIndex = 9;
            this.SetTranscript.Text = "SetTranscript";
            this.SetTranscript.UseVisualStyleBackColor = true;
            // 
            // FillDbFromExistingCourseXml
            // 
            this.FillDbFromExistingCourseXml.Location = new System.Drawing.Point(205, 158);
            this.FillDbFromExistingCourseXml.Name = "FillDbFromExistingCourseXml";
            this.FillDbFromExistingCourseXml.Size = new System.Drawing.Size(245, 27);
            this.FillDbFromExistingCourseXml.TabIndex = 12;
            this.FillDbFromExistingCourseXml.Text = "Fill Db from Existing course XML";
            this.FillDbFromExistingCourseXml.UseVisualStyleBackColor = true;
            this.FillDbFromExistingCourseXml.Click += new System.EventHandler(this.FillDbFromExistingCourseXml_Click);
            // 
            // UploadMainStructureExcelBtn
            // 
            this.UploadMainStructureExcelBtn.Location = new System.Drawing.Point(25, 12);
            this.UploadMainStructureExcelBtn.Name = "UploadMainStructureExcelBtn";
            this.UploadMainStructureExcelBtn.Size = new System.Drawing.Size(152, 23);
            this.UploadMainStructureExcelBtn.TabIndex = 1;
            this.UploadMainStructureExcelBtn.Text = "Upload Main Structure Excel";
            this.UploadMainStructureExcelBtn.UseVisualStyleBackColor = true;
            this.UploadMainStructureExcelBtn.Click += new System.EventHandler(this.UploadMainStructureExcelBtn_Click);
            // 
            // UploadQuestionsExcelBtn
            // 
            this.UploadQuestionsExcelBtn.Location = new System.Drawing.Point(25, 41);
            this.UploadQuestionsExcelBtn.Name = "UploadQuestionsExcelBtn";
            this.UploadQuestionsExcelBtn.Size = new System.Drawing.Size(152, 23);
            this.UploadQuestionsExcelBtn.TabIndex = 2;
            this.UploadQuestionsExcelBtn.Text = "Upload Questions Excel";
            this.UploadQuestionsExcelBtn.UseVisualStyleBackColor = true;
            this.UploadQuestionsExcelBtn.Click += new System.EventHandler(this.UploadQuestionsExcelBtn_Click);
            // 
            // UploadLOSExcelBtn
            // 
            this.UploadLOSExcelBtn.Location = new System.Drawing.Point(25, 70);
            this.UploadLOSExcelBtn.Name = "UploadLOSExcelBtn";
            this.UploadLOSExcelBtn.Size = new System.Drawing.Size(152, 23);
            this.UploadLOSExcelBtn.TabIndex = 3;
            this.UploadLOSExcelBtn.Text = "Upload LOS Excel";
            this.UploadLOSExcelBtn.UseVisualStyleBackColor = true;
            this.UploadLOSExcelBtn.Click += new System.EventHandler(this.UploadLOSExcelBtn_Click);
            // 
            // UploadAcceptanceCriteriaExcel
            // 
            this.UploadAcceptanceCriteriaExcel.Location = new System.Drawing.Point(25, 99);
            this.UploadAcceptanceCriteriaExcel.Name = "UploadAcceptanceCriteriaExcel";
            this.UploadAcceptanceCriteriaExcel.Size = new System.Drawing.Size(152, 23);
            this.UploadAcceptanceCriteriaExcel.TabIndex = 4;
            this.UploadAcceptanceCriteriaExcel.Text = "Upload Acceptance Criteria Excel";
            this.UploadAcceptanceCriteriaExcel.UseVisualStyleBackColor = true;
            this.UploadAcceptanceCriteriaExcel.Click += new System.EventHandler(this.UploadAcceptanceCriteriaExcel_Click);
            // 
            // GenerateCourseXmlBtn
            // 
            this.GenerateCourseXmlBtn.Location = new System.Drawing.Point(205, 12);
            this.GenerateCourseXmlBtn.Name = "GenerateCourseXmlBtn";
            this.GenerateCourseXmlBtn.Size = new System.Drawing.Size(245, 139);
            this.GenerateCourseXmlBtn.TabIndex = 5;
            this.GenerateCourseXmlBtn.Text = "Generate Course XML";
            this.GenerateCourseXmlBtn.UseVisualStyleBackColor = true;
            this.GenerateCourseXmlBtn.Click += new System.EventHandler(this.GenerateCourseXmlBtn_Click);
            // 
            // MainStructureExcelCheckImg
            // 
            this.MainStructureExcelCheckImg.Image = ((System.Drawing.Image)(resources.GetObject("MainStructureExcelCheckImg.Image")));
            this.MainStructureExcelCheckImg.Location = new System.Drawing.Point(183, 12);
            this.MainStructureExcelCheckImg.Name = "MainStructureExcelCheckImg";
            this.MainStructureExcelCheckImg.Size = new System.Drawing.Size(16, 23);
            this.MainStructureExcelCheckImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.MainStructureExcelCheckImg.TabIndex = 13;
            this.MainStructureExcelCheckImg.TabStop = false;
            this.MainStructureExcelCheckImg.Visible = false;
            // 
            // LosExcelCheckImg
            // 
            this.LosExcelCheckImg.Image = ((System.Drawing.Image)(resources.GetObject("LosExcelCheckImg.Image")));
            this.LosExcelCheckImg.Location = new System.Drawing.Point(183, 70);
            this.LosExcelCheckImg.Name = "LosExcelCheckImg";
            this.LosExcelCheckImg.Size = new System.Drawing.Size(16, 23);
            this.LosExcelCheckImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.LosExcelCheckImg.TabIndex = 14;
            this.LosExcelCheckImg.TabStop = false;
            this.LosExcelCheckImg.Visible = false;
            // 
            // QuestionExcelCheckImg
            // 
            this.QuestionExcelCheckImg.Image = ((System.Drawing.Image)(resources.GetObject("QuestionExcelCheckImg.Image")));
            this.QuestionExcelCheckImg.Location = new System.Drawing.Point(183, 41);
            this.QuestionExcelCheckImg.Name = "QuestionExcelCheckImg";
            this.QuestionExcelCheckImg.Size = new System.Drawing.Size(16, 23);
            this.QuestionExcelCheckImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.QuestionExcelCheckImg.TabIndex = 15;
            this.QuestionExcelCheckImg.TabStop = false;
            this.QuestionExcelCheckImg.Visible = false;
            // 
            // AcceptanceCriteriaCheckImg
            // 
            this.AcceptanceCriteriaCheckImg.Image = ((System.Drawing.Image)(resources.GetObject("AcceptanceCriteriaCheckImg.Image")));
            this.AcceptanceCriteriaCheckImg.Location = new System.Drawing.Point(183, 99);
            this.AcceptanceCriteriaCheckImg.Name = "AcceptanceCriteriaCheckImg";
            this.AcceptanceCriteriaCheckImg.Size = new System.Drawing.Size(16, 23);
            this.AcceptanceCriteriaCheckImg.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.AcceptanceCriteriaCheckImg.TabIndex = 16;
            this.AcceptanceCriteriaCheckImg.TabStop = false;
            this.AcceptanceCriteriaCheckImg.Visible = false;
            // 
            // UploadSSTestsExcel
            // 
            this.UploadSSTestsExcel.Location = new System.Drawing.Point(25, 128);
            this.UploadSSTestsExcel.Name = "UploadSSTestsExcel";
            this.UploadSSTestsExcel.Size = new System.Drawing.Size(152, 23);
            this.UploadSSTestsExcel.TabIndex = 17;
            this.UploadSSTestsExcel.Text = "Upload SS tests Excel";
            this.UploadSSTestsExcel.UseVisualStyleBackColor = true;
            this.UploadSSTestsExcel.Click += new System.EventHandler(this.UploadSSTestsExcel_Click);
            // 
            // UploadSsTestCheckImage
            // 
            this.UploadSsTestCheckImage.Image = ((System.Drawing.Image)(resources.GetObject("UploadSsTestCheckImage.Image")));
            this.UploadSsTestCheckImage.Location = new System.Drawing.Point(183, 128);
            this.UploadSsTestCheckImage.Name = "UploadSsTestCheckImage";
            this.UploadSsTestCheckImage.Size = new System.Drawing.Size(16, 23);
            this.UploadSsTestCheckImage.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.UploadSsTestCheckImage.TabIndex = 18;
            this.UploadSsTestCheckImage.TabStop = false;
            this.UploadSsTestCheckImage.Visible = false;
            // 
            // uploadProgressTestCheckIcon
            // 
            this.uploadProgressTestCheckIcon.Image = ((System.Drawing.Image)(resources.GetObject("uploadProgressTestCheckIcon.Image")));
            this.uploadProgressTestCheckIcon.Location = new System.Drawing.Point(183, 162);
            this.uploadProgressTestCheckIcon.Name = "uploadProgressTestCheckIcon";
            this.uploadProgressTestCheckIcon.Size = new System.Drawing.Size(16, 23);
            this.uploadProgressTestCheckIcon.SizeMode = System.Windows.Forms.PictureBoxSizeMode.CenterImage;
            this.uploadProgressTestCheckIcon.TabIndex = 20;
            this.uploadProgressTestCheckIcon.TabStop = false;
            this.uploadProgressTestCheckIcon.Visible = false;
            // 
            // UploadProgressTestExcell
            // 
            this.UploadProgressTestExcell.Location = new System.Drawing.Point(25, 162);
            this.UploadProgressTestExcell.Name = "UploadProgressTestExcell";
            this.UploadProgressTestExcell.Size = new System.Drawing.Size(152, 23);
            this.UploadProgressTestExcell.TabIndex = 19;
            this.UploadProgressTestExcell.Text = "Upload Progress tests Excel";
            this.UploadProgressTestExcell.UseVisualStyleBackColor = true;
            this.UploadProgressTestExcell.Click += new System.EventHandler(this.UploadProgressTestExcell_Click);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(462, 277);
            this.Controls.Add(this.uploadProgressTestCheckIcon);
            this.Controls.Add(this.UploadProgressTestExcell);
            this.Controls.Add(this.UploadSsTestCheckImage);
            this.Controls.Add(this.UploadSSTestsExcel);
            this.Controls.Add(this.AcceptanceCriteriaCheckImg);
            this.Controls.Add(this.QuestionExcelCheckImg);
            this.Controls.Add(this.LosExcelCheckImg);
            this.Controls.Add(this.MainStructureExcelCheckImg);
            this.Controls.Add(this.GenerateCourseXmlBtn);
            this.Controls.Add(this.UploadAcceptanceCriteriaExcel);
            this.Controls.Add(this.UploadLOSExcelBtn);
            this.Controls.Add(this.UploadQuestionsExcelBtn);
            this.Controls.Add(this.UploadMainStructureExcelBtn);
            this.Controls.Add(this.FillDbFromExistingCourseXml);
            this.Controls.Add(this.SetTranscript);
            this.Controls.Add(this.StatusLabel);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.label6);
            this.Name = "MainForm";
            this.Text = "Course transformer";
            ((System.ComponentModel.ISupportInitialize)(this.MainStructureExcelCheckImg)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.LosExcelCheckImg)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.QuestionExcelCheckImg)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.AcceptanceCriteriaCheckImg)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.UploadSsTestCheckImage)).EndInit();
            ((System.ComponentModel.ISupportInitialize)(this.uploadProgressTestCheckIcon)).EndInit();
            this.ResumeLayout(false);
            this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label StatusLabel;
		private System.Windows.Forms.CheckBox SetTranscript;
		private System.Windows.Forms.Button FillDbFromExistingCourseXml;
		private System.Windows.Forms.Button UploadMainStructureExcelBtn;
		private System.Windows.Forms.Button UploadQuestionsExcelBtn;
		private System.Windows.Forms.Button UploadLOSExcelBtn;
		private System.Windows.Forms.Button UploadAcceptanceCriteriaExcel;
		private System.Windows.Forms.Button GenerateCourseXmlBtn;
		private System.Windows.Forms.PictureBox MainStructureExcelCheckImg;
		private System.Windows.Forms.PictureBox LosExcelCheckImg;
		private System.Windows.Forms.PictureBox QuestionExcelCheckImg;
		private System.Windows.Forms.PictureBox AcceptanceCriteriaCheckImg;
		private System.Windows.Forms.Button UploadSSTestsExcel;
		private System.Windows.Forms.PictureBox UploadSsTestCheckImage;
		private System.Windows.Forms.PictureBox uploadProgressTestCheckIcon;
		private System.Windows.Forms.Button UploadProgressTestExcell;
	}
}