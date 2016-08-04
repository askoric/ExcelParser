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
			this.label1 = new System.Windows.Forms.Label();
			this.button1 = new System.Windows.Forms.Button();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.label4 = new System.Windows.Forms.Label();
			this.label5 = new System.Windows.Forms.Label();
			this.label6 = new System.Windows.Forms.Label();
			this.label7 = new System.Windows.Forms.Label();
			this.StatusLabel = new System.Windows.Forms.Label();
			this.SetTranscript = new System.Windows.Forms.CheckBox();
			this.SuspendLayout();
			// 
			// label1
			// 
			this.label1.AutoSize = true;
			this.label1.Location = new System.Drawing.Point(110, 33);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(58, 13);
			this.label1.TabIndex = 0;
			this.label1.Text = "Klikenzi =>";
			// 
			// button1
			// 
			this.button1.Location = new System.Drawing.Point(174, 28);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(75, 23);
			this.button1.TabIndex = 1;
			this.button1.Text = "I feel lucky";
			this.button1.UseVisualStyleBackColor = true;
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(32, 70);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(65, 13);
			this.label2.TabIndex = 2;
			this.label2.Text = "Instructione:";
			// 
			// label3
			// 
			this.label3.AutoSize = true;
			this.label3.Location = new System.Drawing.Point(46, 92);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(318, 13);
			this.label3.TabIndex = 3;
			this.label3.Text = "1) Click button above and first upload excel file with main structure";
			// 
			// label4
			// 
			this.label4.AutoSize = true;
			this.label4.Location = new System.Drawing.Point(50, 114);
			this.label4.Name = "label4";
			this.label4.Size = new System.Drawing.Size(321, 13);
			this.label4.TabIndex = 4;
			this.label4.Text = "2) After uploading of first excel upload excel with question structure";
			// 
			// label5
			// 
			this.label5.AutoSize = true;
			this.label5.Location = new System.Drawing.Point(50, 136);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(231, 13);
			this.label5.TabIndex = 5;
			this.label5.Text = "3) Save output result as xml. Viola you are done";
			// 
			// label6
			// 
			this.label6.AutoSize = true;
			this.label6.Location = new System.Drawing.Point(29, 159);
			this.label6.Name = "label6";
			this.label6.Size = new System.Drawing.Size(275, 13);
			this.label6.TabIndex = 6;
			this.label6.Text = "!important excel files need to be in 97-2003 excel format. ";
			// 
			// label7
			// 
			this.label7.AutoSize = true;
			this.label7.Location = new System.Drawing.Point(32, 214);
			this.label7.Name = "label7";
			this.label7.Size = new System.Drawing.Size(43, 13);
			this.label7.TabIndex = 7;
			this.label7.Text = "Status: ";
			// 
			// StatusLabel
			// 
			this.StatusLabel.AutoSize = true;
			this.StatusLabel.Location = new System.Drawing.Point(91, 214);
			this.StatusLabel.Name = "StatusLabel";
			this.StatusLabel.Size = new System.Drawing.Size(0, 13);
			this.StatusLabel.TabIndex = 8;
			// 
			// SetTranscript
			// 
			this.SetTranscript.AutoSize = true;
			this.SetTranscript.Checked = true;
			this.SetTranscript.CheckState = System.Windows.Forms.CheckState.Checked;
			this.SetTranscript.Location = new System.Drawing.Point(284, 32);
			this.SetTranscript.Name = "SetTranscript";
			this.SetTranscript.Size = new System.Drawing.Size(89, 17);
			this.SetTranscript.TabIndex = 9;
			this.SetTranscript.Text = "SetTranscript";
			this.SetTranscript.UseVisualStyleBackColor = true;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(578, 256);
			this.Controls.Add(this.SetTranscript);
			this.Controls.Add(this.StatusLabel);
			this.Controls.Add(this.label7);
			this.Controls.Add(this.label6);
			this.Controls.Add(this.label5);
			this.Controls.Add(this.label4);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.label1);
			this.Name = "MainForm";
			this.Text = "Excel => Xml";
			this.Load += new System.EventHandler(this.MainForm_Load);
			this.ResumeLayout(false);
			this.PerformLayout();

		}

		#endregion

		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label label4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.Label label6;
		private System.Windows.Forms.Label label7;
		private System.Windows.Forms.Label StatusLabel;
		private System.Windows.Forms.CheckBox SetTranscript;
	}
}