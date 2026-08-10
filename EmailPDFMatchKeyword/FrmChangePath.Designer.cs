namespace EmailPDFMatchKeyword
{
	partial class FrmChangePath
	{
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.IContainer components = null;

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		/// <param name="disposing">true if managed resources should be disposed; otherwise, false.</param>
		protected override void Dispose(bool disposing)
		{
			if (disposing && (components != null))
			{
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent()
		{
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(FrmChangePath));
			btnpathsave = new Button();
			btncancel = new Button();
			txtwritepath = new TextBox();
			lblwritepath = new Label();
			btn2Browse = new Button();
			lblsheetid = new Label();
			txtsheetid = new TextBox();
			pictureBox1 = new PictureBox();
			((System.ComponentModel.ISupportInitialize)pictureBox1).BeginInit();
			SuspendLayout();
			// 
			// btnpathsave
			// 
			btnpathsave.BackColor = SystemColors.ActiveCaption;
			btnpathsave.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
			btnpathsave.Location = new Point(53, 286);
			btnpathsave.Name = "btnpathsave";
			btnpathsave.Size = new Size(138, 32);
			btnpathsave.TabIndex = 2;
			btnpathsave.Text = "Save";
			btnpathsave.UseVisualStyleBackColor = false;
			btnpathsave.Click += btnpathsave_Click;
			// 
			// btncancel
			// 
			btncancel.BackColor = SystemColors.ActiveCaption;
			btncancel.Font = new Font("Segoe UI", 9.75F, FontStyle.Regular, GraphicsUnit.Point, 0);
			btncancel.Location = new Point(208, 286);
			btncancel.Name = "btncancel";
			btncancel.Size = new Size(138, 32);
			btncancel.TabIndex = 3;
			btncancel.Text = "Cancel";
			btncancel.UseVisualStyleBackColor = false;
			btncancel.Click += btncancel_Click;
			// 
			// txtwritepath
			// 
			txtwritepath.Location = new Point(52, 57);
			txtwritepath.Name = "txtwritepath";
			txtwritepath.Size = new Size(282, 23);
			txtwritepath.TabIndex = 4;
			// 
			// lblwritepath
			// 
			lblwritepath.AutoSize = true;
			lblwritepath.Location = new Point(52, 39);
			lblwritepath.Name = "lblwritepath";
			lblwritepath.Size = new Size(186, 15);
			lblwritepath.TabIndex = 5;
			lblwritepath.Text = "Folder path to write(Google Drive)";
			// 
			// btn2Browse
			// 
			btn2Browse.Location = new Point(340, 57);
			btn2Browse.Name = "btn2Browse";
			btn2Browse.Size = new Size(102, 23);
			btn2Browse.TabIndex = 7;
			btn2Browse.Text = "Browse...";
			btn2Browse.UseVisualStyleBackColor = true;
			btn2Browse.Click += btn2Browse_Click;
			// 
			// lblsheetid
			// 
			lblsheetid.AutoSize = true;
			lblsheetid.Location = new Point(52, 100);
			lblsheetid.Name = "lblsheetid";
			lblsheetid.Size = new Size(90, 15);
			lblsheetid.TabIndex = 8;
			lblsheetid.Text = "SpreadsheetId : ";
			// 
			// txtsheetid
			// 
			txtsheetid.Location = new Point(52, 118);
			txtsheetid.Name = "txtsheetid";
			txtsheetid.Size = new Size(282, 23);
			txtsheetid.TabIndex = 9;
			// 
			// pictureBox1
			// 
			pictureBox1.Image = (Image)resources.GetObject("pictureBox1.Image");
			pictureBox1.Location = new Point(52, 168);
			pictureBox1.Name = "pictureBox1";
			pictureBox1.Size = new Size(294, 86);
			pictureBox1.TabIndex = 10;
			pictureBox1.TabStop = false;
			// 
			// FrmChangePath
			// 
			AutoScaleDimensions = new SizeF(7F, 15F);
			AutoScaleMode = AutoScaleMode.Font;
			BackColor = Color.LightGray;
			ClientSize = new Size(491, 352);
			Controls.Add(pictureBox1);
			Controls.Add(txtsheetid);
			Controls.Add(lblsheetid);
			Controls.Add(btn2Browse);
			Controls.Add(lblwritepath);
			Controls.Add(txtwritepath);
			Controls.Add(btncancel);
			Controls.Add(btnpathsave);
			MaximizeBox = false;
			MinimizeBox = false;
			Name = "FrmChangePath";
			Text = "Folder Path Settings";
			((System.ComponentModel.ISupportInitialize)pictureBox1).EndInit();
			ResumeLayout(false);
			PerformLayout();
		}

		#endregion
		private Button btnpathsave;
		private Button btncancel;
		private TextBox txtwritepath;
		private Label lblwritepath;
		private Button btn2Browse;
		private Label lblsheetid;
		private TextBox txtsheetid;
		private PictureBox pictureBox1;
	}
}