using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;

namespace OutputExcel
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class Form1 : System.Windows.Forms.Form
	{
        private System.Windows.Forms.TextBox tbFileName;
        private System.Windows.Forms.Button cmdCreateFileName;
        private System.Windows.Forms.TextBox tbPath;
        private System.Windows.Forms.LinkLabel lnkPath;
        private System.Windows.Forms.FolderBrowserDialog folderBrowserDialog1;
        private Button btnReadFileName;
        private Label label1;
        private TextBox txtStart;
        private TextBox txtStop;
        private Label label2;
        private Button btnProcessCenterStudentUpdate;
        private Button btnSessionStudents;
		/// <summary>
		/// Required designer variable.
		/// </summary>
		private System.ComponentModel.Container components = null;

		public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();

			//
			// TODO: Add any constructor code after InitializeComponent call
			//
		}

		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			if( disposing )
			{
				if (components != null) 
				{
					components.Dispose();
				}
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
            this.tbFileName = new System.Windows.Forms.TextBox();
            this.cmdCreateFileName = new System.Windows.Forms.Button();
            this.tbPath = new System.Windows.Forms.TextBox();
            this.lnkPath = new System.Windows.Forms.LinkLabel();
            this.folderBrowserDialog1 = new System.Windows.Forms.FolderBrowserDialog();
            this.btnReadFileName = new System.Windows.Forms.Button();
            this.label1 = new System.Windows.Forms.Label();
            this.txtStart = new System.Windows.Forms.TextBox();
            this.txtStop = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.btnProcessCenterStudentUpdate = new System.Windows.Forms.Button();
            this.btnSessionStudents = new System.Windows.Forms.Button();
            this.SuspendLayout();
            // 
            // tbFileName
            // 
            this.tbFileName.Location = new System.Drawing.Point(67, 74);
            this.tbFileName.Name = "tbFileName";
            this.tbFileName.Size = new System.Drawing.Size(120, 22);
            this.tbFileName.TabIndex = 0;
            this.tbFileName.Text = "Bugs.xlsx";
            // 
            // cmdCreateFileName
            // 
            this.cmdCreateFileName.Location = new System.Drawing.Point(67, 102);
            this.cmdCreateFileName.Name = "cmdCreateFileName";
            this.cmdCreateFileName.Size = new System.Drawing.Size(188, 27);
            this.cmdCreateFileName.TabIndex = 1;
            this.cmdCreateFileName.Text = "Create FileName";
            this.cmdCreateFileName.Click += new System.EventHandler(this.cmdCreateFileName_Click);
            // 
            // tbPath
            // 
            this.tbPath.Location = new System.Drawing.Point(67, 46);
            this.tbPath.Name = "tbPath";
            this.tbPath.Size = new System.Drawing.Size(281, 22);
            this.tbPath.TabIndex = 5;
            this.tbPath.Text = "c:\\temp";
            // 
            // lnkPath
            // 
            this.lnkPath.Location = new System.Drawing.Point(10, 46);
            this.lnkPath.Name = "lnkPath";
            this.lnkPath.Size = new System.Drawing.Size(38, 28);
            this.lnkPath.TabIndex = 4;
            this.lnkPath.TabStop = true;
            this.lnkPath.Text = "Path";
            this.lnkPath.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
            this.lnkPath.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.lnkPath_LinkClicked);
            // 
            // btnReadFileName
            // 
            this.btnReadFileName.Location = new System.Drawing.Point(67, 218);
            this.btnReadFileName.Name = "btnReadFileName";
            this.btnReadFileName.Size = new System.Drawing.Size(188, 27);
            this.btnReadFileName.TabIndex = 6;
            this.btnReadFileName.Text = "Read FileName";
            this.btnReadFileName.UseVisualStyleBackColor = true;
            this.btnReadFileName.Click += new System.EventHandler(this.btnReadFileName_Click);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(68, 164);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(38, 17);
            this.label1.TabIndex = 7;
            this.label1.Text = "Start";
            // 
            // txtStart
            // 
            this.txtStart.Location = new System.Drawing.Point(120, 161);
            this.txtStart.Name = "txtStart";
            this.txtStart.Size = new System.Drawing.Size(100, 22);
            this.txtStart.TabIndex = 8;
            this.txtStart.Text = "A1";
            // 
            // txtStop
            // 
            this.txtStop.Location = new System.Drawing.Point(120, 189);
            this.txtStop.Name = "txtStop";
            this.txtStop.Size = new System.Drawing.Size(100, 22);
            this.txtStop.TabIndex = 10;
            this.txtStop.Text = "AF85";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(68, 192);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(37, 17);
            this.label2.TabIndex = 9;
            this.label2.Text = "Stop";
            // 
            // btnProcessCenterStudentUpdate
            // 
            this.btnProcessCenterStudentUpdate.Location = new System.Drawing.Point(71, 261);
            this.btnProcessCenterStudentUpdate.Name = "btnProcessCenterStudentUpdate";
            this.btnProcessCenterStudentUpdate.Size = new System.Drawing.Size(227, 23);
            this.btnProcessCenterStudentUpdate.TabIndex = 11;
            this.btnProcessCenterStudentUpdate.Text = "For Center Student Update";
            this.btnProcessCenterStudentUpdate.UseVisualStyleBackColor = true;
            this.btnProcessCenterStudentUpdate.Click += new System.EventHandler(this.btnProcessCenterStudentUpdate_Click);
            // 
            // btnSessionStudents
            // 
            this.btnSessionStudents.Location = new System.Drawing.Point(241, 161);
            this.btnSessionStudents.Name = "btnSessionStudents";
            this.btnSessionStudents.Size = new System.Drawing.Size(130, 23);
            this.btnSessionStudents.TabIndex = 12;
            this.btnSessionStudents.Text = "Session Students";
            this.btnSessionStudents.UseVisualStyleBackColor = true;
            this.btnSessionStudents.Click += new System.EventHandler(this.btnSessionStudents_Click);
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(6, 15);
            this.ClientSize = new System.Drawing.Size(400, 296);
            this.Controls.Add(this.btnSessionStudents);
            this.Controls.Add(this.btnProcessCenterStudentUpdate);
            this.Controls.Add(this.txtStop);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.txtStart);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.btnReadFileName);
            this.Controls.Add(this.tbPath);
            this.Controls.Add(this.lnkPath);
            this.Controls.Add(this.cmdCreateFileName);
            this.Controls.Add(this.tbFileName);
            this.Name = "Form1";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "Form1";
            this.ResumeLayout(false);
            this.PerformLayout();

        }
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

        private void cmdCreateFileName_Click(object sender, System.EventArgs e)
        {
            WriteFileDal wfd = new WriteFileDal();
            string sPath = tbPath.Text;
            string sFile = tbFileName.Text;
            string sPathFile = "";

            this.Cursor = Cursors.WaitCursor;

            sPathFile = sPath;
            
            if (sPathFile.EndsWith("\\") == false)
            {
                sPathFile += "\\";
            }

            sPathFile += sFile;
            if (sPathFile.Length>0)
            {
                wfd.CreateExcel(sPathFile);

                this.Cursor = Cursors.Default;
                
                MessageBox.Show("Finished Outputting Excel", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);
                
            }


        }

        private void lnkPath_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
        {
            string sPath = "";
            string sScrub = "";
            string sPathSelected = "";
            DialogResult result;

            sPath = tbPath.Text;
            if (sPath.Length == 0)
            {
                sPath = Application.StartupPath;
            }


            sScrub = sPath.Replace("bin\\Debug","");
            sPath = sScrub.Replace("bin\\Release","");
            folderBrowserDialog1.SelectedPath = sPath;
            result = folderBrowserDialog1.ShowDialog(this);
            
            if (result == DialogResult.OK)
            {
                sPathSelected = folderBrowserDialog1.SelectedPath;
                if (sPathSelected.Length > 0)
                {
                    if (sPathSelected.EndsWith("\\") == false)
                    {
                        sPathSelected += "\\";
                    }

                    tbPath.Text = sPathSelected;
                }
                else
                {
                    tbPath.Text = "";
                }

            }
            else
            {
                tbPath.Text = "";
            }
        }

        private void btnReadFileName_Click(object sender, EventArgs e)
        {
            string startCell = "";
            string stopCell = "";
            ReadFileDal rfd = new ReadFileDal();
            string sPath = tbPath.Text;
            string sFile = tbFileName.Text;
            string sPathFile = "";

            this.Cursor = Cursors.WaitCursor;

            sPathFile = sPath;

            if (sPathFile.EndsWith("\\") == false)
            {
                sPathFile += "\\";
            }

            sPathFile += sFile;
            if (sPathFile.Length > 0)
            {

                startCell = txtStart.Text.Trim();
                stopCell = txtStop.Text.Trim();

                rfd.ReadExcel(sPath,sPathFile, startCell, stopCell);

                this.Cursor = Cursors.Default;

                MessageBox.Show("Finished Reading Excel", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }


        }

        private void btnProcessCenterStudentUpdate_Click(object sender, EventArgs e)
        {

            string startCell = "";
            string stopCell = "";
            ReadFileDal rfd = new ReadFileDal();
            string sPath = tbPath.Text;
            string sFile = tbFileName.Text;
            string sPathFile = "";

            this.Cursor = Cursors.WaitCursor;

            sPathFile = sPath;

            if (sPathFile.EndsWith("\\") == false)
            {
                sPathFile += "\\";
            }

            sPathFile += sFile;
            if (sPathFile.Length > 0)
            {

                startCell = txtStart.Text.Trim();
                stopCell = txtStop.Text.Trim();

                rfd.ReadExcelForCenterStudentUpdate(sPath, sPathFile, startCell, stopCell);

                this.Cursor = Cursors.Default;

                MessageBox.Show("Finished Reading Excel", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }

        }

        private void btnSessionStudents_Click(object sender, EventArgs e)
        {

            string startCell = "";
            string stopCell = "";
            ReadFileDal rfd = new ReadFileDal();
            string sPath = tbPath.Text;
            string sFile = tbFileName.Text;
            string sPathFile = "";

            this.Cursor = Cursors.WaitCursor;

            sPathFile = sPath;

            if (sPathFile.EndsWith("\\") == false)
            {
                sPathFile += "\\";
            }

            sPathFile += sFile;
            if (sPathFile.Length > 0)
            {

                startCell = txtStart.Text.Trim();
                stopCell = txtStop.Text.Trim();

                rfd.ReadExcelForSessionStudentsUpdate(sPath, sPathFile, startCell, stopCell);

                this.Cursor = Cursors.Default;

                MessageBox.Show("Finished Reading Excel", "Completed", MessageBoxButtons.OK, MessageBoxIcon.Information);

            }
        }
	}
}
