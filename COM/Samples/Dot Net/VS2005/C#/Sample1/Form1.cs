using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.IO;

namespace Sample1
{
	public class Form1 : System.Windows.Forms.Form
	{
		private const int maxTime  = 20;

		private PDFCreator.clsPDFCreator _PDFCreator;
		private PDFCreator.clsPDFCreatorError pErr;

        private bool ReadyState;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.Button button5;
		private System.Windows.Forms.Button button6;
		private System.Windows.Forms.Timer timer1;
		private System.Windows.Forms.OpenFileDialog openFileDialog1;
        private StatusStrip statusStrip1;
        private ToolStripStatusLabel toolStripStatusLabel1;
		private System.ComponentModel.IContainer components;
	
		public Form1()
		{
			InitializeComponent();
		}
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

		#region Vom Windows Form-Designer generierter Code
		private void InitializeComponent()
		{
            this.components = new System.ComponentModel.Container();
            this.button1 = new System.Windows.Forms.Button();
            this.button2 = new System.Windows.Forms.Button();
            this.button3 = new System.Windows.Forms.Button();
            this.button4 = new System.Windows.Forms.Button();
            this.button5 = new System.Windows.Forms.Button();
            this.button6 = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.openFileDialog1 = new System.Windows.Forms.OpenFileDialog();
            this.statusStrip1 = new System.Windows.Forms.StatusStrip();
            this.toolStripStatusLabel1 = new System.Windows.Forms.ToolStripStatusLabel();
            this.statusStrip1.SuspendLayout();
            this.SuspendLayout();
            // 
            // button1
            // 
            this.button1.Enabled = false;
            this.button1.Location = new System.Drawing.Point(8, 8);
            this.button1.Name = "button1";
            this.button1.Size = new System.Drawing.Size(152, 40);
            this.button1.TabIndex = 0;
            this.button1.Text = "Show options dialog";
            this.button1.Click += new System.EventHandler(this.button1_Click);
            // 
            // button2
            // 
            this.button2.Enabled = false;
            this.button2.Location = new System.Drawing.Point(8, 56);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(152, 40);
            this.button2.TabIndex = 1;
            this.button2.Text = "Show logfile dialog";
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // button3
            // 
            this.button3.Enabled = false;
            this.button3.Location = new System.Drawing.Point(192, 8);
            this.button3.Name = "button3";
            this.button3.Size = new System.Drawing.Size(152, 40);
            this.button3.TabIndex = 2;
            this.button3.Text = "Print printer testpage";
            this.button3.Click += new System.EventHandler(this.button3_Click);
            // 
            // button4
            // 
            this.button4.Enabled = false;
            this.button4.Location = new System.Drawing.Point(192, 56);
            this.button4.Name = "button4";
            this.button4.Size = new System.Drawing.Size(152, 40);
            this.button4.TabIndex = 3;
            this.button4.Text = "Print PDFCreator testpage";
            this.button4.Click += new System.EventHandler(this.button4_Click);
            // 
            // button5
            // 
            this.button5.Enabled = false;
            this.button5.Location = new System.Drawing.Point(376, 8);
            this.button5.Name = "button5";
            this.button5.Size = new System.Drawing.Size(152, 40);
            this.button5.TabIndex = 4;
            this.button5.Text = "Convert to PDF";
            this.button5.Click += new System.EventHandler(this.button5_Click);
            // 
            // button6
            // 
            this.button6.Enabled = false;
            this.button6.Location = new System.Drawing.Point(376, 56);
            this.button6.Name = "button6";
            this.button6.Size = new System.Drawing.Size(152, 40);
            this.button6.TabIndex = 5;
            this.button6.Text = "Convert to TIFF";
            this.button6.Click += new System.EventHandler(this.button6_Click);
            // 
            // timer1
            // 
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // statusStrip1
            // 
            this.statusStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.toolStripStatusLabel1});
            this.statusStrip1.Location = new System.Drawing.Point(0, 112);
            this.statusStrip1.Name = "statusStrip1";
            this.statusStrip1.Size = new System.Drawing.Size(536, 22);
            this.statusStrip1.TabIndex = 6;
            this.statusStrip1.Text = "statusStrip1";
            // 
            // toolStripStatusLabel1
            // 
            this.toolStripStatusLabel1.Name = "toolStripStatusLabel1";
            this.toolStripStatusLabel1.Size = new System.Drawing.Size(42, 17);
            this.toolStripStatusLabel1.Text = "Status:";
            // 
            // Form1
            // 
            this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
            this.ClientSize = new System.Drawing.Size(536, 134);
            this.Controls.Add(this.statusStrip1);
            this.Controls.Add(this.button1);
            this.Controls.Add(this.button2);
            this.Controls.Add(this.button3);
            this.Controls.Add(this.button4);
            this.Controls.Add(this.button5);
            this.Controls.Add(this.button6);
            this.Name = "Form1";
            this.Text = "Sample1 - PDFCreator COM interface";
            this.Closing += new System.ComponentModel.CancelEventHandler(this.Form1_Closing);
            this.Load += new System.EventHandler(this.Form1_Load);
            this.statusStrip1.ResumeLayout(false);
            this.statusStrip1.PerformLayout();
            this.ResumeLayout(false);
            this.PerformLayout();

		}
		#endregion

		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			string parameters;
            toolStripStatusLabel1.Text = "Status: Program is started.";

			pErr = new PDFCreator.clsPDFCreatorError();

			_PDFCreator = new PDFCreator.clsPDFCreator();
			_PDFCreator.eError += new PDFCreator.__clsPDFCreator_eErrorEventHandler(_PDFCreator_eError); 
			_PDFCreator.eReady += new PDFCreator.__clsPDFCreator_eReadyEventHandler(_PDFCreator_eReady); 
        
			parameters = "/NoProcessingAtStartup";

			if (!_PDFCreator.cStart(parameters, false))
			{
                toolStripStatusLabel1.Text = "Status: Error[" + pErr.Number + "]: " + pErr.Description;
            }
			else
			{
				button1.Enabled = true;
				button2.Enabled = true;
				button3.Enabled = true;
				button4.Enabled = true;
				button5.Enabled = true;
				button6.Enabled = true;
			}
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
            _PDFCreator.cShowOptionsDialog(true);
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			_PDFCreator.cShowLogfileDialog(true);
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			_PDFCreator.cDefaultPrinter = "PDFCreator";
			// Wait 1 second
			timer1.Interval = 1000;
			timer1.Enabled = true;
			while (timer1.Enabled)
			{
				Application.DoEvents();
			}
			_PDFCreator.cPrinterStop = false;
			_PDFCreator.cPrintPrinterTestpage("");
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
			_PDFCreator.cPrintPDFCreatorTestpage();
			_PDFCreator.cPrinterStop = true;
		}

		private void button5_Click(object sender, System.EventArgs e)
		{
	        PrintIt(0);
		}

		private void button6_Click(object sender, System.EventArgs e)
		{
	        PrintIt(0);
		}

		private void PrintIt(int FileTyp)
		{
			string fname, DefaultPrinter;
			FileInfo fi;
			PDFCreator.clsPDFCreatorOptions opt;
			openFileDialog1.Multiselect = false;
			openFileDialog1.CheckFileExists = true;
			openFileDialog1.CheckPathExists = true;
			openFileDialog1.Filter = "Text files (*.txt)|*.txt|All files (*.*)|*.*";
			if (openFileDialog1.ShowDialog() == DialogResult.OK)
			{
				fi = new FileInfo(openFileDialog1.FileName);
				if (fi.Name.Length > 0)
				{
					if (fi.Name.IndexOf(".") > 1)
					{
						fname = fi.Name.Substring(0,fi.Name.IndexOf("."));
					}
					else
					{
						fname = fi.Name;
					}
					if (!_PDFCreator.cIsPrintable(fi.FullName))
					{
						MessageBox.Show("File '" + fi.FullName + "' is not printable!", this.Text,MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
						return;
					}
					opt = _PDFCreator.cOptions;
					opt.UseAutosave = 1;
					opt.UseAutosaveDirectory = 1;
					opt.AutosaveDirectory = fi.DirectoryName;
					opt.AutosaveFormat = FileTyp;
					if (FileTyp == 5)
					{
						opt.BitmapResolution = 72;
					}
					opt.AutosaveFilename = fname;
					_PDFCreator.cOptions = opt;
					_PDFCreator.cClearCache();
					DefaultPrinter = _PDFCreator.cDefaultPrinter;
					_PDFCreator.cDefaultPrinter = "PDFCreator";
					_PDFCreator.cPrintFile(fi.FullName);
                    ReadyState = false;
                    _PDFCreator.cPrinterStop = false;
					timer1.Interval = maxTime * 1000;
					timer1.Enabled = true;
					while (!ReadyState && timer1.Enabled)
					{
						Application.DoEvents();
					}
					timer1.Enabled = false;
					if (!ReadyState)
					{
						MessageBox.Show("Creating printer test page as pdf.\n\r\n\r" + 
							"An error is occured: Time is up!", this.Text,MessageBoxButtons.OK, MessageBoxIcon.Exclamation);
					}
					_PDFCreator.cPrinterStop = true;
					_PDFCreator.cDefaultPrinter = DefaultPrinter;
				}
			}
		}

		private void _PDFCreator_eReady()
		{
            toolStripStatusLabel1.Text = "Status: \"" + _PDFCreator.cOutputFilename + "\" was created!"; 
            _PDFCreator.cPrinterStop = true;
			ReadyState = true;
		}

		private void _PDFCreator_eError()
		{
			pErr = _PDFCreator.cError;
		}

		private void timer1_Tick(object sender, System.EventArgs e)
		{
			timer1.Enabled=false;
		}

		private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			_PDFCreator.cClose();
			System.Runtime.InteropServices.Marshal.ReleaseComObject(_PDFCreator);
			System.Runtime.InteropServices.Marshal.ReleaseComObject(pErr);
			pErr = null;
			GC.Collect();
            GC.WaitForPendingFinalizers();
            GC.Collect();
            GC.WaitForPendingFinalizers();
        }
	}
}
