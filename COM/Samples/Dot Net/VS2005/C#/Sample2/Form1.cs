using System;
using System.Collections;
using System.Drawing;
using System.Drawing.Printing;
using System.Windows.Forms;

namespace Sample2
{
	public class Form1 : System.Windows.Forms.Form
	{
		private PDFCreator.clsPDFCreator _PDFCreator;
		private PDFCreator.clsPDFCreatorError pErr;

		private PrintDocument pd;
		
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.Button button2;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Timer timer1;
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
			this.textBox1 = new System.Windows.Forms.TextBox();
			this.timer1 = new System.Windows.Forms.Timer(this.components);
			this.SuspendLayout();
			// 
			// button1
			// 
			this.button1.Enabled = false;
			this.button1.Location = new System.Drawing.Point(16, 8);
			this.button1.Name = "button1";
			this.button1.Size = new System.Drawing.Size(152, 40);
			this.button1.TabIndex = 0;
			this.button1.Text = "&Start";
			this.button1.Click += new System.EventHandler(this.button1_Click);
			// 
			// button2
			// 
			this.button2.Enabled = false;
			this.button2.Location = new System.Drawing.Point(288, 8);
			this.button2.Name = "button2";
			this.button2.Size = new System.Drawing.Size(152, 40);
			this.button2.TabIndex = 1;
			this.button2.Text = "&Preview";
			this.button2.Click += new System.EventHandler(this.button2_Click);
			// 
			// textBox1
			// 
			this.textBox1.BackColor = System.Drawing.Color.FromArgb(((System.Byte)(255)), ((System.Byte)(255)), ((System.Byte)(192)));
			this.textBox1.Location = new System.Drawing.Point(0, 56);
			this.textBox1.Multiline = true;
			this.textBox1.Name = "textBox1";
			this.textBox1.ReadOnly = true;
			this.textBox1.ScrollBars = System.Windows.Forms.ScrollBars.Both;
			this.textBox1.Size = new System.Drawing.Size(464, 80);
			this.textBox1.TabIndex = 2;
			this.textBox1.Text = "textBox1";
			this.textBox1.WordWrap = false;
			// 
			// Form1
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(466, 142);
			this.Controls.Add(this.button1);
			this.Controls.Add(this.button2);
			this.Controls.Add(this.textBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.MaximizeBox = false;
			this.MinimizeBox = false;
			this.Name = "Form1";
			this.Text = "Sample2 - PDFCreator COM interface";
			this.Closing += new System.ComponentModel.CancelEventHandler(this.Form1_Closing);
			this.Load += new System.EventHandler(this.Form1_Load);
			this.ResumeLayout(false);

		}
		#endregion

		[STAThread]
		static void Main() 
		{
			Application.Run(new Form1());
		}

		private void AddStatus(string Str1, bool ClearStatus)
		{
			if (ClearStatus)
			{
				textBox1.Text = Str1;
				textBox1.SelectionStart = 0;
			}
			else
			{
				if (textBox1.Text.Length == 0)
				{
					textBox1.Text = Str1;
					textBox1.SelectionStart = 0;
				}
				else
				{
					textBox1.Text = textBox1.Text + "\r\n" + Str1;
				}
			}
		}
 
		private void Form1_Load(object sender, System.EventArgs e)
		{
			string parameters;
			AddStatus("Status: Program is started.", true);

			pErr = new PDFCreator.clsPDFCreatorError();

			_PDFCreator = new PDFCreator.clsPDFCreator();
			_PDFCreator.eError  += new PDFCreator.__clsPDFCreator_eErrorEventHandler(_PDFCreator_eError); 
			_PDFCreator.eReady  += new PDFCreator.__clsPDFCreator_eReadyEventHandler(_PDFCreator_eReady); 
			
			parameters = "/NoProcessingAtStartup";

			if (_PDFCreator.cStart(parameters, false))
			{
				button1.Enabled = true;
				button2.Enabled = true;
				_PDFCreator.cClearCache();
				_PDFCreator.set_cOption("UseAutosave", 0);
				_PDFCreator.cPrinterStop = false;
			}
		}

		private void _PDFCreator_eReady()
		{
			AddStatus("Status: \"" + _PDFCreator.cOutputFilename + "\" was created!", false);
			_PDFCreator.cPrinterStop = true;
		}

		private void _PDFCreator_eError()
		{
			pErr = _PDFCreator.cError;
			AddStatus("Status: Error[" + pErr.Number + "]: " + pErr.Description, false);
		}

		private void button1_Click(object sender, System.EventArgs e)
		{
			pd = new PrintDocument();
			pd.PrintPage += new PrintPageEventHandler(pd_PrintPage);
			pd.PrinterSettings.PrinterName = "PDFCreator";
			pd.DocumentName = "PDFCreator Dot Net - Sample2";
			pd.Print();
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
			pd = new PrintDocument();
			pd.PrintPage += new PrintPageEventHandler(this.pd_PrintPage);
			pd.PrinterSettings.PrinterName = "PDFCreator";
			pd.DocumentName = "PDFCreator Dot Net - Sample2";
			PrintPreviewDialog ppdlg = new PrintPreviewDialog();
			ppdlg.Document = pd;
			ppdlg.WindowState = FormWindowState.Maximized;
			ppdlg.ShowDialog();
			ppdlg.Dispose();
		}

		private void pd_PrintPage(object sender, PrintPageEventArgs ev)
		{
			float x, y, r;
			x = pd.PrinterSettings.DefaultPageSettings.PaperSize.Width / 2; 
			y = pd.PrinterSettings.DefaultPageSettings.PaperSize.Height / 2; 
			r = pd.PrinterSettings.DefaultPageSettings.PaperSize.Width / 4; 
			DrawCircles(x - r / 2, y - r / 2, r, 5, ev);
		}
		private void DrawCircles(float x, float y, float r, long rec, PrintPageEventArgs ev)
		{
			if (rec != 0)
			{
				Pen p = new Pen(Color.Red);
				ev.Graphics.DrawString("PDFCreator", new Font("Arial", 16), Brushes.Black, 100, 100, new StringFormat());
				ev.Graphics.DrawEllipse(p, x - r, y, r, r);
				ev.Graphics.DrawEllipse(p, x + r, y, r, r);
				ev.Graphics.DrawEllipse(p, x, y - r, r, r);
				ev.Graphics.DrawEllipse(p, x, y + r, r, r);
				ev.Graphics.DrawEllipse(p, x, y, r, r);
				DrawCircles(x - r / 2, y - r / 2, r / 2, rec - 1, ev);
				DrawCircles(x - r / 2, y + r, r / 2, rec - 1, ev);
				DrawCircles(x + r, y - r / 2, r / 2, rec - 1, ev);
				DrawCircles(x + r, y + r, r / 2, rec - 1, ev);
			}
		}

		private void Form1_Closing(object sender, System.ComponentModel.CancelEventArgs e)
		{
			_PDFCreator.cClose();
            while (_PDFCreator.cProgramIsRunning)
            {
                Application.DoEvents();
                System.Threading.Thread.Sleep(100);
            }
			pErr = null;
            _PDFCreator = null;
			GC.Collect();
            GC.WaitForPendingFinalizers();
        }
	}
}