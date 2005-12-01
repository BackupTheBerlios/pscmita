using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using System.Runtime.Remoting;
using System.Runtime.Remoting.Channels;
using System.Runtime.Remoting.Channels.Tcp;
using System.Diagnostics;
using System.Runtime.Serialization.Formatters;

using pscMitaRunner;

namespace CpscMitaAgent
{
	/// <summary>
	/// Summary description for Form1.
	/// </summary>
	public class mitaAgent : System.Windows.Forms.Form
	{
		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.NotifyIcon trayIcon;
		private System.Windows.Forms.ContextMenu cmTray;
		private System.Windows.Forms.MenuItem mnuShow;
		private System.Windows.Forms.MenuItem mnuExit;
		private System.Windows.Forms.Button btExit;
		private System.Windows.Forms.LinkLabel linkLabel1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.Label lblVersion;
		private System.ComponentModel.IContainer components;

		public mitaAgent()
		{
			try
			{
				InitializeComponent();
				System.Reflection.Assembly ass =  System.Reflection.Assembly.GetExecutingAssembly();
				System.Reflection.AssemblyName name = ass.GetName();
				this.lblVersion.Text = name.Version.Major + "." + name.Version.Minor + "." + name.Version.Revision;
				BinaryClientFormatterSinkProvider clientProvider = null;
				BinaryServerFormatterSinkProvider serverProvider = new BinaryServerFormatterSinkProvider();
				serverProvider.TypeFilterLevel = System.Runtime.Serialization.Formatters.TypeFilterLevel.Full;
				
				IDictionary props = new Hashtable();
				props["port"] = 7500;
				props["typeFilterLevel"] = TypeFilterLevel.Full;
				TcpChannel chan = new TcpChannel(
					props,clientProvider,serverProvider);

				ChannelServices.RegisterChannel(chan);

				//				string strConfig  = 
				//					Application.ExecutablePath + ".config";
				//				RemotingConfiguration.Configure(strConfig);
				RemotingConfiguration.RegisterWellKnownServiceType(
					typeof(pscMitaRunner.CpscMitaRunner),
					"pscMitaRunner",
					WellKnownObjectMode.Singleton);
				this.WindowState = FormWindowState.Minimized;
			}
			catch(Exception)
			{
				MessageBox.Show("Error!!!");
			}
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
			this.components = new System.ComponentModel.Container();
			System.Resources.ResourceManager resources = new System.Resources.ResourceManager(typeof(mitaAgent));
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.trayIcon = new System.Windows.Forms.NotifyIcon(this.components);
			this.cmTray = new System.Windows.Forms.ContextMenu();
			this.mnuShow = new System.Windows.Forms.MenuItem();
			this.mnuExit = new System.Windows.Forms.MenuItem();
			this.btExit = new System.Windows.Forms.Button();
			this.linkLabel1 = new System.Windows.Forms.LinkLabel();
			this.label2 = new System.Windows.Forms.Label();
			this.label3 = new System.Windows.Forms.Label();
			this.lblVersion = new System.Windows.Forms.Label();
			this.SuspendLayout();
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(8, 16);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(36, 36);
			this.pictureBox1.TabIndex = 0;
			this.pictureBox1.TabStop = false;
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Arial", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label1.Location = new System.Drawing.Point(60, 12);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(164, 16);
			this.label1.TabIndex = 1;
			this.label1.Text = "pscMita Process Agent";
			// 
			// trayIcon
			// 
			this.trayIcon.ContextMenu = this.cmTray;
			this.trayIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("trayIcon.Icon")));
			this.trayIcon.Text = "pscMita Process Agent";
			this.trayIcon.Visible = true;
			this.trayIcon.DoubleClick += new System.EventHandler(this.trayIcon_DoubleClick);
			// 
			// cmTray
			// 
			this.cmTray.MenuItems.AddRange(new System.Windows.Forms.MenuItem[] {
																																					 this.mnuShow,
																																					 this.mnuExit});
			// 
			// mnuShow
			// 
			this.mnuShow.DefaultItem = true;
			this.mnuShow.Index = 0;
			this.mnuShow.Text = "&Show Main Window";
			this.mnuShow.Click += new System.EventHandler(this.mnuShow_Click);
			// 
			// mnuExit
			// 
			this.mnuExit.Index = 1;
			this.mnuExit.Text = "E&xit";
			this.mnuExit.Click += new System.EventHandler(this.mnuExit_Click);
			// 
			// btExit
			// 
			this.btExit.Font = new System.Drawing.Font("Tahoma", 8.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.btExit.Location = new System.Drawing.Point(248, 64);
			this.btExit.Name = "btExit";
			this.btExit.Size = new System.Drawing.Size(76, 23);
			this.btExit.TabIndex = 2;
			this.btExit.Text = "&Done";
			this.btExit.Click += new System.EventHandler(this.btExit_Click);
			// 
			// linkLabel1
			// 
			this.linkLabel1.Location = new System.Drawing.Point(116, 72);
			this.linkLabel1.Name = "linkLabel1";
			this.linkLabel1.Size = new System.Drawing.Size(116, 16);
			this.linkLabel1.TabIndex = 4;
			this.linkLabel1.TabStop = true;
			this.linkLabel1.Text = "psc@my-tools4you.de";
			this.linkLabel1.LinkClicked += new System.Windows.Forms.LinkLabelLinkClickedEventHandler(this.linkLabel1_LinkClicked);
			// 
			// label2
			// 
			this.label2.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label2.Location = new System.Drawing.Point(60, 36);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(48, 20);
			this.label2.TabIndex = 5;
			this.label2.Text = "Version";
			// 
			// label3
			// 
			this.label3.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.label3.Location = new System.Drawing.Point(60, 72);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(52, 20);
			this.label3.TabIndex = 6;
			this.label3.Text = "Support";
			// 
			// lblVersion
			// 
			this.lblVersion.Font = new System.Drawing.Font("Microsoft Sans Serif", 8.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((System.Byte)(0)));
			this.lblVersion.Location = new System.Drawing.Point(116, 36);
			this.lblVersion.Name = "lblVersion";
			this.lblVersion.Size = new System.Drawing.Size(100, 20);
			this.lblVersion.TabIndex = 7;
			// 
			// mitaAgent
			// 
			this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
			this.ClientSize = new System.Drawing.Size(330, 95);
			this.ControlBox = false;
			this.Controls.Add(this.lblVersion);
			this.Controls.Add(this.label3);
			this.Controls.Add(this.label2);
			this.Controls.Add(this.linkLabel1);
			this.Controls.Add(this.btExit);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox1);
			this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.MaximizeBox = false;
			this.Name = "mitaAgent";
			this.ShowInTaskbar = false;
			this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
			this.Text = "pscMita Process Agent";
			this.TopMost = true;
			this.LocationChanged += new System.EventHandler(this.CServer_LocationChanged);
			this.ResumeLayout(false);

		}
		#endregion

		/// <summary>
		/// The main entry point for the application.
		/// </summary>
		[STAThread]
		static void Main() 
			{
			Application.Run(new mitaAgent());
		}

		private void mnuShow_Click(object sender, System.EventArgs e)
		{
			this.Show();
			this.WindowState = FormWindowState.Normal;
		}

		private void mnuExit_Click(object sender, System.EventArgs e)
		{
			this.Close();
		}

		private void CServer_LocationChanged(object sender, System.EventArgs e)
		{
			if (this.WindowState == FormWindowState.Minimized)
			{
				this.Hide();
			}
		}

		private void btExit_Click(object sender, System.EventArgs e)
		{
			this.WindowState = FormWindowState.Minimized;
		}

		private void linkLabel1_LinkClicked(object sender, System.Windows.Forms.LinkLabelLinkClickedEventArgs e)
		{
			try
			{
				Process proc = new Process();
				proc.StartInfo.FileName = 
					"mailto:psc@exec-se.de";
				proc.Start();
			}
			catch(Exception)
			{
				MessageBox.Show(this,"Cannot send mail because there is no default mail client.",
					"Can't send mail",
					MessageBoxButtons.OK,
					MessageBoxIcon.Information);
			}

		}

		private void trayIcon_DoubleClick(object sender, System.EventArgs e)
		{
			this.mnuShow_Click(this,null);
		}

	}
}
