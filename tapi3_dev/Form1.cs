using System;
using System.Drawing;
using System.Collections;
using System.ComponentModel;
using System.Windows.Forms;
using System.Data;
using TAPI3Lib;
using System.IO;
using System.Collections.Generic;
using Newtonsoft.Json;
using System.Configuration;

namespace tapi3_dev
{


  /// <summary>
  /// Summary description for Form1.
  /// </summary>
  public class Form1 : System.Windows.Forms.Form
	{
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.ListBox listBox1;
		private System.Windows.Forms.Button button1;
		private System.Windows.Forms.TextBox textBox1;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.Button button2;
		private TAPIClass tobj;
		private ITAddress[] ia=new TAPI3Lib.ITAddress[10];
		private ITBasicCallControl bcc;
		private callnotification cn;
		private bool h323,reject;
		uint lines;
		int line;
		int[] registertoken=new int[10];
		private System.Windows.Forms.Button button3;
		private System.Windows.Forms.Button button4;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.CheckBox checkBox1;
		private System.Windows.Forms.CheckBox checkBox2;
		private System.Windows.Forms.Button button5;
		private System.Windows.Forms.TextBox textBox2;
		private System.Windows.Forms.Label label3;
    private NotifyIcon notifyIcon;
    private ContextMenuStrip contextMenuStrip1;
    private ToolStripMenuItem tEstToolStripMenuItem;
    private IContainer components;
    List<classHistory> lstHistory = new List<classHistory>();
    private TabControl tabControl;
    private TabPage tabPage1;
    private Button button6;
    private Label label1;
    private ComboBox comboBox1;
    private TabPage tabPage2;
    private DataGridView dataGridView1;
    private ContextMenuStrip cMenuGrid;
    private ToolStripMenuItem anrufenToolStripMenuItem;
    private ToolStripMenuItem löschenToolStripMenuItem;
    private ToolStripMenuItem beendenToolStripMenuItem;
    string destPath = Path.Combine(AppDomain.CurrentDomain.BaseDirectory, "history.txt");

    public Form1()
		{
			//
			// Required for Windows Form Designer support
			//
			InitializeComponent();
			initializetapi3();
			h323=true;
			reject=false;
      notifyIcon.Visible = true;
      DoDeserialize();
      string mydefaultLine = ConfigurationManager.AppSettings["defaultLine"].ToString();
      if (mydefaultLine != "")
      {
        int myLine = comboBox1.Items.IndexOf(mydefaultLine);
        if (myLine >= 0)
        {
          line = myLine;
          comboBox1.SelectedIndex = myLine;
          Hide();
          this.WindowState = FormWindowState.Minimized;
          //ShowInTaskbar = false;
          RegisterLine();
        }
      }

      //MessageBox.Show("lines : "+lines,"Lines avaialble are");
      //
      // TODO: Add any constructor code after InitializeComponent call
      //
    }
		void initializetapi3()
		{
			try
			{
				tobj = new TAPIClass();
				tobj.Initialize();
				IEnumAddress ea=tobj.EnumerateAddresses();
				ITAddress ln;
				uint arg3=0;
				lines=0;
			
				cn=new callnotification();
				cn.addtolist=new callnotification.listshow(this.status);
        cn.addToCallNotify = new callnotification.addCallNotify(this.addCallNotify);
        cn.addToHistory = new callnotification.addCallHistory(this.addCallHistory);
     
        tobj.ITTAPIEventNotification_Event_Event+= new TAPI3Lib.ITTAPIEventNotification_EventEventHandler(cn.Event);
				tobj.EventFilter=(int)(TAPI_EVENT.TE_CALLNOTIFICATION|
					TAPI_EVENT.TE_DIGITEVENT|
					TAPI_EVENT.TE_PHONEEVENT|
					TAPI_EVENT.TE_CALLSTATE|
					TAPI_EVENT.TE_GENERATEEVENT|
					TAPI_EVENT.TE_GATHERDIGITS|
					TAPI_EVENT.TE_REQUEST);
			
				for(int i=0;i<10;i++)
				{
					ea.Next(1,out ln,ref arg3);
					ia[i]=ln;
					if(ln!=null)
					{
						comboBox1.Items.Add(ia[i].AddressName);
						lines++;
					}
					else
						break;
				}

        //tobj.ITTAPIEventNotification_Event_Event+= new TAPI3Lib.ITTAPIEventNotification_EventEventHandler(cn.Event);
        //tobj.EventFilter=(int)(TAPI_EVENT.TE_CALLNOTIFICATION|TAPI_EVENT.TE_DIGITEVENT|TAPI_EVENT.TE_PHONEEVENT|TAPI_EVENT.TE_CALLSTATE);
        //registertoken=tobj.RegisterCallNotifications(ia[6],true,true,TapiConstants.TAPIMEDIATYPE_AUDIO|TapiConstants.TAPIMEDIATYPE_DATAMODEM,1);	
        //MessageBox.Show("Registration token :-"+registertoken,"Regitration complete");

      }
			catch(Exception e)
			{
				MessageBox.Show(e.ToString());
			}
		}
		public void status(string str)
		{
			listBox1.Items.Add(str);
		}
		
		/// <summary>
		/// Clean up any resources being used.
		/// </summary>
		protected override void Dispose( bool disposing )
		{
			tobj.Shutdown();
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
      System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(Form1));
      this.groupBox1 = new System.Windows.Forms.GroupBox();
      this.button1 = new System.Windows.Forms.Button();
      this.listBox1 = new System.Windows.Forms.ListBox();
      this.textBox1 = new System.Windows.Forms.TextBox();
      this.label2 = new System.Windows.Forms.Label();
      this.button2 = new System.Windows.Forms.Button();
      this.button3 = new System.Windows.Forms.Button();
      this.button4 = new System.Windows.Forms.Button();
      this.groupBox2 = new System.Windows.Forms.GroupBox();
      this.checkBox2 = new System.Windows.Forms.CheckBox();
      this.checkBox1 = new System.Windows.Forms.CheckBox();
      this.button5 = new System.Windows.Forms.Button();
      this.textBox2 = new System.Windows.Forms.TextBox();
      this.label3 = new System.Windows.Forms.Label();
      this.notifyIcon = new System.Windows.Forms.NotifyIcon(this.components);
      this.contextMenuStrip1 = new System.Windows.Forms.ContextMenuStrip(this.components);
      this.tEstToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.beendenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.tabControl = new System.Windows.Forms.TabControl();
      this.tabPage2 = new System.Windows.Forms.TabPage();
      this.dataGridView1 = new System.Windows.Forms.DataGridView();
      this.cMenuGrid = new System.Windows.Forms.ContextMenuStrip(this.components);
      this.anrufenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.löschenToolStripMenuItem = new System.Windows.Forms.ToolStripMenuItem();
      this.tabPage1 = new System.Windows.Forms.TabPage();
      this.button6 = new System.Windows.Forms.Button();
      this.label1 = new System.Windows.Forms.Label();
      this.comboBox1 = new System.Windows.Forms.ComboBox();
      this.groupBox1.SuspendLayout();
      this.groupBox2.SuspendLayout();
      this.contextMenuStrip1.SuspendLayout();
      this.tabControl.SuspendLayout();
      this.tabPage2.SuspendLayout();
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).BeginInit();
      this.cMenuGrid.SuspendLayout();
      this.tabPage1.SuspendLayout();
      this.SuspendLayout();
      // 
      // groupBox1
      // 
      this.groupBox1.Controls.Add(this.button1);
      this.groupBox1.Controls.Add(this.listBox1);
      this.groupBox1.Location = new System.Drawing.Point(59, 45);
      this.groupBox1.Name = "groupBox1";
      this.groupBox1.Size = new System.Drawing.Size(408, 160);
      this.groupBox1.TabIndex = 5;
      this.groupBox1.TabStop = false;
      this.groupBox1.Text = "Call status";
      // 
      // button1
      // 
      this.button1.Location = new System.Drawing.Point(296, 120);
      this.button1.Name = "button1";
      this.button1.Size = new System.Drawing.Size(75, 23);
      this.button1.TabIndex = 1;
      this.button1.Text = "Clear status";
      this.button1.Click += new System.EventHandler(this.button1_Click);
      // 
      // listBox1
      // 
      this.listBox1.Location = new System.Drawing.Point(32, 24);
      this.listBox1.Name = "listBox1";
      this.listBox1.Size = new System.Drawing.Size(344, 82);
      this.listBox1.TabIndex = 0;
      // 
      // textBox1
      // 
      this.textBox1.Location = new System.Drawing.Point(58, 234);
      this.textBox1.Name = "textBox1";
      this.textBox1.Size = new System.Drawing.Size(152, 20);
      this.textBox1.TabIndex = 6;
      this.textBox1.KeyDown += new System.Windows.Forms.KeyEventHandler(this.textBox1_KeyDown);
      // 
      // label2
      // 
      this.label2.Location = new System.Drawing.Point(58, 210);
      this.label2.Name = "label2";
      this.label2.Size = new System.Drawing.Size(80, 16);
      this.label2.TabIndex = 7;
      this.label2.Text = "Call number/IP";
      // 
      // button2
      // 
      this.button2.Location = new System.Drawing.Point(378, 234);
      this.button2.Name = "button2";
      this.button2.Size = new System.Drawing.Size(75, 23);
      this.button2.TabIndex = 8;
      this.button2.Text = "CALL";
      this.button2.Click += new System.EventHandler(this.button2_Click);
      // 
      // button3
      // 
      this.button3.Location = new System.Drawing.Point(483, 182);
      this.button3.Name = "button3";
      this.button3.Size = new System.Drawing.Size(75, 23);
      this.button3.TabIndex = 9;
      this.button3.Text = "Answer";
      this.button3.Click += new System.EventHandler(this.button3_Click);
      // 
      // button4
      // 
      this.button4.Location = new System.Drawing.Point(483, 234);
      this.button4.Name = "button4";
      this.button4.Size = new System.Drawing.Size(75, 23);
      this.button4.TabIndex = 11;
      this.button4.Text = "Disconnect";
      this.button4.Click += new System.EventHandler(this.button4_Click);
      // 
      // groupBox2
      // 
      this.groupBox2.Controls.Add(this.checkBox2);
      this.groupBox2.Location = new System.Drawing.Point(483, 112);
      this.groupBox2.Name = "groupBox2";
      this.groupBox2.Size = new System.Drawing.Size(104, 64);
      this.groupBox2.TabIndex = 12;
      this.groupBox2.TabStop = false;
      this.groupBox2.Text = "Answer mode";
      // 
      // checkBox2
      // 
      this.checkBox2.Location = new System.Drawing.Point(16, 24);
      this.checkBox2.Name = "checkBox2";
      this.checkBox2.Size = new System.Drawing.Size(80, 24);
      this.checkBox2.TabIndex = 0;
      this.checkBox2.Text = "Reject";
      this.checkBox2.CheckedChanged += new System.EventHandler(this.checkBox2_CheckedChanged);
      // 
      // checkBox1
      // 
      this.checkBox1.Checked = true;
      this.checkBox1.CheckState = System.Windows.Forms.CheckState.Checked;
      this.checkBox1.Location = new System.Drawing.Point(242, 234);
      this.checkBox1.Name = "checkBox1";
      this.checkBox1.Size = new System.Drawing.Size(112, 16);
      this.checkBox1.TabIndex = 13;
      this.checkBox1.Text = "H.323 call(IP call)";
      this.checkBox1.CheckedChanged += new System.EventHandler(this.checkBox1_CheckedChanged);
      // 
      // button5
      // 
      this.button5.Location = new System.Drawing.Point(597, 69);
      this.button5.Name = "button5";
      this.button5.Size = new System.Drawing.Size(75, 23);
      this.button5.TabIndex = 14;
      this.button5.Text = "Transfer";
      this.button5.Click += new System.EventHandler(this.button5_Click);
      // 
      // textBox2
      // 
      this.textBox2.Location = new System.Drawing.Point(470, 69);
      this.textBox2.Name = "textBox2";
      this.textBox2.Size = new System.Drawing.Size(100, 20);
      this.textBox2.TabIndex = 15;
      // 
      // label3
      // 
      this.label3.Location = new System.Drawing.Point(470, 45);
      this.label3.Name = "label3";
      this.label3.Size = new System.Drawing.Size(88, 16);
      this.label3.TabIndex = 16;
      this.label3.Text = "tranfer address";
      // 
      // notifyIcon
      // 
      this.notifyIcon.ContextMenuStrip = this.contextMenuStrip1;
      this.notifyIcon.Icon = ((System.Drawing.Icon)(resources.GetObject("notifyIcon.Icon")));
      this.notifyIcon.Text = "Wildix";
      this.notifyIcon.Visible = true;
      this.notifyIcon.DoubleClick += new System.EventHandler(this.notifyIcon_DoubleClick);
      // 
      // contextMenuStrip1
      // 
      this.contextMenuStrip1.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.tEstToolStripMenuItem,
            this.beendenToolStripMenuItem});
      this.contextMenuStrip1.Name = "contextMenuStrip1";
      this.contextMenuStrip1.Size = new System.Drawing.Size(199, 48);
      // 
      // tEstToolStripMenuItem
      // 
      this.tEstToolStripMenuItem.Name = "tEstToolStripMenuItem";
      this.tEstToolStripMenuItem.Size = new System.Drawing.Size(198, 22);
      this.tEstToolStripMenuItem.Text = "Teste Benachrichtigung";
      this.tEstToolStripMenuItem.Click += new System.EventHandler(this.tEstToolStripMenuItem_Click);
      // 
      // beendenToolStripMenuItem
      // 
      this.beendenToolStripMenuItem.Name = "beendenToolStripMenuItem";
      this.beendenToolStripMenuItem.Size = new System.Drawing.Size(198, 22);
      this.beendenToolStripMenuItem.Text = "Beenden";
      this.beendenToolStripMenuItem.Click += new System.EventHandler(this.beendenToolStripMenuItem_Click);
      // 
      // tabControl
      // 
      this.tabControl.Controls.Add(this.tabPage2);
      this.tabControl.Controls.Add(this.tabPage1);
      this.tabControl.Location = new System.Drawing.Point(12, 17);
      this.tabControl.Name = "tabControl";
      this.tabControl.SelectedIndex = 0;
      this.tabControl.Size = new System.Drawing.Size(803, 350);
      this.tabControl.TabIndex = 2;
      // 
      // tabPage2
      // 
      this.tabPage2.Controls.Add(this.dataGridView1);
      this.tabPage2.Location = new System.Drawing.Point(4, 22);
      this.tabPage2.Name = "tabPage2";
      this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
      this.tabPage2.Size = new System.Drawing.Size(795, 324);
      this.tabPage2.TabIndex = 1;
      this.tabPage2.Text = "History";
      this.tabPage2.UseVisualStyleBackColor = true;
      // 
      // dataGridView1
      // 
      this.dataGridView1.AllowUserToOrderColumns = true;
      this.dataGridView1.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
      this.dataGridView1.ContextMenuStrip = this.cMenuGrid;
      this.dataGridView1.Dock = System.Windows.Forms.DockStyle.Fill;
      this.dataGridView1.Location = new System.Drawing.Point(3, 3);
      this.dataGridView1.MultiSelect = false;
      this.dataGridView1.Name = "dataGridView1";
      this.dataGridView1.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
      this.dataGridView1.Size = new System.Drawing.Size(789, 318);
      this.dataGridView1.TabIndex = 0;
      this.dataGridView1.CellMouseDown += new System.Windows.Forms.DataGridViewCellMouseEventHandler(this.dataGridView1_CellMouseDown);
      // 
      // cMenuGrid
      // 
      this.cMenuGrid.Items.AddRange(new System.Windows.Forms.ToolStripItem[] {
            this.anrufenToolStripMenuItem,
            this.löschenToolStripMenuItem});
      this.cMenuGrid.Name = "cMenuGrid";
      this.cMenuGrid.Size = new System.Drawing.Size(119, 48);
      // 
      // anrufenToolStripMenuItem
      // 
      this.anrufenToolStripMenuItem.Name = "anrufenToolStripMenuItem";
      this.anrufenToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
      this.anrufenToolStripMenuItem.Text = "Anrufen";
      this.anrufenToolStripMenuItem.Click += new System.EventHandler(this.anrufenToolStripMenuItem_Click);
      // 
      // löschenToolStripMenuItem
      // 
      this.löschenToolStripMenuItem.Name = "löschenToolStripMenuItem";
      this.löschenToolStripMenuItem.Size = new System.Drawing.Size(118, 22);
      this.löschenToolStripMenuItem.Text = "Löschen";
      this.löschenToolStripMenuItem.Click += new System.EventHandler(this.löschenToolStripMenuItem_Click);
      // 
      // tabPage1
      // 
      this.tabPage1.Controls.Add(this.button6);
      this.tabPage1.Controls.Add(this.label1);
      this.tabPage1.Controls.Add(this.checkBox1);
      this.tabPage1.Controls.Add(this.label3);
      this.tabPage1.Controls.Add(this.button4);
      this.tabPage1.Controls.Add(this.textBox2);
      this.tabPage1.Controls.Add(this.button3);
      this.tabPage1.Controls.Add(this.button2);
      this.tabPage1.Controls.Add(this.comboBox1);
      this.tabPage1.Controls.Add(this.label2);
      this.tabPage1.Controls.Add(this.button5);
      this.tabPage1.Controls.Add(this.textBox1);
      this.tabPage1.Controls.Add(this.groupBox2);
      this.tabPage1.Controls.Add(this.groupBox1);
      this.tabPage1.Location = new System.Drawing.Point(4, 22);
      this.tabPage1.Name = "tabPage1";
      this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
      this.tabPage1.Size = new System.Drawing.Size(795, 324);
      this.tabPage1.TabIndex = 0;
      this.tabPage1.Text = "Function";
      this.tabPage1.UseVisualStyleBackColor = true;
      // 
      // button6
      // 
      this.button6.Location = new System.Drawing.Point(499, 18);
      this.button6.Name = "button6";
      this.button6.Size = new System.Drawing.Size(75, 23);
      this.button6.TabIndex = 20;
      this.button6.Text = "Register";
      // 
      // label1
      // 
      this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
      this.label1.Location = new System.Drawing.Point(3, 26);
      this.label1.Name = "label1";
      this.label1.Size = new System.Drawing.Size(50, 23);
      this.label1.TabIndex = 19;
      this.label1.Text = "Line";
      // 
      // comboBox1
      // 
      this.comboBox1.Location = new System.Drawing.Point(59, 18);
      this.comboBox1.Name = "comboBox1";
      this.comboBox1.Size = new System.Drawing.Size(408, 21);
      this.comboBox1.TabIndex = 18;
      this.comboBox1.Text = "Select Line of communication";
      // 
      // Form1
      // 
      this.AutoScaleBaseSize = new System.Drawing.Size(5, 13);
      this.ClientSize = new System.Drawing.Size(885, 379);
      this.Controls.Add(this.tabControl);
      this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
      this.Name = "Form1";
      this.Text = "Wildix";
      this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.Form1_FormClosing);
      this.Load += new System.EventHandler(this.Form1_Load);
      this.Resize += new System.EventHandler(this.Form1_Resize);
      this.groupBox1.ResumeLayout(false);
      this.groupBox2.ResumeLayout(false);
      this.contextMenuStrip1.ResumeLayout(false);
      this.tabControl.ResumeLayout(false);
      this.tabPage2.ResumeLayout(false);
      ((System.ComponentModel.ISupportInitialize)(this.dataGridView1)).EndInit();
      this.cMenuGrid.ResumeLayout(false);
      this.tabPage1.ResumeLayout(false);
      this.tabPage1.PerformLayout();
      this.ResumeLayout(false);

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

		private void comboBox1_SelectedIndexChanged(object sender, System.EventArgs e)
		{
			line=comboBox1.SelectedIndex;
		}

		private void button2_Click(object sender, System.EventArgs e)
		{
      MakeCall(textBox1.Text);
		}

    private void MakeCall(String Number)
    {
      TAPI3Lib.ITAddress ln = null;
      ln = ia[line];
      if (Number.Length != 0)
      {
        cn.addtolist("Rufe " + Number + " an");			
        try
        {
         /* if (!h323)
          {
            bcc = ln.CreateCall(textBox1.Text, TapiConstants.LINEADDRESSTYPE_PHONENUMBER | TapiConstants.LINEADDRESSTYPE_IPADDRESS, TapiConstants.TAPIMEDIATYPE_DATAMODEM | TapiConstants.TAPIMEDIATYPE_AUDIO);
            bcc.SetQOS(TapiConstants.TAPIMEDIATYPE_DATAMODEM | TapiConstants.TAPIMEDIATYPE_AUDIO, QOS_SERVICE_LEVEL.QSL_BEST_EFFORT);
            bcc.Connect(false);
          }
          else
          {*/
            bcc = ln.CreateCall(Number, TapiConstants.LINEADDRESSTYPE_IPADDRESS, TapiConstants.TAPIMEDIATYPE_AUDIO);
            bcc.Connect(false);
          //}
        }
        catch 
        {
          MessageBox.Show("Failed to create call!", "TAPI3");
        }
      }
      else
      {
        MessageBox.Show("Please enter number to dial.. ");
      }
    }

		private void button1_Click(object sender, System.EventArgs e)
		{
			listBox1.Items.Clear();
		}

		private void button3_Click(object sender, System.EventArgs e)
		{
			IEnumCall ec = ia[line].EnumerateCalls();
			uint arg=0;
			ITCallInfo ici;
			try
			{
				ec.Next(1,out ici,ref arg);
				ITBasicCallControl bc=(TAPI3Lib.ITBasicCallControl)ici;
				if(!reject)
				{
					bc.Answer();
				}
				else
				{
					bc.Disconnect(DISCONNECT_CODE.DC_REJECTED);
					ici.ReleaseUserUserInfo();
				}
			}
			catch(Exception exp)
			{
				MessageBox.Show("There may not be any calls to answer! \n\n"+exp.ToString(),"TAPI3");
			}
		}

		private void button4_Click(object sender, System.EventArgs e)
		{
			IEnumCall ec = ia[line].EnumerateCalls();
			uint arg = 0;
			ITCallInfo ici;
			try
			{
				ec.Next(1,out ici,ref arg);
				ITBasicCallControl bc=(ITBasicCallControl)ici;
				bc.Disconnect(DISCONNECT_CODE.DC_NORMAL);
				ici.ReleaseUserUserInfo();
			}
			catch(Exception exp)
			{
				MessageBox.Show("No call to disconnect!","TAPI3");
			}
		}

		private void checkBox1_CheckedChanged(object sender, System.EventArgs e)
		{
			h323=checkBox1.Checked;
		}

		private void checkBox2_CheckedChanged(object sender, System.EventArgs e)
		{
			reject=checkBox2.Checked;
		}

		private void button5_Click(object sender, System.EventArgs e)
		{
			IEnumCall ec = ia[line].EnumerateCalls();
			uint arg = 0;
			ITCallInfo ici;
			try
			{
				ec.Next(1,out ici,ref arg);
				ITBasicCallControl bc=(ITBasicCallControl)ici;
				bc.BlindTransfer(textBox2.Text);
				ici.ReleaseUserUserInfo();
			}
			catch(Exception exp)
			{
				MessageBox.Show("May not have any call to disconnect!\n\n"+exp.ToString(),"TAPI3");
			}
		}

		private void Form1_Load(object sender, System.EventArgs e)
		{
			//MessageBox.Show("To recieve calls from any line you need to register on that line\n,you can do this by selecting the line ansd press the register button!","Instruction");
		}


    private void Form1_Resize(object sender, EventArgs e)
    {
      if (this.WindowState == FormWindowState.Minimized)
      {
        Hide();
        notifyIcon.Visible = true;
        ShowInTaskbar = true;
      }

    }

    private void button6_Click(object sender, System.EventArgs e)
		{
      RegisterLine();
		}

    private void RegisterLine()
    {
      try
      {
        registertoken[line] = tobj.RegisterCallNotifications(ia[line], true, true, TapiConstants.TAPIMEDIATYPE_AUDIO, 2);
        cn.addtolist("Registration token : " + registertoken[line]);
        addCallNotify("Registriert " + registertoken[line]);
        //MessageBox.Show("Registration token : "+registertoken[line],"Registration Succeed for line "+line);
      }
      catch (Exception ein)
      {
        MessageBox.Show("Failed to register on line " + line, "Registration for calls");
      }
    }

    private void tEstToolStripMenuItem_Click(object sender, EventArgs e)
    {
      addCallNotify("context");
      classHistory mycl = new classHistory();
      mycl.FromName = "test "+(DateTime.Now).ToShortTimeString();
      addCallHistory(mycl);
    }

    private void notifyIcon_DoubleClick(object sender, EventArgs e)
    {
      Show();
      this.WindowState = FormWindowState.Normal;
    }

    public void addCallNotify(string s)
    {
      notifyIcon.BalloonTipTitle = "Wildix";
      notifyIcon.BalloonTipIcon = ToolTipIcon.Info;
      notifyIcon.BalloonTipText = s;
      notifyIcon.ShowBalloonTip(3000);
    }

    private void addCallHistory(classHistory myHistory)
    {
      myHistory.StartTime = DateTime.Now;
      lstHistory.Insert(0, myHistory);
      DoSerialize();
    }

    private void DoSerialize()
    {
      string output = JsonConvert.SerializeObject(lstHistory);
      System.IO.File.WriteAllText(destPath, output);
      dataGridView1.DataSource = FillGrid(lstHistory);
    }

    private void DoDeserialize()
    {
      lstHistory = JsonConvert.DeserializeObject<List<classHistory>>(File.ReadAllText(destPath));
      dataGridView1.DataSource = FillGrid(lstHistory);
    }

    private void anrufenToolStripMenuItem_Click(object sender, EventArgs e)
    {
      if (dataGridView1.SelectedCells.Count > 0)
      {
        int selectedrowindex = dataGridView1.SelectedCells[0].RowIndex;

        DataGridViewRow selectedRow = dataGridView1.Rows[selectedrowindex];
        string direction = Convert.ToString(selectedRow.Cells["Direction"].Value);
        string numberToCall = "";
        if (direction == "Outgoing")
        {
          numberToCall = Convert.ToString(selectedRow.Cells["ToNumber"].Value);
        }
        else
        {
          numberToCall = Convert.ToString(selectedRow.Cells["FromNumber"].Value);
        }
        MakeCall(numberToCall);
      }
    }

    private void löschenToolStripMenuItem_Click(object sender, EventArgs e)
    {
      if (dataGridView1.SelectedCells.Count > 0)
      {
        int selectedrowindex = dataGridView1.SelectedCells[0].RowIndex;
        DataGridViewRow selectedRow = dataGridView1.Rows[selectedrowindex];
        for (int i = 0; i < lstHistory.Count; i++)
        {
          if
            (
              (lstHistory[i].StartTime == Convert.ToDateTime(selectedRow.Cells["StartTime"].Value)) &
              (lstHistory[i].Direction == Convert.ToString(selectedRow.Cells["Direction"].Value)) &
              (lstHistory[i].FromNumber == Convert.ToString(selectedRow.Cells["FromNumber"].Value)) &
              (lstHistory[i].ToNumber == Convert.ToString(selectedRow.Cells["ToNumber"].Value)) &
              (lstHistory[i].CallID == Convert.ToInt32(selectedRow.Cells["CallID"].Value))
            )
          {
            lstHistory.RemoveAt(i);
            DoSerialize();
            break;
          }
        }
      }
    }

    private void dataGridView1_CellMouseDown(object sender, DataGridViewCellMouseEventArgs e)
    {
      if (e.Button == MouseButtons.Right)
      {
        int rowSelected = e.RowIndex;
        if (e.RowIndex != -1)
        {
          this.dataGridView1.Rows[rowSelected].Selected = true;
        }
      }
    }

    private void Form1_FormClosing(object sender, FormClosingEventArgs e)
    {
      if (e.CloseReason == CloseReason.UserClosing)
      {
        e.Cancel = true;
        Hide();
        this.WindowState = FormWindowState.Minimized;
      }
    }

    private void beendenToolStripMenuItem_Click(object sender, EventArgs e)
    {
      Application.Exit();
    }

    private void textBox1_KeyDown(object sender, KeyEventArgs e)
    {
      if (e.KeyCode == Keys.Enter)
      {
        MakeCall(textBox1.Text);
      }
    }

    public static DataTable FillGrid<T>( IList<T> data)
    {
      PropertyDescriptorCollection props = TypeDescriptor.GetProperties(typeof(T));
      DataTable table = new DataTable();
      for (int i = 0; i < props.Count; i++)
      {
        PropertyDescriptor prop = props[i];
        table.Columns.Add(prop.Name, prop.PropertyType);
      }
      object[] values = new object[props.Count];
      foreach (T item in data)
      {
        for (int i = 0; i < values.Length; i++)
        {
          values[i] = props[i].GetValue(item);
        }
        try { table.Rows.Add(values); } catch { }
      }
      return table;
    }

  }
	class callnotification:TAPI3Lib.ITTAPIEventNotification
	{
		public delegate void listshow(string str);
    public delegate void addCallNotify(string s);
    public delegate void addCallHistory(classHistory myHistory);
    public listshow addtolist;
    public addCallNotify addToCallNotify;
    public addCallHistory addToHistory;
    

    public void Event(TAPI3Lib.TAPI_EVENT te,object eobj)
		{
      //ITCallStateEvent.Call.get_CallInfoLong(CALLINFO_LONG.CIL_CALLID); damit bekommt man vielleicht eine CallID
      classHistory clHistory = new classHistory();
      clHistory.Clear();
      TAPI3Lib.ITCallNotificationEvent cn = eobj as TAPI3Lib.ITCallNotificationEvent;
      switch (te)
			{
				case TAPI3Lib.TAPI_EVENT.TE_CALLNOTIFICATION:
					addtolist("call notification event has occured");
          //TAPI3Lib.ITCallNotificationEvent cn = eobj as TAPI3Lib.ITCallNotificationEvent;
          if (cn.Call.CallState == TAPI3Lib.CALL_STATE.CS_OFFERING)
          {
            string c = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER);
            addtolist("Call Offering: " + c + " -> " + cn.Call.Address.DialableAddress);

            //classHistory clHistory = new classHistory();
            clHistory.FromName = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME);
            clHistory.FromNumber = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER);
            clHistory.ToName = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME);
            clHistory.ToNumber = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNUMBER);
            clHistory.Direction = "Incoming";
            try { clHistory.CallID = cn.Call.get_CallInfoLong(CALLINFO_LONG.CIL_CALLID); } catch { }
            addToHistory(clHistory);

            addToCallNotify(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME) + " (" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER) + ")");
            addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME));//Anita
            //addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDPARTYFRIENDLYNAME));
            addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME));//EDV-Support
            addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER));//410
            //addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLINGPARTYID));
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNAME)); } catch { }
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNUMBER)); } catch { }
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_DISPLAYABLEADDRESS)); } catch { }
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNAME)); } catch { }
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNUMBER)); } catch { }
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNAME)); } catch { }
            try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNUMBER)); } catch { }
          }
          break;
				case TAPI3Lib.TAPI_EVENT.TE_DIGITEVENT:
					TAPI3Lib.ITDigitDetectionEvent dd=(TAPI3Lib.ITDigitDetectionEvent)eobj;
					addtolist("Dialed digit"+dd.ToString());
          break;
				case TAPI3Lib.TAPI_EVENT.TE_GENERATEEVENT:
					TAPI3Lib.ITDigitGenerationEvent dg=(TAPI3Lib.ITDigitGenerationEvent)eobj;
					MessageBox.Show("digit dialed!");
					addtolist("Dialed digit"+dg.ToString());
					break;
				case TAPI3Lib.TAPI_EVENT.TE_PHONEEVENT:
					addtolist("A phone event!");
					break;
				case TAPI3Lib.TAPI_EVENT.TE_GATHERDIGITS:
					addtolist("Gather digit event!");
					break;
				case TAPI3Lib.TAPI_EVENT.TE_CALLSTATE:
					TAPI3Lib.ITCallStateEvent a=(TAPI3Lib.ITCallStateEvent)eobj;
					TAPI3Lib.ITCallInfo b=a.Call;
				  switch(b.CallState)
				  {
					  case TAPI3Lib.CALL_STATE.CS_INPROGRESS:
						addtolist("dialing");
              try
              {
                //TAPI3Lib.ITCallStateEvent cn1 = eobj as TAPI3Lib.ITCallStateEvent;
                //cn1.Call.get_c
                addtolist("**********");
                try { addtolist("01"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNAME)); } catch { }
                try { addtolist("02"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNUMBER)); } catch { }
                try { addtolist("03"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_DISPLAYABLEADDRESS)); } catch { }
                try { addtolist("04"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNAME)); } catch { }
                try { addtolist("05"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNUMBER)); } catch { }
                try { addtolist("06"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNAME)); } catch { }
                try { addtolist("07"+ cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNUMBER)); } catch { }
                try { addtolist("08" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME)); } catch { }
                try { addtolist("09" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNUMBER)); } catch { }
                try { addtolist("30" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER)); } catch { }
                try { addtolist("31" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME)); } catch { }



                try { addtolist("11" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNAME)); } catch { }
                try { addtolist("12" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNUMBER)); } catch { }
                try { addtolist("13" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_DISPLAYABLEADDRESS)); } catch { }
                try { addtolist("14" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNAME)); } catch { }
                try { addtolist("15" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNUMBER)); } catch { }
                try { addtolist("16" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNAME)); } catch { }
                try { addtolist("17" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNUMBER)); } catch { }
                try { addtolist("18" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME)); } catch { }
                try { addtolist("19" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNUMBER)); } catch { }
                try { addtolist("20" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER)); } catch { }
                try { addtolist("21" + b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME)); } catch { }


                /* try { addtolist("21" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNAME)); } catch { }
                 try { addtolist("22" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNUMBER)); } catch { }
                 try { addtolist("23" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_DISPLAYABLEADDRESS)); } catch { }
                 try { addtolist("24" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNAME)); } catch { }
                 try { addtolist("25" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNUMBER)); } catch { }
                 try { addtolist("26" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNAME)); } catch { }
                 try { addtolist("27" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNUMBER)); } catch { }*/

                /*clHistory.FromName = b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME);
                clHistory.FromNumber = a.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER);
                clHistory.ToName = a.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME);
                clHistory.ToNumber = a.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNUMBER);*/
                try { clHistory.FromName = b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME); } catch { }
                try { clHistory.FromNumber = b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER); } catch { }
                try { clHistory.ToNumber = b.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNUMBER); } catch { }
                try { clHistory.CallID = b.get_CallInfoLong(CALLINFO_LONG.CIL_CALLID); } catch { }
                clHistory.Direction  = "Outgoing";
                addToHistory(clHistory);
              }
              catch
              {
                clHistory.FromName = "Error";
                addToHistory(clHistory);
              }
              

              /*if (cn.Call.CallState == TAPI3Lib.CALL_STATE.CS_OFFERING)
              {
                string c = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER);
                addtolist("Call Offering: " + c + " -> " + cn.Call.Address.DialableAddress);

                //classHistory clHistory = new classHistory();
                clHistory.FromName = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME);
                clHistory.FromNumber = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER);
                clHistory.ToName = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME);
                clHistory.ToNumber = cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNUMBER);
                clHistory.Direction = "Incoming";
                addToHistory(clHistory);

                addToCallNotify(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME) + " (" + cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER) + ")");
                addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDIDNAME));//Anita
                                                                                                 //addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLEDPARTYFRIENDLYNAME));
                addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNAME));//EDV-Support
                addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLERIDNUMBER));//410
                                                                                                   //addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CALLINGPARTYID));
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNAME)); } catch { }
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_CONNECTEDIDNUMBER)); } catch { }
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_DISPLAYABLEADDRESS)); } catch { }
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNAME)); } catch { }
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTINGIDNUMBER)); } catch { }
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNAME)); } catch { }
                try { addtolist(cn.Call.get_CallInfoString(TAPI3Lib.CALLINFO_STRING.CIS_REDIRECTIONIDNUMBER)); } catch { }
              }*/

              break;
					case TAPI3Lib.CALL_STATE.CS_CONNECTED:
						addtolist("Connected");
						break;
					case TAPI3Lib.CALL_STATE.CS_DISCONNECTED:
						addtolist("Disconnected");
						break;
					case TAPI3Lib.CALL_STATE.CS_OFFERING:
						addtolist("A party wants to communicate with you!");
						break;
					case TAPI3Lib.CALL_STATE.CS_IDLE:
						addtolist("Call is created!");
						break;
				}
				break;
			}
		}
	}
}
