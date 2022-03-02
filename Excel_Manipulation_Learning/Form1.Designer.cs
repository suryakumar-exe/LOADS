
namespace Excel_Manipulation_Learning
{
    partial class Form1
    {
        /// <summary>
        ///  Required designer variable.
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        ///  Clean up any resources being used.
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
        ///  Required method for Designer support - do not modify
        ///  the contents of this method with the code editor.
        /// </summary>
        private void InitializeComponent()
        {
            this.button2 = new System.Windows.Forms.Button();
            this.Main_compare = new System.Windows.Forms.Button();
            this.Fatigue = new System.Windows.Forms.ComboBox();
            this.QuickView = new System.Windows.Forms.ComboBox();
            this.Fatigue_lbl = new System.Windows.Forms.Label();
            this.Quickview_lbl = new System.Windows.Forms.Label();
            this.btn_fatigue = new System.Windows.Forms.Button();
            this.btn_quickview = new System.Windows.Forms.Button();
            this.quickview_cb = new System.Windows.Forms.ComboBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label2 = new System.Windows.Forms.Label();
            this.hub_cb = new System.Windows.Forms.ComboBox();
            this.label3 = new System.Windows.Forms.Label();
            this.mainshaft_cb = new System.Windows.Forms.ComboBox();
            this.label7 = new System.Windows.Forms.Label();
            this.blade_cb = new System.Windows.Forms.ComboBox();
            this.label8 = new System.Windows.Forms.Label();
            this.yaw_cb = new System.Windows.Forms.ComboBox();
            this.label9 = new System.Windows.Forms.Label();
            this.mainframe_cb = new System.Windows.Forms.ComboBox();
            this.label10 = new System.Windows.Forms.Label();
            this.tower_cb = new System.Windows.Forms.ComboBox();
            this.label11 = new System.Windows.Forms.Label();
            this.genframe_cb = new System.Windows.Forms.ComboBox();
            this.label12 = new System.Windows.Forms.Label();
            this.steeltower_cb = new System.Windows.Forms.ComboBox();
            this.label13 = new System.Windows.Forms.Label();
            this.towerca_cb = new System.Windows.Forms.ComboBox();
            this.gearbox_cb = new System.Windows.Forms.ComboBox();
            this.label4 = new System.Windows.Forms.Label();
            this.gbzf_cb = new System.Windows.Forms.ComboBox();
            this.label5 = new System.Windows.Forms.Label();
            this.gbelickoff_cb = new System.Windows.Forms.ComboBox();
            this.label6 = new System.Windows.Forms.Label();
            this.progressBar1 = new System.Windows.Forms.ProgressBar();
            this.backgroundWorker1 = new System.ComponentModel.BackgroundWorker();
            this.label14 = new System.Windows.Forms.Label();
            this.btn_refernceload = new System.Windows.Forms.Button();
            this.label15 = new System.Windows.Forms.Label();
            this.refload_main = new System.Windows.Forms.ComboBox();
            this.refload_cb = new System.Windows.Forms.ComboBox();
            this.label16 = new System.Windows.Forms.Label();
            this.label17 = new System.Windows.Forms.Label();
            this.gpsupport = new System.Windows.Forms.ComboBox();
            this.label18 = new System.Windows.Forms.Label();
            this.roterbearing = new System.Windows.Forms.ComboBox();
            this.btn_additionalsensor = new System.Windows.Forms.Button();
            this.label19 = new System.Windows.Forms.Label();
            this.cbx_asn = new System.Windows.Forms.ComboBox();
            this.SuspendLayout();
            // 
            // button2
            // 
            this.button2.Location = new System.Drawing.Point(713, 12);
            this.button2.Name = "button2";
            this.button2.Size = new System.Drawing.Size(75, 23);
            this.button2.TabIndex = 1;
            this.button2.Text = "button2";
            this.button2.UseVisualStyleBackColor = true;
            this.button2.Click += new System.EventHandler(this.button2_Click);
            // 
            // Main_compare
            // 
            this.Main_compare.Location = new System.Drawing.Point(419, 376);
            this.Main_compare.Name = "Main_compare";
            this.Main_compare.Size = new System.Drawing.Size(75, 23);
            this.Main_compare.TabIndex = 2;
            this.Main_compare.Text = "Compare";
            this.Main_compare.UseVisualStyleBackColor = true;
            this.Main_compare.Click += new System.EventHandler(this.button3_Click);
            // 
            // Fatigue
            // 
            this.Fatigue.FormattingEnabled = true;
            this.Fatigue.Location = new System.Drawing.Point(250, 68);
            this.Fatigue.Name = "Fatigue";
            this.Fatigue.Size = new System.Drawing.Size(274, 23);
            this.Fatigue.TabIndex = 3;
            this.Fatigue.SelectedIndexChanged += new System.EventHandler(this.Fatigue_SelectedIndexChanged);
            // 
            // QuickView
            // 
            this.QuickView.FormattingEnabled = true;
            this.QuickView.Location = new System.Drawing.Point(250, 97);
            this.QuickView.Name = "QuickView";
            this.QuickView.Size = new System.Drawing.Size(274, 23);
            this.QuickView.TabIndex = 4;
            // 
            // Fatigue_lbl
            // 
            this.Fatigue_lbl.AutoSize = true;
            this.Fatigue_lbl.Location = new System.Drawing.Point(183, 76);
            this.Fatigue_lbl.Name = "Fatigue_lbl";
            this.Fatigue_lbl.Size = new System.Drawing.Size(46, 15);
            this.Fatigue_lbl.TabIndex = 5;
            this.Fatigue_lbl.Text = "Fatigue";
            // 
            // Quickview_lbl
            // 
            this.Quickview_lbl.AutoSize = true;
            this.Quickview_lbl.Location = new System.Drawing.Point(183, 105);
            this.Quickview_lbl.Name = "Quickview_lbl";
            this.Quickview_lbl.Size = new System.Drawing.Size(63, 15);
            this.Quickview_lbl.TabIndex = 6;
            this.Quickview_lbl.Text = "QuickView";
            // 
            // btn_fatigue
            // 
            this.btn_fatigue.Location = new System.Drawing.Point(543, 68);
            this.btn_fatigue.Name = "btn_fatigue";
            this.btn_fatigue.Size = new System.Drawing.Size(75, 23);
            this.btn_fatigue.TabIndex = 7;
            this.btn_fatigue.Text = "Browse";
            this.btn_fatigue.UseVisualStyleBackColor = true;
            this.btn_fatigue.Click += new System.EventHandler(this.button3_Click_1);
            // 
            // btn_quickview
            // 
            this.btn_quickview.Location = new System.Drawing.Point(543, 97);
            this.btn_quickview.Name = "btn_quickview";
            this.btn_quickview.Size = new System.Drawing.Size(75, 23);
            this.btn_quickview.TabIndex = 8;
            this.btn_quickview.Text = "Browse";
            this.btn_quickview.UseVisualStyleBackColor = true;
            this.btn_quickview.Click += new System.EventHandler(this.button4_Click);
            // 
            // quickview_cb
            // 
            this.quickview_cb.FormattingEnabled = true;
            this.quickview_cb.Location = new System.Drawing.Point(213, 212);
            this.quickview_cb.Name = "quickview_cb";
            this.quickview_cb.Size = new System.Drawing.Size(111, 23);
            this.quickview_cb.TabIndex = 9;
            this.quickview_cb.SelectedIndexChanged += new System.EventHandler(this.quickview_cb_SelectedIndexChanged);
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(143, 220);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(34, 15);
            this.label1.TabIndex = 10;
            this.label1.Text = "Pitch";
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(330, 220);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(30, 15);
            this.label2.TabIndex = 11;
            this.label2.Text = "Hub";
            // 
            // hub_cb
            // 
            this.hub_cb.FormattingEnabled = true;
            this.hub_cb.Location = new System.Drawing.Point(419, 212);
            this.hub_cb.Name = "hub_cb";
            this.hub_cb.Size = new System.Drawing.Size(105, 23);
            this.hub_cb.TabIndex = 12;
            this.hub_cb.SelectedIndexChanged += new System.EventHandler(this.hub_cb_SelectedIndexChanged);
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(143, 244);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(64, 15);
            this.label3.TabIndex = 14;
            this.label3.Text = "Main Shaft";
            // 
            // mainshaft_cb
            // 
            this.mainshaft_cb.FormattingEnabled = true;
            this.mainshaft_cb.Location = new System.Drawing.Point(213, 241);
            this.mainshaft_cb.Name = "mainshaft_cb";
            this.mainshaft_cb.Size = new System.Drawing.Size(111, 23);
            this.mainshaft_cb.TabIndex = 13;
            this.mainshaft_cb.SelectedIndexChanged += new System.EventHandler(this.mainshaft_cb_SelectedIndexChanged);
            // 
            // label7
            // 
            this.label7.AutoSize = true;
            this.label7.Location = new System.Drawing.Point(143, 273);
            this.label7.Name = "label7";
            this.label7.Size = new System.Drawing.Size(36, 15);
            this.label7.TabIndex = 22;
            this.label7.Text = "Blade";
            this.label7.Click += new System.EventHandler(this.label7_Click);
            // 
            // blade_cb
            // 
            this.blade_cb.FormattingEnabled = true;
            this.blade_cb.Location = new System.Drawing.Point(213, 270);
            this.blade_cb.Name = "blade_cb";
            this.blade_cb.Size = new System.Drawing.Size(111, 23);
            this.blade_cb.TabIndex = 21;
            this.blade_cb.SelectedIndexChanged += new System.EventHandler(this.blade_cb_SelectedIndexChanged);
            // 
            // label8
            // 
            this.label8.AutoSize = true;
            this.label8.Location = new System.Drawing.Point(330, 249);
            this.label8.Name = "label8";
            this.label8.Size = new System.Drawing.Size(28, 15);
            this.label8.TabIndex = 24;
            this.label8.Text = "Yaw";
            // 
            // yaw_cb
            // 
            this.yaw_cb.FormattingEnabled = true;
            this.yaw_cb.Location = new System.Drawing.Point(419, 241);
            this.yaw_cb.Name = "yaw_cb";
            this.yaw_cb.Size = new System.Drawing.Size(105, 23);
            this.yaw_cb.TabIndex = 23;
            this.yaw_cb.SelectedIndexChanged += new System.EventHandler(this.yaw_cb_SelectedIndexChanged);
            // 
            // label9
            // 
            this.label9.AutoSize = true;
            this.label9.Location = new System.Drawing.Point(143, 307);
            this.label9.Name = "label9";
            this.label9.Size = new System.Drawing.Size(70, 15);
            this.label9.TabIndex = 26;
            this.label9.Text = "Main Frame";
            // 
            // mainframe_cb
            // 
            this.mainframe_cb.FormattingEnabled = true;
            this.mainframe_cb.Location = new System.Drawing.Point(213, 299);
            this.mainframe_cb.Name = "mainframe_cb";
            this.mainframe_cb.Size = new System.Drawing.Size(111, 23);
            this.mainframe_cb.TabIndex = 25;
            this.mainframe_cb.SelectedIndexChanged += new System.EventHandler(this.comboBox1_SelectedIndexChanged);
            // 
            // label10
            // 
            this.label10.AutoSize = true;
            this.label10.Location = new System.Drawing.Point(330, 307);
            this.label10.Name = "label10";
            this.label10.Size = new System.Drawing.Size(84, 15);
            this.label10.TabIndex = 28;
            this.label10.Text = "Conrete Tower";
            // 
            // tower_cb
            // 
            this.tower_cb.FormattingEnabled = true;
            this.tower_cb.Location = new System.Drawing.Point(419, 299);
            this.tower_cb.Name = "tower_cb";
            this.tower_cb.Size = new System.Drawing.Size(105, 23);
            this.tower_cb.TabIndex = 27;
            this.tower_cb.SelectedIndexChanged += new System.EventHandler(this.tower_cb_SelectedIndexChanged);
            // 
            // label11
            // 
            this.label11.AutoSize = true;
            this.label11.Location = new System.Drawing.Point(143, 336);
            this.label11.Name = "label11";
            this.label11.Size = new System.Drawing.Size(64, 15);
            this.label11.TabIndex = 30;
            this.label11.Text = "Gen Frame";
            // 
            // genframe_cb
            // 
            this.genframe_cb.FormattingEnabled = true;
            this.genframe_cb.Location = new System.Drawing.Point(213, 328);
            this.genframe_cb.Name = "genframe_cb";
            this.genframe_cb.Size = new System.Drawing.Size(111, 23);
            this.genframe_cb.TabIndex = 29;
            this.genframe_cb.SelectedIndexChanged += new System.EventHandler(this.genframe_cb_SelectedIndexChanged);
            // 
            // label12
            // 
            this.label12.AutoSize = true;
            this.label12.Location = new System.Drawing.Point(330, 336);
            this.label12.Name = "label12";
            this.label12.Size = new System.Drawing.Size(67, 15);
            this.label12.TabIndex = 32;
            this.label12.Text = "Steel Tower";
            // 
            // steeltower_cb
            // 
            this.steeltower_cb.FormattingEnabled = true;
            this.steeltower_cb.Location = new System.Drawing.Point(419, 328);
            this.steeltower_cb.Name = "steeltower_cb";
            this.steeltower_cb.Size = new System.Drawing.Size(105, 23);
            this.steeltower_cb.TabIndex = 31;
            this.steeltower_cb.SelectedIndexChanged += new System.EventHandler(this.steeltower_cb_SelectedIndexChanged);
            // 
            // label13
            // 
            this.label13.AutoSize = true;
            this.label13.Location = new System.Drawing.Point(330, 274);
            this.label13.Name = "label13";
            this.label13.Size = new System.Drawing.Size(58, 15);
            this.label13.TabIndex = 34;
            this.label13.Text = "Tower CA";
            // 
            // towerca_cb
            // 
            this.towerca_cb.FormattingEnabled = true;
            this.towerca_cb.Location = new System.Drawing.Point(419, 270);
            this.towerca_cb.Name = "towerca_cb";
            this.towerca_cb.Size = new System.Drawing.Size(105, 23);
            this.towerca_cb.TabIndex = 33;
            this.towerca_cb.SelectedIndexChanged += new System.EventHandler(this.towerca_cb_SelectedIndexChanged);
            // 
            // gearbox_cb
            // 
            this.gearbox_cb.FormattingEnabled = true;
            this.gearbox_cb.Location = new System.Drawing.Point(612, 212);
            this.gearbox_cb.Name = "gearbox_cb";
            this.gearbox_cb.Size = new System.Drawing.Size(205, 23);
            this.gearbox_cb.TabIndex = 15;
            this.gearbox_cb.SelectedIndexChanged += new System.EventHandler(this.gearbox_cb_SelectedIndexChanged);
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(536, 215);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(63, 15);
            this.label4.TabIndex = 16;
            this.label4.Text = "GB Winerg";
            // 
            // gbzf_cb
            // 
            this.gbzf_cb.FormattingEnabled = true;
            this.gbzf_cb.Location = new System.Drawing.Point(612, 266);
            this.gbzf_cb.Name = "gbzf_cb";
            this.gbzf_cb.Size = new System.Drawing.Size(205, 23);
            this.gbzf_cb.TabIndex = 17;
            this.gbzf_cb.SelectedIndexChanged += new System.EventHandler(this.gbzf_cb_SelectedIndexChanged);
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(536, 273);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(38, 15);
            this.label5.TabIndex = 18;
            this.label5.Text = "GB ZF";
            // 
            // gbelickoff_cb
            // 
            this.gbelickoff_cb.FormattingEnabled = true;
            this.gbelickoff_cb.Location = new System.Drawing.Point(612, 241);
            this.gbelickoff_cb.Name = "gbelickoff_cb";
            this.gbelickoff_cb.Size = new System.Drawing.Size(205, 23);
            this.gbelickoff_cb.TabIndex = 19;
            this.gbelickoff_cb.SelectedIndexChanged += new System.EventHandler(this.gbelickoff_cb_SelectedIndexChanged);
            // 
            // label6
            // 
            this.label6.AutoSize = true;
            this.label6.Location = new System.Drawing.Point(536, 244);
            this.label6.Name = "label6";
            this.label6.Size = new System.Drawing.Size(64, 15);
            this.label6.TabIndex = 20;
            this.label6.Text = "GB Elickoff";
            // 
            // progressBar1
            // 
            this.progressBar1.Location = new System.Drawing.Point(12, 12);
            this.progressBar1.Name = "progressBar1";
            this.progressBar1.Size = new System.Drawing.Size(695, 23);
            this.progressBar1.TabIndex = 35;
            // 
            // backgroundWorker1
            // 
            this.backgroundWorker1.WorkerReportsProgress = true;
            this.backgroundWorker1.DoWork += new System.ComponentModel.DoWorkEventHandler(this.backgroundWorker1_DoWork);
            this.backgroundWorker1.ProgressChanged += new System.ComponentModel.ProgressChangedEventHandler(this.backgroundWorker1_ProgressChanged);
            this.backgroundWorker1.RunWorkerCompleted += new System.ComponentModel.RunWorkerCompletedEventHandler(this.backgroundWorker1_RunWorkerCompleted);
            // 
            // label14
            // 
            this.label14.AutoSize = true;
            this.label14.Location = new System.Drawing.Point(13, 42);
            this.label14.Name = "label14";
            this.label14.Size = new System.Drawing.Size(23, 15);
            this.label14.TabIndex = 36;
            this.label14.Text = "0%";
            // 
            // btn_refernceload
            // 
            this.btn_refernceload.Location = new System.Drawing.Point(667, 126);
            this.btn_refernceload.Name = "btn_refernceload";
            this.btn_refernceload.Size = new System.Drawing.Size(75, 23);
            this.btn_refernceload.TabIndex = 39;
            this.btn_refernceload.Text = "Browse";
            this.btn_refernceload.UseVisualStyleBackColor = true;
            this.btn_refernceload.Click += new System.EventHandler(this.button1_Click);
            // 
            // label15
            // 
            this.label15.AutoSize = true;
            this.label15.BackColor = System.Drawing.SystemColors.Menu;
            this.label15.Location = new System.Drawing.Point(183, 134);
            this.label15.Name = "label15";
            this.label15.Size = new System.Drawing.Size(53, 15);
            this.label15.TabIndex = 38;
            this.label15.Text = "Ref Load";
            // 
            // refload_main
            // 
            this.refload_main.FormattingEnabled = true;
            this.refload_main.Location = new System.Drawing.Point(250, 126);
            this.refload_main.Name = "refload_main";
            this.refload_main.Size = new System.Drawing.Size(274, 23);
            this.refload_main.TabIndex = 37;
            this.refload_main.SelectedIndexChanged += new System.EventHandler(this.refload_main_SelectedIndexChanged);
            // 
            // refload_cb
            // 
            this.refload_cb.FormattingEnabled = true;
            this.refload_cb.Location = new System.Drawing.Point(530, 126);
            this.refload_cb.Name = "refload_cb";
            this.refload_cb.Size = new System.Drawing.Size(121, 23);
            this.refload_cb.TabIndex = 40;
            this.refload_cb.SelectedIndexChanged += new System.EventHandler(this.refload_cb_SelectedIndexChanged);
            // 
            // label16
            // 
            this.label16.AutoSize = true;
            this.label16.Location = new System.Drawing.Point(603, 384);
            this.label16.Name = "label16";
            this.label16.Size = new System.Drawing.Size(39, 15);
            this.label16.TabIndex = 41;
            this.label16.Text = "Status";
            // 
            // label17
            // 
            this.label17.AutoSize = true;
            this.label17.Location = new System.Drawing.Point(536, 307);
            this.label17.Name = "label17";
            this.label17.Size = new System.Drawing.Size(67, 15);
            this.label17.TabIndex = 43;
            this.label17.Text = "GB Support";
            // 
            // gpsupport
            // 
            this.gpsupport.FormattingEnabled = true;
            this.gpsupport.Location = new System.Drawing.Point(612, 295);
            this.gpsupport.Name = "gpsupport";
            this.gpsupport.Size = new System.Drawing.Size(205, 23);
            this.gpsupport.TabIndex = 42;
            // 
            // label18
            // 
            this.label18.AutoSize = true;
            this.label18.Location = new System.Drawing.Point(536, 332);
            this.label18.Name = "label18";
            this.label18.Size = new System.Drawing.Size(78, 15);
            this.label18.TabIndex = 45;
            this.label18.Text = "Roter bearing";
            // 
            // roterbearing
            // 
            this.roterbearing.FormattingEnabled = true;
            this.roterbearing.Location = new System.Drawing.Point(614, 324);
            this.roterbearing.Name = "roterbearing";
            this.roterbearing.Size = new System.Drawing.Size(205, 23);
            this.roterbearing.TabIndex = 44;
            this.roterbearing.SelectedIndexChanged += new System.EventHandler(this.roterbearing_SelectedIndexChanged);
            // 
            // btn_additionalsensor
            // 
            this.btn_additionalsensor.Location = new System.Drawing.Point(543, 155);
            this.btn_additionalsensor.Name = "btn_additionalsensor";
            this.btn_additionalsensor.Size = new System.Drawing.Size(75, 23);
            this.btn_additionalsensor.TabIndex = 48;
            this.btn_additionalsensor.Text = "Browse";
            this.btn_additionalsensor.UseVisualStyleBackColor = true;
            this.btn_additionalsensor.Click += new System.EventHandler(this.btn_asn_Click);
            // 
            // label19
            // 
            this.label19.AutoSize = true;
            this.label19.Location = new System.Drawing.Point(146, 158);
            this.label19.Name = "label19";
            this.label19.Size = new System.Drawing.Size(100, 15);
            this.label19.TabIndex = 47;
            this.label19.Text = "Additional Sensor";
            // 
            // cbx_asn
            // 
            this.cbx_asn.FormattingEnabled = true;
            this.cbx_asn.Location = new System.Drawing.Point(250, 155);
            this.cbx_asn.Name = "cbx_asn";
            this.cbx_asn.Size = new System.Drawing.Size(274, 23);
            this.cbx_asn.TabIndex = 46;
            // 
            // Form1
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(7F, 15F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1045, 562);
            this.Controls.Add(this.btn_additionalsensor);
            this.Controls.Add(this.label19);
            this.Controls.Add(this.cbx_asn);
            this.Controls.Add(this.label18);
            this.Controls.Add(this.roterbearing);
            this.Controls.Add(this.label17);
            this.Controls.Add(this.gpsupport);
            this.Controls.Add(this.label16);
            this.Controls.Add(this.refload_cb);
            this.Controls.Add(this.btn_refernceload);
            this.Controls.Add(this.label15);
            this.Controls.Add(this.refload_main);
            this.Controls.Add(this.label14);
            this.Controls.Add(this.progressBar1);
            this.Controls.Add(this.label13);
            this.Controls.Add(this.towerca_cb);
            this.Controls.Add(this.label12);
            this.Controls.Add(this.steeltower_cb);
            this.Controls.Add(this.label11);
            this.Controls.Add(this.genframe_cb);
            this.Controls.Add(this.label10);
            this.Controls.Add(this.tower_cb);
            this.Controls.Add(this.label9);
            this.Controls.Add(this.mainframe_cb);
            this.Controls.Add(this.label8);
            this.Controls.Add(this.yaw_cb);
            this.Controls.Add(this.label7);
            this.Controls.Add(this.blade_cb);
            this.Controls.Add(this.label6);
            this.Controls.Add(this.gbelickoff_cb);
            this.Controls.Add(this.label5);
            this.Controls.Add(this.gbzf_cb);
            this.Controls.Add(this.label4);
            this.Controls.Add(this.gearbox_cb);
            this.Controls.Add(this.label3);
            this.Controls.Add(this.mainshaft_cb);
            this.Controls.Add(this.hub_cb);
            this.Controls.Add(this.label2);
            this.Controls.Add(this.label1);
            this.Controls.Add(this.quickview_cb);
            this.Controls.Add(this.btn_quickview);
            this.Controls.Add(this.btn_fatigue);
            this.Controls.Add(this.Quickview_lbl);
            this.Controls.Add(this.Fatigue_lbl);
            this.Controls.Add(this.QuickView);
            this.Controls.Add(this.Fatigue);
            this.Controls.Add(this.Main_compare);
            this.Controls.Add(this.button2);
            this.Name = "Form1";
            this.Text = "Loads Comparison";
            this.Load += new System.EventHandler(this.Form1_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion
        private System.Windows.Forms.Button button2;
        private System.Windows.Forms.Button Main_compare;
        private System.Windows.Forms.ComboBox Fatigue;
        private System.Windows.Forms.ComboBox QuickView;
        private System.Windows.Forms.Label Fatigue_lbl;
        public System.Windows.Forms.Label Quickview_lbl;
        private System.Windows.Forms.Button btn_fatigue;
        private System.Windows.Forms.Button btn_quickview;
        private System.Windows.Forms.ComboBox quickview_cb;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.ComboBox hub_cb;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.ComboBox mainshaft_cb;
        private System.Windows.Forms.Label label7;
        private System.Windows.Forms.ComboBox blade_cb;
        private System.Windows.Forms.Label label8;
        private System.Windows.Forms.ComboBox yaw_cb;
        private System.Windows.Forms.Label label9;
        private System.Windows.Forms.ComboBox mainframe_cb;
        private System.Windows.Forms.Label label10;
        private System.Windows.Forms.ComboBox tower_cb;
        private System.Windows.Forms.Label label11;
        private System.Windows.Forms.ComboBox genframe_cb;
        private System.Windows.Forms.Label label12;
        private System.Windows.Forms.ComboBox steeltower_cb;
        private System.Windows.Forms.Label label13;
        private System.Windows.Forms.ComboBox towerca_cb;
        private System.Windows.Forms.ComboBox gearbox_cb;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.ComboBox gbzf_cb;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.ComboBox gbelickoff_cb;
        private System.Windows.Forms.Label label6;
        private System.Windows.Forms.ProgressBar progressBar1;
        private System.ComponentModel.BackgroundWorker backgroundWorker1;
        private System.Windows.Forms.Label label14;
        private System.Windows.Forms.Button btn_refernceload;
        public System.Windows.Forms.Label label15;
        private System.Windows.Forms.ComboBox refload_main;
        private System.Windows.Forms.ComboBox refload_cb;
        private System.Windows.Forms.Label label16;
        private System.Windows.Forms.Label label17;
        private System.Windows.Forms.ComboBox gpsupport;
        private System.Windows.Forms.Label label18;
        private System.Windows.Forms.ComboBox roterbearing;
        private System.Windows.Forms.Button btn_additionalsensor;
        private System.Windows.Forms.Label label19;
        private System.Windows.Forms.ComboBox cbx_asn;
    }
}

