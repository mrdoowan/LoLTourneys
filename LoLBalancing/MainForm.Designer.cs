namespace LoLBalancing
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
		protected override void Dispose(bool disposing) {
			if (disposing && (components != null)) {
				components.Dispose();
			}
			base.Dispose(disposing);
		}

		#region Windows Form Designer generated code

		/// <summary>
		/// Required method for Designer support - do not modify
		/// the contents of this method with the code editor.
		/// </summary>
		private void InitializeComponent() {
			this.components = new System.ComponentModel.Container();
			System.ComponentModel.ComponentResourceManager resources = new System.ComponentModel.ComponentResourceManager(typeof(MainForm));
			this.pictureBox1 = new System.Windows.Forms.PictureBox();
			this.pictureBox2 = new System.Windows.Forms.PictureBox();
			this.label1 = new System.Windows.Forms.Label();
			this.tabControl1 = new System.Windows.Forms.TabControl();
			this.tabPage1 = new System.Windows.Forms.TabPage();
			this.label_Total = new System.Windows.Forms.Label();
			this.button_SavePlayers = new System.Windows.Forms.Button();
			this.button_LoadPlayers = new System.Windows.Forms.Button();
			this.dataGridView_Players = new System.Windows.Forms.DataGridView();
			this.DelPlayer = new System.Windows.Forms.DataGridViewButtonColumn();
			this.NamePlayer = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.UniqPlayer = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.SummName = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.RankPlayer = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.RolesPlayer = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.DuoPlayer = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.button_AddPlayer = new System.Windows.Forms.Button();
			this.tabPage2 = new System.Windows.Forms.TabPage();
			this.pictureBox3 = new System.Windows.Forms.PictureBox();
			this.button_Balance = new System.Windows.Forms.Button();
			this.button_ResetPoints = new System.Windows.Forms.Button();
			this.groupBox_Setting = new System.Windows.Forms.GroupBox();
			this.checkBox_EveryRole = new System.Windows.Forms.CheckBox();
			this.label2 = new System.Windows.Forms.Label();
			this.numericUpDown_RandNum = new System.Windows.Forms.NumericUpDown();
			this.checkBox_BalByRole = new System.Windows.Forms.CheckBox();
			this.dataGridView_Ranks = new System.Windows.Forms.DataGridView();
			this.Rank = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.Value = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.tabPage3 = new System.Windows.Forms.TabPage();
			this.groupBox2 = new System.Windows.Forms.GroupBox();
			this.label3 = new System.Windows.Forms.Label();
			this.label_TeamPoints = new System.Windows.Forms.Label();
			this.comboBox_Teams = new System.Windows.Forms.ComboBox();
			this.dataGridView_Team = new System.Windows.Forms.DataGridView();
			this.dataGridViewTextBoxColumn1 = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.dataGridViewTextBoxColumn2 = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.RolesPlay = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.RankPlay = new System.Windows.Forms.DataGridViewTextBoxColumn();
			this.groupBox1 = new System.Windows.Forms.GroupBox();
			this.button_LoadTeams = new System.Windows.Forms.Button();
			this.label_Version = new System.Windows.Forms.Label();
			this.groupBox4 = new System.Windows.Forms.GroupBox();
			this.label5 = new System.Windows.Forms.Label();
			this.textBox_APIKey = new System.Windows.Forms.TextBox();
			this.button_GenStats = new System.Windows.Forms.Button();
			this.button_HelpStats = new System.Windows.Forms.Button();
			this.checkBox_ProdKey = new System.Windows.Forms.CheckBox();
			this.toolTip1 = new System.Windows.Forms.ToolTip(this.components);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).BeginInit();
			this.tabControl1.SuspendLayout();
			this.tabPage1.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_Players)).BeginInit();
			this.tabPage2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).BeginInit();
			this.groupBox_Setting.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.numericUpDown_RandNum)).BeginInit();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_Ranks)).BeginInit();
			this.tabPage3.SuspendLayout();
			this.groupBox2.SuspendLayout();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_Team)).BeginInit();
			this.groupBox1.SuspendLayout();
			this.groupBox4.SuspendLayout();
			this.SuspendLayout();
			// 
			// pictureBox1
			// 
			this.pictureBox1.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox1.Image")));
			this.pictureBox1.Location = new System.Drawing.Point(12, 27);
			this.pictureBox1.Name = "pictureBox1";
			this.pictureBox1.Size = new System.Drawing.Size(100, 100);
			this.pictureBox1.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox1.TabIndex = 0;
			this.pictureBox1.TabStop = false;
			// 
			// pictureBox2
			// 
			this.pictureBox2.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox2.Image")));
			this.pictureBox2.Location = new System.Drawing.Point(549, 27);
			this.pictureBox2.Name = "pictureBox2";
			this.pictureBox2.Size = new System.Drawing.Size(100, 100);
			this.pictureBox2.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox2.TabIndex = 1;
			this.pictureBox2.TabStop = false;
			// 
			// label1
			// 
			this.label1.Font = new System.Drawing.Font("Microsoft Sans Serif", 20.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
			this.label1.Location = new System.Drawing.Point(118, 27);
			this.label1.Name = "label1";
			this.label1.Size = new System.Drawing.Size(425, 100);
			this.label1.TabIndex = 2;
			this.label1.Text = "League of Legends\r\nTeam Balancer\r\nStats Collector";
			this.label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter;
			// 
			// tabControl1
			// 
			this.tabControl1.Controls.Add(this.tabPage1);
			this.tabControl1.Controls.Add(this.tabPage2);
			this.tabControl1.Controls.Add(this.tabPage3);
			this.tabControl1.Location = new System.Drawing.Point(12, 150);
			this.tabControl1.Name = "tabControl1";
			this.tabControl1.SelectedIndex = 0;
			this.tabControl1.Size = new System.Drawing.Size(637, 399);
			this.tabControl1.TabIndex = 3;
			// 
			// tabPage1
			// 
			this.tabPage1.Controls.Add(this.label_Total);
			this.tabPage1.Controls.Add(this.button_SavePlayers);
			this.tabPage1.Controls.Add(this.button_LoadPlayers);
			this.tabPage1.Controls.Add(this.dataGridView_Players);
			this.tabPage1.Controls.Add(this.button_AddPlayer);
			this.tabPage1.Location = new System.Drawing.Point(4, 22);
			this.tabPage1.Name = "tabPage1";
			this.tabPage1.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage1.Size = new System.Drawing.Size(629, 373);
			this.tabPage1.TabIndex = 0;
			this.tabPage1.Text = "Player Roster";
			this.tabPage1.UseVisualStyleBackColor = true;
			// 
			// label_Total
			// 
			this.label_Total.ForeColor = System.Drawing.Color.Teal;
			this.label_Total.Location = new System.Drawing.Point(264, 6);
			this.label_Total.Name = "label_Total";
			this.label_Total.Size = new System.Drawing.Size(359, 23);
			this.label_Total.TabIndex = 5;
			this.label_Total.Text = "Total Players: 0";
			this.label_Total.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// button_SavePlayers
			// 
			this.button_SavePlayers.Location = new System.Drawing.Point(178, 6);
			this.button_SavePlayers.Name = "button_SavePlayers";
			this.button_SavePlayers.Size = new System.Drawing.Size(80, 23);
			this.button_SavePlayers.TabIndex = 3;
			this.button_SavePlayers.Text = "Save Players";
			this.button_SavePlayers.UseVisualStyleBackColor = true;
			this.button_SavePlayers.Click += new System.EventHandler(this.button_SavePlayers_Click);
			// 
			// button_LoadPlayers
			// 
			this.button_LoadPlayers.Location = new System.Drawing.Point(92, 6);
			this.button_LoadPlayers.Name = "button_LoadPlayers";
			this.button_LoadPlayers.Size = new System.Drawing.Size(80, 23);
			this.button_LoadPlayers.TabIndex = 2;
			this.button_LoadPlayers.Text = "Load Players";
			this.button_LoadPlayers.UseVisualStyleBackColor = true;
			this.button_LoadPlayers.Click += new System.EventHandler(this.button_LoadPlayers_Click);
			// 
			// dataGridView_Players
			// 
			this.dataGridView_Players.AllowUserToAddRows = false;
			this.dataGridView_Players.AllowUserToDeleteRows = false;
			this.dataGridView_Players.AllowUserToResizeRows = false;
			this.dataGridView_Players.BackgroundColor = System.Drawing.SystemColors.Window;
			this.dataGridView_Players.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_Players.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.DelPlayer,
            this.NamePlayer,
            this.UniqPlayer,
            this.SummName,
            this.RankPlayer,
            this.RolesPlayer,
            this.DuoPlayer});
			this.dataGridView_Players.Location = new System.Drawing.Point(6, 35);
			this.dataGridView_Players.MultiSelect = false;
			this.dataGridView_Players.Name = "dataGridView_Players";
			this.dataGridView_Players.RowHeadersVisible = false;
			this.dataGridView_Players.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.dataGridView_Players.Size = new System.Drawing.Size(617, 335);
			this.dataGridView_Players.TabIndex = 1;
			this.dataGridView_Players.CellContentClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_Players_CellContentClick);
			this.dataGridView_Players.CellDoubleClick += new System.Windows.Forms.DataGridViewCellEventHandler(this.dataGridView_Players_CellDoubleClick);
			// 
			// DelPlayer
			// 
			this.DelPlayer.HeaderText = "X";
			this.DelPlayer.Name = "DelPlayer";
			this.DelPlayer.ReadOnly = true;
			this.DelPlayer.Width = 25;
			// 
			// NamePlayer
			// 
			this.NamePlayer.HeaderText = "Name";
			this.NamePlayer.Name = "NamePlayer";
			this.NamePlayer.ReadOnly = true;
			this.NamePlayer.Width = 125;
			// 
			// UniqPlayer
			// 
			this.UniqPlayer.HeaderText = "Uniq";
			this.UniqPlayer.Name = "UniqPlayer";
			this.UniqPlayer.ReadOnly = true;
			this.UniqPlayer.Width = 75;
			// 
			// SummName
			// 
			this.SummName.HeaderText = "Summoner";
			this.SummName.Name = "SummName";
			this.SummName.ReadOnly = true;
			this.SummName.Width = 130;
			// 
			// RankPlayer
			// 
			this.RankPlayer.HeaderText = "Rank";
			this.RankPlayer.Name = "RankPlayer";
			this.RankPlayer.ReadOnly = true;
			this.RankPlayer.Width = 80;
			// 
			// RolesPlayer
			// 
			this.RolesPlayer.HeaderText = "Roles";
			this.RolesPlayer.Name = "RolesPlayer";
			this.RolesPlayer.ReadOnly = true;
			this.RolesPlayer.Width = 45;
			// 
			// DuoPlayer
			// 
			this.DuoPlayer.HeaderText = "Duo";
			this.DuoPlayer.Name = "DuoPlayer";
			this.DuoPlayer.ReadOnly = true;
			this.DuoPlayer.Width = 134;
			// 
			// button_AddPlayer
			// 
			this.button_AddPlayer.Location = new System.Drawing.Point(6, 6);
			this.button_AddPlayer.Name = "button_AddPlayer";
			this.button_AddPlayer.Size = new System.Drawing.Size(80, 23);
			this.button_AddPlayer.TabIndex = 0;
			this.button_AddPlayer.Text = "Add Player";
			this.button_AddPlayer.UseVisualStyleBackColor = true;
			this.button_AddPlayer.Click += new System.EventHandler(this.button_AddPlayer_Click);
			// 
			// tabPage2
			// 
			this.tabPage2.Controls.Add(this.pictureBox3);
			this.tabPage2.Controls.Add(this.button_Balance);
			this.tabPage2.Controls.Add(this.button_ResetPoints);
			this.tabPage2.Controls.Add(this.groupBox_Setting);
			this.tabPage2.Controls.Add(this.dataGridView_Ranks);
			this.tabPage2.Location = new System.Drawing.Point(4, 22);
			this.tabPage2.Name = "tabPage2";
			this.tabPage2.Padding = new System.Windows.Forms.Padding(3);
			this.tabPage2.Size = new System.Drawing.Size(629, 373);
			this.tabPage2.TabIndex = 1;
			this.tabPage2.Text = "Balancing";
			this.tabPage2.UseVisualStyleBackColor = true;
			// 
			// pictureBox3
			// 
			this.pictureBox3.Image = ((System.Drawing.Image)(resources.GetObject("pictureBox3.Image")));
			this.pictureBox3.Location = new System.Drawing.Point(187, 184);
			this.pictureBox3.Name = "pictureBox3";
			this.pictureBox3.Size = new System.Drawing.Size(436, 183);
			this.pictureBox3.SizeMode = System.Windows.Forms.PictureBoxSizeMode.StretchImage;
			this.pictureBox3.TabIndex = 5;
			this.pictureBox3.TabStop = false;
			// 
			// button_Balance
			// 
			this.button_Balance.Location = new System.Drawing.Point(498, 124);
			this.button_Balance.Name = "button_Balance";
			this.button_Balance.Size = new System.Drawing.Size(125, 23);
			this.button_Balance.TabIndex = 3;
			this.button_Balance.Text = "BALANCE TEAMS!";
			this.button_Balance.UseVisualStyleBackColor = true;
			this.button_Balance.Click += new System.EventHandler(this.button_Balance_Click);
			// 
			// button_ResetPoints
			// 
			this.button_ResetPoints.Location = new System.Drawing.Point(187, 124);
			this.button_ResetPoints.Name = "button_ResetPoints";
			this.button_ResetPoints.Size = new System.Drawing.Size(125, 23);
			this.button_ResetPoints.TabIndex = 2;
			this.button_ResetPoints.Text = "Default Rank -> Points";
			this.button_ResetPoints.UseVisualStyleBackColor = true;
			this.button_ResetPoints.Click += new System.EventHandler(this.button_ResetPoints_Click);
			// 
			// groupBox_Setting
			// 
			this.groupBox_Setting.Controls.Add(this.checkBox_EveryRole);
			this.groupBox_Setting.Controls.Add(this.label2);
			this.groupBox_Setting.Controls.Add(this.numericUpDown_RandNum);
			this.groupBox_Setting.Controls.Add(this.checkBox_BalByRole);
			this.groupBox_Setting.Location = new System.Drawing.Point(187, 6);
			this.groupBox_Setting.Name = "groupBox_Setting";
			this.groupBox_Setting.Size = new System.Drawing.Size(436, 112);
			this.groupBox_Setting.TabIndex = 1;
			this.groupBox_Setting.TabStop = false;
			this.groupBox_Setting.Text = "Settings";
			// 
			// checkBox_EveryRole
			// 
			this.checkBox_EveryRole.AutoSize = true;
			this.checkBox_EveryRole.Checked = true;
			this.checkBox_EveryRole.CheckState = System.Windows.Forms.CheckState.Checked;
			this.checkBox_EveryRole.Location = new System.Drawing.Point(6, 42);
			this.checkBox_EveryRole.Name = "checkBox_EveryRole";
			this.checkBox_EveryRole.Size = new System.Drawing.Size(185, 17);
			this.checkBox_EveryRole.TabIndex = 3;
			this.checkBox_EveryRole.Text = "Generate Teams With Every Role";
			this.checkBox_EveryRole.UseVisualStyleBackColor = true;
			// 
			// label2
			// 
			this.label2.AutoSize = true;
			this.label2.Location = new System.Drawing.Point(47, 64);
			this.label2.Name = "label2";
			this.label2.Size = new System.Drawing.Size(378, 39);
			this.label2.TabIndex = 2;
			this.label2.Text = "Random Variable\r\n(A larger number means more variance, but more inconsistency and" +
    " vice-versa)\r\n(If the Ranks -> Points is default, then default number is general" +
    "ly 5)";
			// 
			// numericUpDown_RandNum
			// 
			this.numericUpDown_RandNum.Location = new System.Drawing.Point(6, 74);
			this.numericUpDown_RandNum.Name = "numericUpDown_RandNum";
			this.numericUpDown_RandNum.Size = new System.Drawing.Size(35, 20);
			this.numericUpDown_RandNum.TabIndex = 1;
			this.numericUpDown_RandNum.Value = new decimal(new int[] {
            5,
            0,
            0,
            0});
			// 
			// checkBox_BalByRole
			// 
			this.checkBox_BalByRole.AutoSize = true;
			this.checkBox_BalByRole.Location = new System.Drawing.Point(6, 19);
			this.checkBox_BalByRole.Name = "checkBox_BalByRole";
			this.checkBox_BalByRole.Size = new System.Drawing.Size(231, 17);
			this.checkBox_BalByRole.TabIndex = 0;
			this.checkBox_BalByRole.Text = "Balance With Emphasis On Preferred Roles";
			this.checkBox_BalByRole.UseVisualStyleBackColor = true;
			// 
			// dataGridView_Ranks
			// 
			this.dataGridView_Ranks.AllowUserToAddRows = false;
			this.dataGridView_Ranks.AllowUserToDeleteRows = false;
			this.dataGridView_Ranks.AllowUserToResizeColumns = false;
			this.dataGridView_Ranks.AllowUserToResizeRows = false;
			this.dataGridView_Ranks.BackgroundColor = System.Drawing.SystemColors.Window;
			this.dataGridView_Ranks.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_Ranks.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.Rank,
            this.Value});
			this.dataGridView_Ranks.Location = new System.Drawing.Point(6, 6);
			this.dataGridView_Ranks.MultiSelect = false;
			this.dataGridView_Ranks.Name = "dataGridView_Ranks";
			this.dataGridView_Ranks.RowHeadersVisible = false;
			this.dataGridView_Ranks.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.CellSelect;
			this.dataGridView_Ranks.Size = new System.Drawing.Size(175, 361);
			this.dataGridView_Ranks.TabIndex = 0;
			// 
			// Rank
			// 
			this.Rank.Frozen = true;
			this.Rank.HeaderText = "Rank";
			this.Rank.Name = "Rank";
			this.Rank.ReadOnly = true;
			this.Rank.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.Rank.Width = 125;
			// 
			// Value
			// 
			this.Value.Frozen = true;
			this.Value.HeaderText = "Pts";
			this.Value.Name = "Value";
			this.Value.SortMode = System.Windows.Forms.DataGridViewColumnSortMode.NotSortable;
			this.Value.Width = 30;
			// 
			// tabPage3
			// 
			this.tabPage3.Controls.Add(this.groupBox4);
			this.tabPage3.Controls.Add(this.groupBox2);
			this.tabPage3.Controls.Add(this.groupBox1);
			this.tabPage3.Location = new System.Drawing.Point(4, 22);
			this.tabPage3.Name = "tabPage3";
			this.tabPage3.Size = new System.Drawing.Size(629, 373);
			this.tabPage3.TabIndex = 2;
			this.tabPage3.Text = "Teams / Stats";
			this.tabPage3.UseVisualStyleBackColor = true;
			// 
			// groupBox2
			// 
			this.groupBox2.Controls.Add(this.label3);
			this.groupBox2.Controls.Add(this.label_TeamPoints);
			this.groupBox2.Controls.Add(this.comboBox_Teams);
			this.groupBox2.Controls.Add(this.dataGridView_Team);
			this.groupBox2.Location = new System.Drawing.Point(3, 62);
			this.groupBox2.Name = "groupBox2";
			this.groupBox2.Size = new System.Drawing.Size(623, 216);
			this.groupBox2.TabIndex = 2;
			this.groupBox2.TabStop = false;
			this.groupBox2.Text = "Teams";
			// 
			// label3
			// 
			this.label3.Location = new System.Drawing.Point(6, 43);
			this.label3.Name = "label3";
			this.label3.Size = new System.Drawing.Size(149, 169);
			this.label3.TabIndex = 3;
			this.label3.Text = "Total Points: \r\nNum. of Teams: \r\nAverage Team Points: ";
			this.label3.TextAlign = System.Drawing.ContentAlignment.MiddleLeft;
			// 
			// label_TeamPoints
			// 
			this.label_TeamPoints.Location = new System.Drawing.Point(161, 179);
			this.label_TeamPoints.Name = "label_TeamPoints";
			this.label_TeamPoints.Size = new System.Drawing.Size(456, 34);
			this.label_TeamPoints.TabIndex = 2;
			this.label_TeamPoints.Text = "No Team Points";
			this.label_TeamPoints.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// comboBox_Teams
			// 
			this.comboBox_Teams.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList;
			this.comboBox_Teams.FormattingEnabled = true;
			this.comboBox_Teams.Items.AddRange(new object[] {
            " "});
			this.comboBox_Teams.Location = new System.Drawing.Point(6, 19);
			this.comboBox_Teams.Name = "comboBox_Teams";
			this.comboBox_Teams.Size = new System.Drawing.Size(149, 21);
			this.comboBox_Teams.TabIndex = 1;
			this.comboBox_Teams.SelectedIndexChanged += new System.EventHandler(this.comboBox_Teams_SelectedIndexChanged);
			// 
			// dataGridView_Team
			// 
			this.dataGridView_Team.AllowUserToAddRows = false;
			this.dataGridView_Team.AllowUserToDeleteRows = false;
			this.dataGridView_Team.BackgroundColor = System.Drawing.SystemColors.Window;
			this.dataGridView_Team.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize;
			this.dataGridView_Team.Columns.AddRange(new System.Windows.Forms.DataGridViewColumn[] {
            this.dataGridViewTextBoxColumn1,
            this.dataGridViewTextBoxColumn2,
            this.RolesPlay,
            this.RankPlay});
			this.dataGridView_Team.Location = new System.Drawing.Point(161, 19);
			this.dataGridView_Team.MultiSelect = false;
			this.dataGridView_Team.Name = "dataGridView_Team";
			this.dataGridView_Team.ReadOnly = true;
			this.dataGridView_Team.RowHeadersVisible = false;
			this.dataGridView_Team.ScrollBars = System.Windows.Forms.ScrollBars.Horizontal;
			this.dataGridView_Team.SelectionMode = System.Windows.Forms.DataGridViewSelectionMode.FullRowSelect;
			this.dataGridView_Team.Size = new System.Drawing.Size(455, 156);
			this.dataGridView_Team.TabIndex = 0;
			// 
			// dataGridViewTextBoxColumn1
			// 
			this.dataGridViewTextBoxColumn1.HeaderText = "Name (Uniq)";
			this.dataGridViewTextBoxColumn1.Name = "dataGridViewTextBoxColumn1";
			this.dataGridViewTextBoxColumn1.ReadOnly = true;
			this.dataGridViewTextBoxColumn1.Width = 177;
			// 
			// dataGridViewTextBoxColumn2
			// 
			this.dataGridViewTextBoxColumn2.HeaderText = "Summoner";
			this.dataGridViewTextBoxColumn2.Name = "dataGridViewTextBoxColumn2";
			this.dataGridViewTextBoxColumn2.ReadOnly = true;
			// 
			// RolesPlay
			// 
			this.RolesPlay.HeaderText = "Roles";
			this.RolesPlay.Name = "RolesPlay";
			this.RolesPlay.ReadOnly = true;
			this.RolesPlay.Width = 75;
			// 
			// RankPlay
			// 
			this.RankPlay.HeaderText = "Rank";
			this.RankPlay.Name = "RankPlay";
			this.RankPlay.ReadOnly = true;
			// 
			// groupBox1
			// 
			this.groupBox1.Controls.Add(this.button_LoadTeams);
			this.groupBox1.Location = new System.Drawing.Point(3, 3);
			this.groupBox1.Name = "groupBox1";
			this.groupBox1.Size = new System.Drawing.Size(623, 53);
			this.groupBox1.TabIndex = 0;
			this.groupBox1.TabStop = false;
			this.groupBox1.Text = "Load Teams";
			// 
			// button_LoadTeams
			// 
			this.button_LoadTeams.Location = new System.Drawing.Point(6, 19);
			this.button_LoadTeams.Name = "button_LoadTeams";
			this.button_LoadTeams.Size = new System.Drawing.Size(75, 23);
			this.button_LoadTeams.TabIndex = 1;
			this.button_LoadTeams.Text = "Load Teams";
			this.button_LoadTeams.UseVisualStyleBackColor = true;
			this.button_LoadTeams.Click += new System.EventHandler(this.button_LoadTeams_Click);
			// 
			// label_Version
			// 
			this.label_Version.Location = new System.Drawing.Point(12, 548);
			this.label_Version.Name = "label_Version";
			this.label_Version.Size = new System.Drawing.Size(633, 23);
			this.label_Version.TabIndex = 5;
			this.label_Version.Text = "VERSION MESSAGE";
			this.label_Version.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// groupBox4
			// 
			this.groupBox4.Controls.Add(this.checkBox_ProdKey);
			this.groupBox4.Controls.Add(this.button_HelpStats);
			this.groupBox4.Controls.Add(this.label5);
			this.groupBox4.Controls.Add(this.textBox_APIKey);
			this.groupBox4.Controls.Add(this.button_GenStats);
			this.groupBox4.Location = new System.Drawing.Point(3, 284);
			this.groupBox4.Name = "groupBox4";
			this.groupBox4.Size = new System.Drawing.Size(623, 80);
			this.groupBox4.TabIndex = 3;
			this.groupBox4.TabStop = false;
			this.groupBox4.Text = "Load Stats";
			// 
			// label5
			// 
			this.label5.Location = new System.Drawing.Point(207, 19);
			this.label5.Name = "label5";
			this.label5.Size = new System.Drawing.Size(102, 20);
			this.label5.TabIndex = 2;
			this.label5.Text = "API Key:";
			this.label5.TextAlign = System.Drawing.ContentAlignment.MiddleRight;
			// 
			// textBox_APIKey
			// 
			this.textBox_APIKey.Location = new System.Drawing.Point(315, 19);
			this.textBox_APIKey.Name = "textBox_APIKey";
			this.textBox_APIKey.Size = new System.Drawing.Size(294, 20);
			this.textBox_APIKey.TabIndex = 2;
			// 
			// button_GenStats
			// 
			this.button_GenStats.Location = new System.Drawing.Point(6, 19);
			this.button_GenStats.Name = "button_GenStats";
			this.button_GenStats.Size = new System.Drawing.Size(93, 55);
			this.button_GenStats.TabIndex = 1;
			this.button_GenStats.Text = "Load .txt Games\r\nGenerate Stats";
			this.button_GenStats.UseVisualStyleBackColor = true;
			this.button_GenStats.Click += new System.EventHandler(this.button_GenStats_Click);
			// 
			// button_HelpStats
			// 
			this.button_HelpStats.Location = new System.Drawing.Point(108, 19);
			this.button_HelpStats.Name = "button_HelpStats";
			this.button_HelpStats.Size = new System.Drawing.Size(93, 55);
			this.button_HelpStats.TabIndex = 3;
			this.button_HelpStats.Text = "Help: Format";
			this.button_HelpStats.UseVisualStyleBackColor = true;
			this.button_HelpStats.Click += new System.EventHandler(this.button_HelpStats_Click);
			// 
			// checkBox_ProdKey
			// 
			this.checkBox_ProdKey.Location = new System.Drawing.Point(210, 50);
			this.checkBox_ProdKey.Name = "checkBox_ProdKey";
			this.checkBox_ProdKey.Size = new System.Drawing.Size(107, 24);
			this.checkBox_ProdKey.TabIndex = 4;
			this.checkBox_ProdKey.Text = "Production Key?";
			this.toolTip1.SetToolTip(this.checkBox_ProdKey, "Check if your API is Production approved by Riot");
			this.checkBox_ProdKey.UseVisualStyleBackColor = true;
			// 
			// toolTip1
			// 
			this.toolTip1.AutoPopDelay = 5000;
			this.toolTip1.InitialDelay = 50;
			this.toolTip1.ReshowDelay = 100;
			// 
			// MainForm
			// 
			this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
			this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
			this.ClientSize = new System.Drawing.Size(661, 575);
			this.Controls.Add(this.label_Version);
			this.Controls.Add(this.tabControl1);
			this.Controls.Add(this.label1);
			this.Controls.Add(this.pictureBox2);
			this.Controls.Add(this.pictureBox1);
			this.Icon = ((System.Drawing.Icon)(resources.GetObject("$this.Icon")));
			this.Name = "MainForm";
			this.Text = "League of Legends Balancer";
			this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
			this.Load += new System.EventHandler(this.MainForm_Load);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox1)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.pictureBox2)).EndInit();
			this.tabControl1.ResumeLayout(false);
			this.tabPage1.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_Players)).EndInit();
			this.tabPage2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.pictureBox3)).EndInit();
			this.groupBox_Setting.ResumeLayout(false);
			this.groupBox_Setting.PerformLayout();
			((System.ComponentModel.ISupportInitialize)(this.numericUpDown_RandNum)).EndInit();
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_Ranks)).EndInit();
			this.tabPage3.ResumeLayout(false);
			this.groupBox2.ResumeLayout(false);
			((System.ComponentModel.ISupportInitialize)(this.dataGridView_Team)).EndInit();
			this.groupBox1.ResumeLayout(false);
			this.groupBox4.ResumeLayout(false);
			this.groupBox4.PerformLayout();
			this.ResumeLayout(false);

		}

		#endregion

		private System.Windows.Forms.PictureBox pictureBox1;
		private System.Windows.Forms.PictureBox pictureBox2;
		private System.Windows.Forms.Label label1;
		private System.Windows.Forms.TabControl tabControl1;
		private System.Windows.Forms.TabPage tabPage1;
		private System.Windows.Forms.TabPage tabPage2;
		private System.Windows.Forms.TabPage tabPage3;
		private System.Windows.Forms.Button button_AddPlayer;
		private System.Windows.Forms.DataGridView dataGridView_Players;
		private System.Windows.Forms.Button button_SavePlayers;
		private System.Windows.Forms.Button button_LoadPlayers;
		private System.Windows.Forms.DataGridView dataGridView_Ranks;
		private System.Windows.Forms.GroupBox groupBox_Setting;
		private System.Windows.Forms.CheckBox checkBox_BalByRole;
		private System.Windows.Forms.Label label2;
		private System.Windows.Forms.NumericUpDown numericUpDown_RandNum;
		private System.Windows.Forms.CheckBox checkBox_EveryRole;
		private System.Windows.Forms.Button button_ResetPoints;
		private System.Windows.Forms.Button button_Balance;
		private System.Windows.Forms.PictureBox pictureBox3;
		private System.Windows.Forms.GroupBox groupBox1;
		private System.Windows.Forms.Button button_LoadTeams;
		private System.Windows.Forms.GroupBox groupBox2;
		private System.Windows.Forms.ComboBox comboBox_Teams;
		private System.Windows.Forms.DataGridView dataGridView_Team;
		private System.Windows.Forms.Label label_TeamPoints;
		private System.Windows.Forms.Label label_Total;
		private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn1;
		private System.Windows.Forms.DataGridViewTextBoxColumn dataGridViewTextBoxColumn2;
		private System.Windows.Forms.DataGridViewTextBoxColumn RolesPlay;
		private System.Windows.Forms.DataGridViewTextBoxColumn RankPlay;
		private System.Windows.Forms.DataGridViewButtonColumn DelPlayer;
		private System.Windows.Forms.DataGridViewTextBoxColumn NamePlayer;
		private System.Windows.Forms.DataGridViewTextBoxColumn UniqPlayer;
		private System.Windows.Forms.DataGridViewTextBoxColumn SummName;
		private System.Windows.Forms.DataGridViewTextBoxColumn RankPlayer;
		private System.Windows.Forms.DataGridViewTextBoxColumn RolesPlayer;
		private System.Windows.Forms.DataGridViewTextBoxColumn DuoPlayer;
		private System.Windows.Forms.Label label_Version;
		private System.Windows.Forms.DataGridViewTextBoxColumn Rank;
		private System.Windows.Forms.DataGridViewTextBoxColumn Value;
		private System.Windows.Forms.Label label3;
		private System.Windows.Forms.GroupBox groupBox4;
		private System.Windows.Forms.Label label5;
		private System.Windows.Forms.TextBox textBox_APIKey;
		private System.Windows.Forms.Button button_GenStats;
		private System.Windows.Forms.Button button_HelpStats;
		private System.Windows.Forms.CheckBox checkBox_ProdKey;
		private System.Windows.Forms.ToolTip toolTip1;
	}
}

