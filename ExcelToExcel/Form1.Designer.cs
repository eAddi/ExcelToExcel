namespace ExcelToExcel
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
            this.TabControl = new System.Windows.Forms.TabControl();
            this.Tab_ExcelFile = new System.Windows.Forms.TabPage();
            this.L_Path = new System.Windows.Forms.Label();
            this.L_1 = new System.Windows.Forms.Label();
            this.BT_Load = new System.Windows.Forms.Button();
            this.Tab_Hierarchy = new System.Windows.Forms.TabPage();
            this.groupBox4 = new System.Windows.Forms.GroupBox();
            this.BT_Create_Node = new System.Windows.Forms.Button();
            this.myTreeView = new System.Windows.Forms.TreeView();
            this.BT_collapse = new System.Windows.Forms.Button();
            this.BT_Delete_Node = new System.Windows.Forms.Button();
            this.BT_expand = new System.Windows.Forms.Button();
            this.groupBox2 = new System.Windows.Forms.GroupBox();
            this.LB_Entetes = new System.Windows.Forms.ListBox();
            this.BT_Add_Property = new System.Windows.Forms.Button();
            this.BT_Delete_Property = new System.Windows.Forms.Button();
            this.Tab_Generate = new System.Windows.Forms.TabPage();
            this.groupBox3 = new System.Windows.Forms.GroupBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.BT_FE_Confirm = new System.Windows.Forms.Button();
            this.TB_FE_BR = new System.Windows.Forms.TextBox();
            this.TB_FE_UL = new System.Windows.Forms.TextBox();
            this.TB_FE_Name = new System.Windows.Forms.TextBox();
            this.L_FE = new System.Windows.Forms.Label();
            this.RB_FE_False = new System.Windows.Forms.RadioButton();
            this.RB_FE_True = new System.Windows.Forms.RadioButton();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.RB_New_Excel = new System.Windows.Forms.RadioButton();
            this.RB_Existing_Excel = new System.Windows.Forms.RadioButton();
            this.BT_generate = new System.Windows.Forms.Button();
            this.OF_xls = new System.Windows.Forms.OpenFileDialog();
            this.TabControl.SuspendLayout();
            this.Tab_ExcelFile.SuspendLayout();
            this.Tab_Hierarchy.SuspendLayout();
            this.groupBox4.SuspendLayout();
            this.groupBox2.SuspendLayout();
            this.Tab_Generate.SuspendLayout();
            this.groupBox3.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // TabControl
            // 
            this.TabControl.Controls.Add(this.Tab_ExcelFile);
            this.TabControl.Controls.Add(this.Tab_Hierarchy);
            this.TabControl.Controls.Add(this.Tab_Generate);
            this.TabControl.Location = new System.Drawing.Point(-1, 0);
            this.TabControl.Name = "TabControl";
            this.TabControl.SelectedIndex = 0;
            this.TabControl.Size = new System.Drawing.Size(1026, 564);
            this.TabControl.TabIndex = 0;
            // 
            // Tab_ExcelFile
            // 
            this.Tab_ExcelFile.Controls.Add(this.L_Path);
            this.Tab_ExcelFile.Controls.Add(this.L_1);
            this.Tab_ExcelFile.Controls.Add(this.BT_Load);
            this.Tab_ExcelFile.Location = new System.Drawing.Point(4, 22);
            this.Tab_ExcelFile.Name = "Tab_ExcelFile";
            this.Tab_ExcelFile.Padding = new System.Windows.Forms.Padding(3);
            this.Tab_ExcelFile.Size = new System.Drawing.Size(1018, 538);
            this.Tab_ExcelFile.TabIndex = 0;
            this.Tab_ExcelFile.Text = "1 - Fichier Excel";
            this.Tab_ExcelFile.UseVisualStyleBackColor = true;
            // 
            // L_Path
            // 
            this.L_Path.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.L_Path.Location = new System.Drawing.Point(228, 327);
            this.L_Path.Name = "L_Path";
            this.L_Path.Size = new System.Drawing.Size(563, 94);
            this.L_Path.TabIndex = 2;
            this.L_Path.TextAlign = System.Drawing.ContentAlignment.TopCenter;
            // 
            // L_1
            // 
            this.L_1.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.L_1.AutoSize = true;
            this.L_1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.L_1.Location = new System.Drawing.Point(413, 199);
            this.L_1.Name = "L_1";
            this.L_1.Size = new System.Drawing.Size(193, 18);
            this.L_1.TabIndex = 1;
            this.L_1.Text = "Sélectionner un fichier Excel";
            // 
            // BT_Load
            // 
            this.BT_Load.Anchor = System.Windows.Forms.AnchorStyles.None;
            this.BT_Load.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_Load.Location = new System.Drawing.Point(457, 233);
            this.BT_Load.Name = "BT_Load";
            this.BT_Load.Size = new System.Drawing.Size(104, 54);
            this.BT_Load.TabIndex = 0;
            this.BT_Load.Text = "Charger";
            this.BT_Load.UseVisualStyleBackColor = true;
            this.BT_Load.Click += new System.EventHandler(this.BT_Load_Click);
            // 
            // Tab_Hierarchy
            // 
            this.Tab_Hierarchy.Controls.Add(this.groupBox4);
            this.Tab_Hierarchy.Controls.Add(this.groupBox2);
            this.Tab_Hierarchy.Location = new System.Drawing.Point(4, 22);
            this.Tab_Hierarchy.Name = "Tab_Hierarchy";
            this.Tab_Hierarchy.Padding = new System.Windows.Forms.Padding(3);
            this.Tab_Hierarchy.Size = new System.Drawing.Size(1018, 538);
            this.Tab_Hierarchy.TabIndex = 1;
            this.Tab_Hierarchy.Text = "2 - Hiérarchie";
            this.Tab_Hierarchy.UseVisualStyleBackColor = true;
            // 
            // groupBox4
            // 
            this.groupBox4.Controls.Add(this.BT_Create_Node);
            this.groupBox4.Controls.Add(this.myTreeView);
            this.groupBox4.Controls.Add(this.BT_collapse);
            this.groupBox4.Controls.Add(this.BT_Delete_Node);
            this.groupBox4.Controls.Add(this.BT_expand);
            this.groupBox4.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox4.Location = new System.Drawing.Point(394, 19);
            this.groupBox4.Name = "groupBox4";
            this.groupBox4.Size = new System.Drawing.Size(602, 507);
            this.groupBox4.TabIndex = 7;
            this.groupBox4.TabStop = false;
            this.groupBox4.Text = "2 - Noeuds";
            // 
            // BT_Create_Node
            // 
            this.BT_Create_Node.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_Create_Node.Location = new System.Drawing.Point(24, 42);
            this.BT_Create_Node.Name = "BT_Create_Node";
            this.BT_Create_Node.Size = new System.Drawing.Size(75, 34);
            this.BT_Create_Node.TabIndex = 1;
            this.BT_Create_Node.Text = "Ajouter";
            this.BT_Create_Node.UseVisualStyleBackColor = true;
            this.BT_Create_Node.Click += new System.EventHandler(this.BT_Create_Click);
            // 
            // myTreeView
            // 
            this.myTreeView.CheckBoxes = true;
            this.myTreeView.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.myTreeView.Location = new System.Drawing.Point(119, 30);
            this.myTreeView.Name = "myTreeView";
            this.myTreeView.PathSeparator = ".";
            this.myTreeView.ShowNodeToolTips = true;
            this.myTreeView.Size = new System.Drawing.Size(463, 459);
            this.myTreeView.TabIndex = 0;
            // 
            // BT_collapse
            // 
            this.BT_collapse.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_collapse.Location = new System.Drawing.Point(96, 466);
            this.BT_collapse.Name = "BT_collapse";
            this.BT_collapse.Size = new System.Drawing.Size(17, 23);
            this.BT_collapse.TabIndex = 5;
            this.BT_collapse.Text = "-";
            this.BT_collapse.UseVisualStyleBackColor = true;
            this.BT_collapse.Click += new System.EventHandler(this.BT_collapse_Click);
            // 
            // BT_Delete_Node
            // 
            this.BT_Delete_Node.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_Delete_Node.Location = new System.Drawing.Point(24, 82);
            this.BT_Delete_Node.Name = "BT_Delete_Node";
            this.BT_Delete_Node.Size = new System.Drawing.Size(75, 34);
            this.BT_Delete_Node.TabIndex = 2;
            this.BT_Delete_Node.Text = "Supprimer";
            this.BT_Delete_Node.UseVisualStyleBackColor = true;
            this.BT_Delete_Node.Click += new System.EventHandler(this.BT_Delete_Node_Click);
            // 
            // BT_expand
            // 
            this.BT_expand.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_expand.Location = new System.Drawing.Point(96, 437);
            this.BT_expand.Name = "BT_expand";
            this.BT_expand.RightToLeft = System.Windows.Forms.RightToLeft.No;
            this.BT_expand.Size = new System.Drawing.Size(17, 23);
            this.BT_expand.TabIndex = 4;
            this.BT_expand.Text = "+";
            this.BT_expand.UseVisualStyleBackColor = true;
            this.BT_expand.Click += new System.EventHandler(this.BT_expand_Click);
            // 
            // groupBox2
            // 
            this.groupBox2.Controls.Add(this.LB_Entetes);
            this.groupBox2.Controls.Add(this.BT_Add_Property);
            this.groupBox2.Controls.Add(this.BT_Delete_Property);
            this.groupBox2.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox2.Location = new System.Drawing.Point(19, 17);
            this.groupBox2.Name = "groupBox2";
            this.groupBox2.Size = new System.Drawing.Size(334, 509);
            this.groupBox2.TabIndex = 6;
            this.groupBox2.TabStop = false;
            this.groupBox2.Text = "1 - Entêtes";
            // 
            // LB_Entetes
            // 
            this.LB_Entetes.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.LB_Entetes.FormattingEnabled = true;
            this.LB_Entetes.ItemHeight = 15;
            this.LB_Entetes.Location = new System.Drawing.Point(18, 32);
            this.LB_Entetes.Name = "LB_Entetes";
            this.LB_Entetes.SelectionMode = System.Windows.Forms.SelectionMode.MultiExtended;
            this.LB_Entetes.Size = new System.Drawing.Size(199, 454);
            this.LB_Entetes.TabIndex = 0;
            // 
            // BT_Add_Property
            // 
            this.BT_Add_Property.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_Add_Property.Location = new System.Drawing.Point(241, 44);
            this.BT_Add_Property.Name = "BT_Add_Property";
            this.BT_Add_Property.Size = new System.Drawing.Size(75, 34);
            this.BT_Add_Property.TabIndex = 1;
            this.BT_Add_Property.Text = "Ajouter";
            this.BT_Add_Property.UseVisualStyleBackColor = true;
            this.BT_Add_Property.Click += new System.EventHandler(this.BT_Add_Property_Click);
            // 
            // BT_Delete_Property
            // 
            this.BT_Delete_Property.Font = new System.Drawing.Font("Microsoft Sans Serif", 9F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_Delete_Property.Location = new System.Drawing.Point(241, 84);
            this.BT_Delete_Property.Name = "BT_Delete_Property";
            this.BT_Delete_Property.Size = new System.Drawing.Size(75, 34);
            this.BT_Delete_Property.TabIndex = 2;
            this.BT_Delete_Property.Text = "Supprimer";
            this.BT_Delete_Property.UseVisualStyleBackColor = true;
            this.BT_Delete_Property.Click += new System.EventHandler(this.BT_Delete_Property_Click);
            // 
            // Tab_Generate
            // 
            this.Tab_Generate.Controls.Add(this.groupBox3);
            this.Tab_Generate.Controls.Add(this.groupBox1);
            this.Tab_Generate.Controls.Add(this.BT_generate);
            this.Tab_Generate.Location = new System.Drawing.Point(4, 22);
            this.Tab_Generate.Name = "Tab_Generate";
            this.Tab_Generate.Padding = new System.Windows.Forms.Padding(3);
            this.Tab_Generate.Size = new System.Drawing.Size(1018, 538);
            this.Tab_Generate.TabIndex = 2;
            this.Tab_Generate.Text = "3 - Générer";
            this.Tab_Generate.UseVisualStyleBackColor = true;
            // 
            // groupBox3
            // 
            this.groupBox3.Controls.Add(this.label1);
            this.groupBox3.Controls.Add(this.label3);
            this.groupBox3.Controls.Add(this.BT_FE_Confirm);
            this.groupBox3.Controls.Add(this.TB_FE_BR);
            this.groupBox3.Controls.Add(this.TB_FE_UL);
            this.groupBox3.Controls.Add(this.TB_FE_Name);
            this.groupBox3.Controls.Add(this.L_FE);
            this.groupBox3.Controls.Add(this.RB_FE_False);
            this.groupBox3.Controls.Add(this.RB_FE_True);
            this.groupBox3.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox3.Location = new System.Drawing.Point(120, 53);
            this.groupBox3.Name = "groupBox3";
            this.groupBox3.Size = new System.Drawing.Size(514, 150);
            this.groupBox3.TabIndex = 10;
            this.groupBox3.TabStop = false;
            this.groupBox3.Text = "Fiche Espace";
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(198, 84);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(145, 18);
            this.label1.TabIndex = 17;
            this.label1.Text = "Portée des cellules : ";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(282, 108);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(21, 18);
            this.label3.TabIndex = 16;
            this.label3.Text = " - ";
            // 
            // BT_FE_Confirm
            // 
            this.BT_FE_Confirm.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_FE_Confirm.Location = new System.Drawing.Point(396, 105);
            this.BT_FE_Confirm.Name = "BT_FE_Confirm";
            this.BT_FE_Confirm.Size = new System.Drawing.Size(94, 24);
            this.BT_FE_Confirm.TabIndex = 15;
            this.BT_FE_Confirm.Text = "Confirmer";
            this.BT_FE_Confirm.UseVisualStyleBackColor = true;
            this.BT_FE_Confirm.Click += new System.EventHandler(this.BT_FE_Confirm_Click);
            // 
            // TB_FE_BR
            // 
            this.TB_FE_BR.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.TB_FE_BR.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TB_FE_BR.Location = new System.Drawing.Point(305, 105);
            this.TB_FE_BR.Name = "TB_FE_BR";
            this.TB_FE_BR.Size = new System.Drawing.Size(80, 22);
            this.TB_FE_BR.TabIndex = 14;
            this.TB_FE_BR.Text = "FF667";
            // 
            // TB_FE_UL
            // 
            this.TB_FE_UL.CharacterCasing = System.Windows.Forms.CharacterCasing.Upper;
            this.TB_FE_UL.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TB_FE_UL.Location = new System.Drawing.Point(201, 105);
            this.TB_FE_UL.Name = "TB_FE_UL";
            this.TB_FE_UL.Size = new System.Drawing.Size(80, 22);
            this.TB_FE_UL.TabIndex = 13;
            this.TB_FE_UL.Text = "A2";
            // 
            // TB_FE_Name
            // 
            this.TB_FE_Name.Font = new System.Drawing.Font("Microsoft Sans Serif", 9.75F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.TB_FE_Name.Location = new System.Drawing.Point(24, 105);
            this.TB_FE_Name.Name = "TB_FE_Name";
            this.TB_FE_Name.Size = new System.Drawing.Size(147, 22);
            this.TB_FE_Name.TabIndex = 11;
            this.TB_FE_Name.Text = "Fiches Espaces";
            // 
            // L_FE
            // 
            this.L_FE.AutoSize = true;
            this.L_FE.Location = new System.Drawing.Point(21, 84);
            this.L_FE.Name = "L_FE";
            this.L_FE.Size = new System.Drawing.Size(49, 18);
            this.L_FE.TabIndex = 10;
            this.L_FE.Text = "Nom :";
            // 
            // RB_FE_False
            // 
            this.RB_FE_False.AutoSize = true;
            this.RB_FE_False.Location = new System.Drawing.Point(22, 52);
            this.RB_FE_False.Name = "RB_FE_False";
            this.RB_FE_False.Size = new System.Drawing.Size(251, 22);
            this.RB_FE_False.TabIndex = 1;
            this.RB_FE_False.Text = "Ne Possède pas une fiche espace";
            this.RB_FE_False.UseVisualStyleBackColor = true;
            // 
            // RB_FE_True
            // 
            this.RB_FE_True.AutoSize = true;
            this.RB_FE_True.Checked = true;
            this.RB_FE_True.Location = new System.Drawing.Point(22, 29);
            this.RB_FE_True.Name = "RB_FE_True";
            this.RB_FE_True.Size = new System.Drawing.Size(200, 22);
            this.RB_FE_True.TabIndex = 0;
            this.RB_FE_True.TabStop = true;
            this.RB_FE_True.Text = "Possède une fiche espace";
            this.RB_FE_True.UseVisualStyleBackColor = true;
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.RB_New_Excel);
            this.groupBox1.Controls.Add(this.RB_Existing_Excel);
            this.groupBox1.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.groupBox1.Location = new System.Drawing.Point(120, 238);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(514, 107);
            this.groupBox1.TabIndex = 8;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "Exporter vers";
            // 
            // RB_New_Excel
            // 
            this.RB_New_Excel.AutoSize = true;
            this.RB_New_Excel.Location = new System.Drawing.Point(38, 36);
            this.RB_New_Excel.Name = "RB_New_Excel";
            this.RB_New_Excel.Size = new System.Drawing.Size(188, 22);
            this.RB_New_Excel.TabIndex = 5;
            this.RB_New_Excel.TabStop = true;
            this.RB_New_Excel.Text = "Un nouveau fichier Excel";
            this.RB_New_Excel.UseVisualStyleBackColor = true;
            // 
            // RB_Existing_Excel
            // 
            this.RB_Existing_Excel.AutoSize = true;
            this.RB_Existing_Excel.Checked = true;
            this.RB_Existing_Excel.Location = new System.Drawing.Point(38, 59);
            this.RB_Existing_Excel.Name = "RB_Existing_Excel";
            this.RB_Existing_Excel.Size = new System.Drawing.Size(179, 22);
            this.RB_Existing_Excel.TabIndex = 6;
            this.RB_Existing_Excel.TabStop = true;
            this.RB_Existing_Excel.Text = "Le fichier Excel existant";
            this.RB_Existing_Excel.UseVisualStyleBackColor = true;
            // 
            // BT_generate
            // 
            this.BT_generate.Font = new System.Drawing.Font("Microsoft Sans Serif", 11.25F, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, ((byte)(0)));
            this.BT_generate.Location = new System.Drawing.Point(304, 383);
            this.BT_generate.Name = "BT_generate";
            this.BT_generate.Size = new System.Drawing.Size(146, 52);
            this.BT_generate.TabIndex = 4;
            this.BT_generate.Text = "Générer";
            this.BT_generate.UseVisualStyleBackColor = true;
            this.BT_generate.Click += new System.EventHandler(this.BT_Generate_Click);
            // 
            // OF_xls
            // 
            this.OF_xls.RestoreDirectory = true;
            this.OF_xls.FileOk += new System.ComponentModel.CancelEventHandler(this.OF_xls_FileOk);
            // 
            // MainForm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 13F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(1022, 560);
            this.Controls.Add(this.TabControl);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedSingle;
            this.Name = "MainForm";
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "ExcelTransformer";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.MainForm_FormClosing);
            this.TabControl.ResumeLayout(false);
            this.Tab_ExcelFile.ResumeLayout(false);
            this.Tab_ExcelFile.PerformLayout();
            this.Tab_Hierarchy.ResumeLayout(false);
            this.groupBox4.ResumeLayout(false);
            this.groupBox2.ResumeLayout(false);
            this.Tab_Generate.ResumeLayout(false);
            this.groupBox3.ResumeLayout(false);
            this.groupBox3.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.TabControl TabControl;
        private System.Windows.Forms.TabPage Tab_ExcelFile;
        private System.Windows.Forms.Label L_Path;
        private System.Windows.Forms.Label L_1;
        private System.Windows.Forms.Button BT_Load;
        private System.Windows.Forms.TabPage Tab_Hierarchy;
        private System.Windows.Forms.OpenFileDialog OF_xls;
        private System.Windows.Forms.TabPage Tab_Generate;
        private System.Windows.Forms.Button BT_Delete_Node;
        private System.Windows.Forms.Button BT_Create_Node;
        private System.Windows.Forms.TreeView myTreeView;
        private System.Windows.Forms.Button BT_collapse;
        private System.Windows.Forms.Button BT_expand;
        private System.Windows.Forms.RadioButton RB_Existing_Excel;
        private System.Windows.Forms.RadioButton RB_New_Excel;
        private System.Windows.Forms.Button BT_generate;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.GroupBox groupBox2;
        private System.Windows.Forms.Button BT_Delete_Property;
        private System.Windows.Forms.Button BT_Add_Property;
        private System.Windows.Forms.ListBox LB_Entetes;
        private System.Windows.Forms.GroupBox groupBox3;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.Button BT_FE_Confirm;
        private System.Windows.Forms.TextBox TB_FE_BR;
        private System.Windows.Forms.TextBox TB_FE_UL;
        private System.Windows.Forms.TextBox TB_FE_Name;
        private System.Windows.Forms.Label L_FE;
        private System.Windows.Forms.RadioButton RB_FE_False;
        private System.Windows.Forms.RadioButton RB_FE_True;
        private System.Windows.Forms.GroupBox groupBox4;
    }
}

