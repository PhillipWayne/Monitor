namespace MonitorPorts
{
    partial class FilterForm
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
            this.GroupBox_Protocol = new System.Windows.Forms.GroupBox();
            this.GroupBox_Source = new System.Windows.Forms.GroupBox();
            this.DestListBox2 = new System.Windows.Forms.CheckedListBox();
            this.SourceListBox1 = new System.Windows.Forms.CheckedListBox();
            this.label1 = new System.Windows.Forms.Label();
            this.label_SourceIP = new System.Windows.Forms.Label();
            this.button_OK = new System.Windows.Forms.Button();
            this.button_Cancel = new System.Windows.Forms.Button();
            this.groupBox1 = new System.Windows.Forms.GroupBox();
            this.label5 = new System.Windows.Forms.Label();
            this.label4 = new System.Windows.Forms.Label();
            this.label3 = new System.Windows.Forms.Label();
            this.textBox2 = new System.Windows.Forms.TextBox();
            this.label2 = new System.Windows.Forms.Label();
            this.textBox1 = new System.Windows.Forms.TextBox();
            this.txtPcapLength = new System.Windows.Forms.TextBox();
            this.Unit1 = new System.Windows.Forms.CheckBox();
            this.Unit2 = new System.Windows.Forms.CheckBox();
            this.GroupBox_Protocol.SuspendLayout();
            this.GroupBox_Source.SuspendLayout();
            this.groupBox1.SuspendLayout();
            this.SuspendLayout();
            // 
            // GroupBox_Protocol
            // 
            this.GroupBox_Protocol.Controls.Add(this.Unit2);
            this.GroupBox_Protocol.Controls.Add(this.Unit1);
            this.GroupBox_Protocol.Controls.Add(this.txtPcapLength);
            this.GroupBox_Protocol.Dock = System.Windows.Forms.DockStyle.Top;
            this.GroupBox_Protocol.Location = new System.Drawing.Point(0, 0);
            this.GroupBox_Protocol.Name = "GroupBox_Protocol";
            this.GroupBox_Protocol.Size = new System.Drawing.Size(279, 45);
            this.GroupBox_Protocol.TabIndex = 4;
            this.GroupBox_Protocol.TabStop = false;
            this.GroupBox_Protocol.Text = "保存pcap文件大小";
            // 
            // GroupBox_Source
            // 
            this.GroupBox_Source.Controls.Add(this.DestListBox2);
            this.GroupBox_Source.Controls.Add(this.SourceListBox1);
            this.GroupBox_Source.Controls.Add(this.label1);
            this.GroupBox_Source.Controls.Add(this.label_SourceIP);
            this.GroupBox_Source.Dock = System.Windows.Forms.DockStyle.Top;
            this.GroupBox_Source.Location = new System.Drawing.Point(0, 45);
            this.GroupBox_Source.Name = "GroupBox_Source";
            this.GroupBox_Source.Size = new System.Drawing.Size(279, 188);
            this.GroupBox_Source.TabIndex = 5;
            this.GroupBox_Source.TabStop = false;
            this.GroupBox_Source.Text = "IP地址筛选";
            // 
            // DestListBox2
            // 
            this.DestListBox2.FormattingEnabled = true;
            this.DestListBox2.Location = new System.Drawing.Point(147, 45);
            this.DestListBox2.Name = "DestListBox2";
            this.DestListBox2.Size = new System.Drawing.Size(123, 132);
            this.DestListBox2.TabIndex = 5;
            // 
            // SourceListBox1
            // 
            this.SourceListBox1.FormattingEnabled = true;
            this.SourceListBox1.Location = new System.Drawing.Point(14, 45);
            this.SourceListBox1.Name = "SourceListBox1";
            this.SourceListBox1.Size = new System.Drawing.Size(123, 132);
            this.SourceListBox1.TabIndex = 5;
            // 
            // label1
            // 
            this.label1.AutoSize = true;
            this.label1.Location = new System.Drawing.Point(150, 23);
            this.label1.Name = "label1";
            this.label1.Size = new System.Drawing.Size(53, 12);
            this.label1.TabIndex = 0;
            this.label1.Text = "目的IP：";
            // 
            // label_SourceIP
            // 
            this.label_SourceIP.AutoSize = true;
            this.label_SourceIP.Location = new System.Drawing.Point(14, 23);
            this.label_SourceIP.Name = "label_SourceIP";
            this.label_SourceIP.Size = new System.Drawing.Size(41, 12);
            this.label_SourceIP.TabIndex = 0;
            this.label_SourceIP.Text = "源IP：";
            // 
            // button_OK
            // 
            this.button_OK.Location = new System.Drawing.Point(114, 314);
            this.button_OK.Name = "button_OK";
            this.button_OK.Size = new System.Drawing.Size(74, 26);
            this.button_OK.TabIndex = 7;
            this.button_OK.Text = "确定";
            this.button_OK.UseVisualStyleBackColor = true;
            this.button_OK.Click += new System.EventHandler(this.button_OK_Click);
            // 
            // button_Cancel
            // 
            this.button_Cancel.DialogResult = System.Windows.Forms.DialogResult.Cancel;
            this.button_Cancel.Location = new System.Drawing.Point(194, 314);
            this.button_Cancel.Name = "button_Cancel";
            this.button_Cancel.Size = new System.Drawing.Size(74, 26);
            this.button_Cancel.TabIndex = 8;
            this.button_Cancel.Text = "取消";
            this.button_Cancel.UseVisualStyleBackColor = true;
            this.button_Cancel.Click += new System.EventHandler(this.button_Cancel_Click);
            // 
            // groupBox1
            // 
            this.groupBox1.Controls.Add(this.label5);
            this.groupBox1.Controls.Add(this.label4);
            this.groupBox1.Controls.Add(this.label3);
            this.groupBox1.Controls.Add(this.textBox2);
            this.groupBox1.Controls.Add(this.label2);
            this.groupBox1.Controls.Add(this.textBox1);
            this.groupBox1.Dock = System.Windows.Forms.DockStyle.Top;
            this.groupBox1.Location = new System.Drawing.Point(0, 233);
            this.groupBox1.Name = "groupBox1";
            this.groupBox1.Size = new System.Drawing.Size(279, 75);
            this.groupBox1.TabIndex = 9;
            this.groupBox1.TabStop = false;
            this.groupBox1.Text = "告警门限";
            // 
            // label5
            // 
            this.label5.AutoSize = true;
            this.label5.Location = new System.Drawing.Point(14, 49);
            this.label5.Name = "label5";
            this.label5.Size = new System.Drawing.Size(77, 12);
            this.label5.TabIndex = 1;
            this.label5.Text = "ZC发送给VOBC";
            // 
            // label4
            // 
            this.label4.AutoSize = true;
            this.label4.Location = new System.Drawing.Point(161, 49);
            this.label4.Name = "label4";
            this.label4.Size = new System.Drawing.Size(65, 12);
            this.label4.TabIndex = 1;
            this.label4.Text = "单位（ms）";
            // 
            // label3
            // 
            this.label3.AutoSize = true;
            this.label3.Location = new System.Drawing.Point(14, 23);
            this.label3.Name = "label3";
            this.label3.Size = new System.Drawing.Size(77, 12);
            this.label3.TabIndex = 1;
            this.label3.Text = "VOBC发送给ZC";
            // 
            // textBox2
            // 
            this.textBox2.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox2.Location = new System.Drawing.Point(99, 45);
            this.textBox2.Name = "textBox2";
            this.textBox2.Size = new System.Drawing.Size(56, 21);
            this.textBox2.TabIndex = 0;
            this.textBox2.Text = "2400";
            this.textBox2.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // label2
            // 
            this.label2.AutoSize = true;
            this.label2.Location = new System.Drawing.Point(161, 23);
            this.label2.Name = "label2";
            this.label2.Size = new System.Drawing.Size(65, 12);
            this.label2.TabIndex = 1;
            this.label2.Text = "单位（ms）";
            // 
            // textBox1
            // 
            this.textBox1.Anchor = ((System.Windows.Forms.AnchorStyles)((System.Windows.Forms.AnchorStyles.Top | System.Windows.Forms.AnchorStyles.Right)));
            this.textBox1.Location = new System.Drawing.Point(99, 19);
            this.textBox1.Name = "textBox1";
            this.textBox1.Size = new System.Drawing.Size(56, 21);
            this.textBox1.TabIndex = 0;
            this.textBox1.Text = "2400";
            this.textBox1.TextAlign = System.Windows.Forms.HorizontalAlignment.Right;
            // 
            // txtPcapLength
            // 
            this.txtPcapLength.Location = new System.Drawing.Point(51, 16);
            this.txtPcapLength.Name = "txtPcapLength";
            this.txtPcapLength.Size = new System.Drawing.Size(75, 21);
            this.txtPcapLength.TabIndex = 0;
            // 
            // Unit1
            // 
            this.Unit1.AutoSize = true;
            this.Unit1.Checked = true;
            this.Unit1.CheckState = System.Windows.Forms.CheckState.Checked;
            this.Unit1.Location = new System.Drawing.Point(147, 18);
            this.Unit1.Name = "Unit1";
            this.Unit1.Size = new System.Drawing.Size(36, 16);
            this.Unit1.TabIndex = 1;
            this.Unit1.Text = "MB";
            this.Unit1.UseVisualStyleBackColor = true;
            this.Unit1.CheckedChanged += new System.EventHandler(this.Unit1_CheckedChanged);
            // 
            // Unit2
            // 
            this.Unit2.AutoSize = true;
            this.Unit2.Location = new System.Drawing.Point(189, 18);
            this.Unit2.Name = "Unit2";
            this.Unit2.Size = new System.Drawing.Size(36, 16);
            this.Unit2.TabIndex = 2;
            this.Unit2.Text = "GB";
            this.Unit2.UseVisualStyleBackColor = true;
            this.Unit2.CheckedChanged += new System.EventHandler(this.Unit2_CheckedChanged);
            // 
            // FilterForm
            // 
            this.AcceptButton = this.button_OK;
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.AutoSize = true;
            this.CancelButton = this.button_Cancel;
            this.ClientSize = new System.Drawing.Size(279, 342);
            this.Controls.Add(this.groupBox1);
            this.Controls.Add(this.button_OK);
            this.Controls.Add(this.button_Cancel);
            this.Controls.Add(this.GroupBox_Source);
            this.Controls.Add(this.GroupBox_Protocol);
            this.FormBorderStyle = System.Windows.Forms.FormBorderStyle.FixedDialog;
            this.MaximizeBox = false;
            this.MinimizeBox = false;
            this.Name = "FilterForm";
            this.ShowIcon = false;
            this.ShowInTaskbar = false;
            this.SizeGripStyle = System.Windows.Forms.SizeGripStyle.Hide;
            this.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen;
            this.Text = "新建设置";
            this.FormClosing += new System.Windows.Forms.FormClosingEventHandler(this.FilterForm_FormClosing);
            this.GroupBox_Protocol.ResumeLayout(false);
            this.GroupBox_Protocol.PerformLayout();
            this.GroupBox_Source.ResumeLayout(false);
            this.GroupBox_Source.PerformLayout();
            this.groupBox1.ResumeLayout(false);
            this.groupBox1.PerformLayout();
            this.ResumeLayout(false);

        }

        #endregion

        private System.Windows.Forms.GroupBox GroupBox_Protocol;
        private System.Windows.Forms.GroupBox GroupBox_Source;
        private System.Windows.Forms.Label label_SourceIP;
        private System.Windows.Forms.Button button_OK;
        private System.Windows.Forms.Button button_Cancel;
        private System.Windows.Forms.CheckedListBox DestListBox2;
        private System.Windows.Forms.CheckedListBox SourceListBox1;
        private System.Windows.Forms.Label label1;
        private System.Windows.Forms.GroupBox groupBox1;
        private System.Windows.Forms.TextBox textBox1;
        private System.Windows.Forms.Label label2;
        private System.Windows.Forms.Label label5;
        private System.Windows.Forms.Label label4;
        private System.Windows.Forms.Label label3;
        private System.Windows.Forms.TextBox textBox2;
        private System.Windows.Forms.CheckBox Unit2;
        private System.Windows.Forms.CheckBox Unit1;
        private System.Windows.Forms.TextBox txtPcapLength;

    }
}