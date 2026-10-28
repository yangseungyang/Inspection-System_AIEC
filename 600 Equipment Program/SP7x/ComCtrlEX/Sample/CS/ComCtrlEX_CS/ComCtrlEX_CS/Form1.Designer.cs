namespace ComCtrlEX_CS
{
    partial class MainFrm
    {
        /// <summary>
        /// 必要なデザイナ変数です。
        /// </summary>
        private System.ComponentModel.IContainer components = null;

        /// <summary>
        /// 使用中のリソースをすべてクリーンアップします。
        /// </summary>
        /// <param name="disposing">マネージ リソースが破棄される場合 true、破棄されない場合は false です。</param>
        protected override void Dispose(bool disposing)
        {
            if (disposing && (components != null))
            {
                components.Dispose();
            }
            base.Dispose(disposing);
        }

        #region Windows フォーム デザイナで生成されたコード

        /// <summary>
        /// デザイナ サポートに必要なメソッドです。このメソッドの内容を
        /// コード エディタで変更しないでください。
        /// </summary>
        private void InitializeComponent()
        {
            this.components = new System.ComponentModel.Container();
            this.txt_PortNo = new System.Windows.Forms.TextBox();
            this.btn_PortNo = new System.Windows.Forms.Button();
            this.txt_Data = new System.Windows.Forms.TextBox();
            this.btn_Start = new System.Windows.Forms.Button();
            this.btn_End = new System.Windows.Forms.Button();
            this.timer1 = new System.Windows.Forms.Timer(this.components);
            this.SuspendLayout();
            // 
            // txt_PortNo
            // 
            this.txt_PortNo.Location = new System.Drawing.Point(12, 9);
            this.txt_PortNo.Name = "txt_PortNo";
            this.txt_PortNo.Size = new System.Drawing.Size(97, 19);
            this.txt_PortNo.TabIndex = 0;
            // 
            // btn_PortNo
            // 
            this.btn_PortNo.Location = new System.Drawing.Point(115, 9);
            this.btn_PortNo.Name = "btn_PortNo";
            this.btn_PortNo.Size = new System.Drawing.Size(76, 19);
            this.btn_PortNo.TabIndex = 1;
            this.btn_PortNo.Text = "ポート設定";
            this.btn_PortNo.UseVisualStyleBackColor = true;
            this.btn_PortNo.Click += new System.EventHandler(this.btn_PortNo_Click);
            // 
            // txt_Data
            // 
            this.txt_Data.Location = new System.Drawing.Point(13, 49);
            this.txt_Data.Name = "txt_Data";
            this.txt_Data.Size = new System.Drawing.Size(95, 19);
            this.txt_Data.TabIndex = 2;
            // 
            // btn_Start
            // 
            this.btn_Start.Location = new System.Drawing.Point(114, 48);
            this.btn_Start.Name = "btn_Start";
            this.btn_Start.Size = new System.Drawing.Size(77, 20);
            this.btn_Start.TabIndex = 3;
            this.btn_Start.Text = "モニタ開始";
            this.btn_Start.UseVisualStyleBackColor = true;
            this.btn_Start.Click += new System.EventHandler(this.btn_Start_Click);
            // 
            // btn_End
            // 
            this.btn_End.Location = new System.Drawing.Point(197, 48);
            this.btn_End.Name = "btn_End";
            this.btn_End.Size = new System.Drawing.Size(76, 20);
            this.btn_End.TabIndex = 4;
            this.btn_End.Text = "終了";
            this.btn_End.UseVisualStyleBackColor = true;
            this.btn_End.Click += new System.EventHandler(this.btn_End_Click);
            // 
            // timer1
            // 
            this.timer1.Interval = 1000;
            this.timer1.Tick += new System.EventHandler(this.timer1_Tick);
            // 
            // MainFrm
            // 
            this.AutoScaleDimensions = new System.Drawing.SizeF(6F, 12F);
            this.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font;
            this.ClientSize = new System.Drawing.Size(371, 147);
            this.Controls.Add(this.btn_End);
            this.Controls.Add(this.btn_Start);
            this.Controls.Add(this.txt_Data);
            this.Controls.Add(this.btn_PortNo);
            this.Controls.Add(this.txt_PortNo);
            this.Name = "MainFrm";
            this.Text = "MainFrm";
            this.Load += new System.EventHandler(this.MainFrm_Load);
            this.ResumeLayout(false);
            this.PerformLayout();

        }

        #endregion

        private System.Windows.Forms.TextBox txt_PortNo;
        private System.Windows.Forms.Button btn_PortNo;
        private System.Windows.Forms.TextBox txt_Data;
        private System.Windows.Forms.Button btn_Start;
        private System.Windows.Forms.Button btn_End;
        private System.Windows.Forms.Timer timer1;
    }
}

