<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frm030_InspectionInformation
    Inherits System.Windows.Forms.Form

    'フォームがコンポーネントの一覧をクリーンアップするために dispose をオーバーライドします。
    <System.Diagnostics.DebuggerNonUserCode()> _
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Windows フォーム デザイナで必要です。
    Private components As System.ComponentModel.IContainer

    'メモ: 以下のプロシージャは Windows フォーム デザイナで必要です。
    'Windows フォーム デザイナを使用して変更できます。  
    'コード エディタを使って変更しないでください。
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container
        Me.Panel2 = New System.Windows.Forms.Panel
        Me.btnNext = New System.Windows.Forms.Button
        Me.btnPrev = New System.Windows.Forms.Button
        Me.btnStop = New System.Windows.Forms.Button
        Me.btnEnd = New System.Windows.Forms.Button
        Me.btnStart = New System.Windows.Forms.Button
        Me.Panel1 = New System.Windows.Forms.Panel
        Me.Label3 = New System.Windows.Forms.Label
        Me.tmInspectionTime = New ToY_StopWatch.ToY_StopWatch
        Me.Label2 = New System.Windows.Forms.Label
        Me.Label1 = New System.Windows.Forms.Label
        Me.tmCount = New ToY_StopWatch.ToY_StopWatch
        Me.tmWaite = New ToY_StopWatch.ToY_StopWatch
        Me.SerialPort1 = New System.IO.Ports.SerialPort(Me.components)
        Me.Panel2.SuspendLayout()
        Me.Panel1.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel2
        '
        Me.Panel2.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor1
        Me.Panel2.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel2.Controls.Add(Me.btnNext)
        Me.Panel2.Controls.Add(Me.btnPrev)
        Me.Panel2.Controls.Add(Me.btnStop)
        Me.Panel2.Controls.Add(Me.btnEnd)
        Me.Panel2.Controls.Add(Me.btnStart)
        Me.Panel2.Location = New System.Drawing.Point(12, 213)
        Me.Panel2.Name = "Panel2"
        Me.Panel2.Size = New System.Drawing.Size(433, 63)
        Me.Panel2.TabIndex = 4
        '
        'btnNext
        '
        Me.btnNext.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor
        Me.btnNext.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnNext.Location = New System.Drawing.Point(147, 3)
        Me.btnNext.Name = "btnNext"
        Me.btnNext.Size = New System.Drawing.Size(30, 30)
        Me.btnNext.TabIndex = 15
        Me.btnNext.UseVisualStyleBackColor = True
        '
        'btnPrev
        '
        Me.btnPrev.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor
        Me.btnPrev.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnPrev.Location = New System.Drawing.Point(111, 3)
        Me.btnPrev.Name = "btnPrev"
        Me.btnPrev.Size = New System.Drawing.Size(30, 30)
        Me.btnPrev.TabIndex = 14
        Me.btnPrev.UseVisualStyleBackColor = True
        '
        'btnStop
        '
        Me.btnStop.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor
        Me.btnStop.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnStop.Location = New System.Drawing.Point(75, 3)
        Me.btnStop.Name = "btnStop"
        Me.btnStop.Size = New System.Drawing.Size(30, 30)
        Me.btnStop.TabIndex = 13
        Me.btnStop.UseVisualStyleBackColor = True
        '
        'btnEnd
        '
        Me.btnEnd.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor
        Me.btnEnd.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnEnd.Location = New System.Drawing.Point(39, 3)
        Me.btnEnd.Name = "btnEnd"
        Me.btnEnd.Size = New System.Drawing.Size(30, 30)
        Me.btnEnd.TabIndex = 12
        Me.btnEnd.UseVisualStyleBackColor = True
        '
        'btnStart
        '
        Me.btnStart.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor
        Me.btnStart.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.btnStart.Location = New System.Drawing.Point(3, 3)
        Me.btnStart.Name = "btnStart"
        Me.btnStart.Size = New System.Drawing.Size(30, 30)
        Me.btnStart.TabIndex = 11
        Me.btnStart.UseVisualStyleBackColor = True
        '
        'Panel1
        '
        Me.Panel1.BackColor = System.Drawing.Color.LightSlateGray
        Me.Panel1.BackgroundImage = Global.Inspection_System.My.Resources.Resources.BackColor1
        Me.Panel1.BackgroundImageLayout = System.Windows.Forms.ImageLayout.Stretch
        Me.Panel1.Controls.Add(Me.Label3)
        Me.Panel1.Controls.Add(Me.tmInspectionTime)
        Me.Panel1.Controls.Add(Me.Label2)
        Me.Panel1.Controls.Add(Me.Label1)
        Me.Panel1.Controls.Add(Me.tmCount)
        Me.Panel1.Controls.Add(Me.tmWaite)
        Me.Panel1.Location = New System.Drawing.Point(12, 161)
        Me.Panel1.Name = "Panel1"
        Me.Panel1.Size = New System.Drawing.Size(433, 46)
        Me.Panel1.TabIndex = 3
        '
        'Label3
        '
        Me.Label3.ForeColor = System.Drawing.Color.White
        Me.Label3.Image = Global.Inspection_System.My.Resources.Resources.BlueLabel
        Me.Label3.Location = New System.Drawing.Point(318, 0)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(79, 20)
        Me.Label3.TabIndex = 12
        Me.Label3.Text = "検査時間"
        Me.Label3.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tmInspectionTime
        '
        Me.tmInspectionTime.BackColor = System.Drawing.Color.Black
        Me.tmInspectionTime.CountUpDown = ToY_StopWatch.ToY_StopWatch.cUpDown.Up
        Me.tmInspectionTime.Font = New System.Drawing.Font("ＭＳ 明朝", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.tmInspectionTime.ForeColor = System.Drawing.Color.White
        Me.tmInspectionTime.Location = New System.Drawing.Point(318, 20)
        Me.tmInspectionTime.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.tmInspectionTime.Name = "tmInspectionTime"
        Me.tmInspectionTime.SettingValue = 0
        Me.tmInspectionTime.Size = New System.Drawing.Size(79, 20)
        Me.tmInspectionTime.TabIndex = 11
        '
        'Label2
        '
        Me.Label2.ForeColor = System.Drawing.Color.White
        Me.Label2.Image = Global.Inspection_System.My.Resources.Resources.BlueLabel
        Me.Label2.Location = New System.Drawing.Point(124, 0)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(77, 20)
        Me.Label2.TabIndex = 7
        Me.Label2.Text = "カウント"
        Me.Label2.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'Label1
        '
        Me.Label1.ForeColor = System.Drawing.Color.White
        Me.Label1.Image = Global.Inspection_System.My.Resources.Resources.BlueLabel
        Me.Label1.Location = New System.Drawing.Point(8, 0)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(81, 20)
        Me.Label1.TabIndex = 6
        Me.Label1.Text = "待ち時間"
        Me.Label1.TextAlign = System.Drawing.ContentAlignment.MiddleCenter
        '
        'tmCount
        '
        Me.tmCount.BackColor = System.Drawing.Color.Black
        Me.tmCount.CountUpDown = ToY_StopWatch.ToY_StopWatch.cUpDown.Up
        Me.tmCount.Font = New System.Drawing.Font("ＭＳ 明朝", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.tmCount.ForeColor = System.Drawing.Color.White
        Me.tmCount.Location = New System.Drawing.Point(124, 20)
        Me.tmCount.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.tmCount.Name = "tmCount"
        Me.tmCount.SettingValue = 0
        Me.tmCount.Size = New System.Drawing.Size(77, 20)
        Me.tmCount.TabIndex = 4
        '
        'tmWaite
        '
        Me.tmWaite.BackColor = System.Drawing.Color.Black
        Me.tmWaite.CountUpDown = ToY_StopWatch.ToY_StopWatch.cUpDown.Up
        Me.tmWaite.Font = New System.Drawing.Font("ＭＳ ゴシック", 12.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(128, Byte))
        Me.tmWaite.ForeColor = System.Drawing.Color.White
        Me.tmWaite.Location = New System.Drawing.Point(10, 20)
        Me.tmWaite.Margin = New System.Windows.Forms.Padding(4, 4, 4, 4)
        Me.tmWaite.Name = "tmWaite"
        Me.tmWaite.SettingValue = 0
        Me.tmWaite.Size = New System.Drawing.Size(79, 20)
        Me.tmWaite.TabIndex = 3
        '
        'frm030_InspectionInformation
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 12.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(471, 332)
        Me.Controls.Add(Me.Panel2)
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frm030_InspectionInformation"
        Me.Text = "frm030_InspectionInformation"
        Me.Panel2.ResumeLayout(False)
        Me.Panel1.ResumeLayout(False)
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As System.Windows.Forms.Panel
    Friend WithEvents Label2 As System.Windows.Forms.Label
    Friend WithEvents Label1 As System.Windows.Forms.Label
    Friend WithEvents tmCount As ToY_StopWatch.ToY_StopWatch
    Friend WithEvents tmWaite As ToY_StopWatch.ToY_StopWatch
    Friend WithEvents Panel2 As System.Windows.Forms.Panel
    Friend WithEvents btnNext As System.Windows.Forms.Button
    Friend WithEvents btnPrev As System.Windows.Forms.Button
    Friend WithEvents btnStop As System.Windows.Forms.Button
    Friend WithEvents btnEnd As System.Windows.Forms.Button
    Friend WithEvents btnStart As System.Windows.Forms.Button
    Friend WithEvents tmInspectionTime As ToY_StopWatch.ToY_StopWatch
    Friend WithEvents Label3 As System.Windows.Forms.Label
    Friend WithEvents SerialPort1 As System.IO.Ports.SerialPort
End Class
