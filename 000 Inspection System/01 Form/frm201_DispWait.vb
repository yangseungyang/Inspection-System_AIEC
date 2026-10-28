Option Strict On : Option Explicit On

'GetProcess
Imports System.Diagnostics.Process
'Doevents
Imports System.Windows.Forms.Application

Public Class frm201_DispWait

    'Local ================================================
    Dim mAppPath As String = Mid(Application.StartupPath, 1, (InStr(Application.StartupPath, Application.ProductName) - 1) + Len(Application.ProductName)) & "\"

    Private Sub frm201_DispWait_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        '==============================================================================================================
        'Function : Display start
        'Remarks  :
        '==============================================================================================================

        With Me
            .Opacity = 0
            .Text = "Disp Wait"
            .FormBorderStyle = Windows.Forms.FormBorderStyle.None
            '.TopMost = True

            .Width = 150 : .Height = 100
            .Top = (Screen.PrimaryScreen.Bounds.Height \ 2) - (.Height \ 2)
            .Left = (Screen.PrimaryScreen.Bounds.Width \ 2) - (.Width \ 2)

            .BackgroundImage = System.Drawing.Image.FromFile(mAppPath & "900 Picture\Picture\Common\DispWait.bmp")
            .BackgroundImageLayout = ImageLayout.Stretch

            DoEvents()

        End With

        Me.Opacity = 1

    End Sub

End Class