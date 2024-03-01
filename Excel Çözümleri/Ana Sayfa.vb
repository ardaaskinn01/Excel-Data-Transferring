Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.InteropServices
Imports NPOI.SS.UserModel
Imports OfficeOpenXml
Imports Excel = Microsoft.Office.Interop.Excel
Public Class Form2
    Private Sub Form2_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim originalImage As Image = My.Resources.logo
        Dim roundImage As New Bitmap(originalImage.Width, originalImage.Height)

        Using g As Graphics = Graphics.FromImage(roundImage)
            Dim path As New GraphicsPath()
            path.AddEllipse(11, 11, 73, 73)
            g.SetClip(path)
            g.DrawImage(originalImage, 0, 0)
        End Using

        PictureBox1.Image = roundImage
        PictureBox1.BackColor = Color.Transparent ' PictureBox'ın arka planını şeffaf yapar

        Label2.Text = "TİRSAN KARDAN"
        ' Font ayarları
        Label2.Font = New Font("Amaranth", 16, FontStyle.Bold)
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Transparent

        Button1.Text = "STATİK BURULMA TESTİ"
        Button1.Font = New Font("Calibri", 16)
        Button1.BackColor = Color.White
        Button1.ForeColor = Color.Black

        Button4.Text = "DİNAMİK DAYANIKLILIK TESTİ"
        Button4.Font = New Font("Calibri", 16)
        Button4.BackColor = Color.White
        Button4.ForeColor = Color.Black

        Button2.Text = "YORULMA TESTİ"
        Button2.Font = New Font("Calibri", 16)
        Button2.BackColor = Color.White
        Button2.ForeColor = Color.Black

        Button3.Text = "ÇAMUR TESTİ"
        Button3.Font = New Font("Calibri", 16)
        Button3.BackColor = Color.White
        Button3.ForeColor = Color.Black

        Me.BackgroundImage = My.Resources.background
        Me.BackgroundImageLayout = ImageLayout.Stretch
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        OpenForm1(1) ' İlk butona tıklanmış
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        OpenForm1(2) ' İlk butona tıklanmış
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        OpenForm1(3) ' İlk butona tıklanmış
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        OpenForm1(4) ' İlk butona tıklanmış
    End Sub

    Private Sub OpenForm1(buttonNumber As Integer)
        Dim form1 As New Çözüm()
        form1.ButtonNumber = buttonNumber ' Form1'e tıklanan butonun numarasını aktarın
        form1.Show()
    End Sub

End Class