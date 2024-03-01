Imports System.Drawing.Drawing2D
Imports System.IO
Imports System.Runtime.InteropServices
Imports NPOI.SS.UserModel
Imports OfficeOpenXml
Imports Excel = Microsoft.Office.Interop.Excel


Public Class Çözüm
    Dim selectedFileName As String = String.Empty
    Private selectedFiles As New List(Of String)
    Private selectedFile As String ' Seçilen dosyanın yolu'
    Private selectedFile2 As String
    Public Property ButtonNumber As Integer

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Dim originalImage As Image = My.Resources.logo
        Dim roundImage As New Bitmap(originalImage.Width, originalImage.Height)

        If ButtonNumber = 1 Then
            selectedFile2 = "C:\Users\ardaa\Downloads\ST.23065.660.07.xlsx"
            RichTextBox2.Text = System.IO.Path.GetFileName(selectedFile2)
        End If

        If ButtonNumber = 2 Then
            selectedFile2 = "C:\Users\ardaa\Downloads\FT.23135.660.07.L.xlsx"
            RichTextBox2.Text = System.IO.Path.GetFileName(selectedFile2)
        End If

        If ButtonNumber = 3 Then
            selectedFile2 = "C:\Users\ardaa\Downloads\MST ÖRNEK RAPOR.xlsx"
            RichTextBox2.Text = System.IO.Path.GetFileName(selectedFile2)
        End If

        If ButtonNumber = 4 Then
            selectedFile2 = "C:\Users\ardaa\Downloads\DDT ÖRNEK RAPOR.xlsx"
            RichTextBox2.Text = System.IO.Path.GetFileName(selectedFile2)
        End If


        Using g As Graphics = Graphics.FromImage(roundImage)
            Dim path As New GraphicsPath()
            path.AddEllipse(11, 11, 73, 73)
            g.SetClip(path)
            g.DrawImage(originalImage, 0, 0)
        End Using

        PictureBox1.Image = roundImage
        PictureBox1.BackColor = Color.Transparent ' PictureBox'ın arka planını şeffaf yapar

        Dim deleteButtonImage As Image = New Bitmap(20, 20)
        Using g As Graphics = Graphics.FromImage(deleteButtonImage)
            Using p As New Pen(Color.Red, 2)
                g.DrawLine(p, 5, 5, 15, 15)
                g.DrawLine(p, 15, 5, 5, 15)
            End Using
        End Using
        Button3.Image = deleteButtonImage

        Dim deleteButtonImage2 As Image = New Bitmap(20, 20)
        Using g As Graphics = Graphics.FromImage(deleteButtonImage2)
            Using p As New Pen(Color.Red, 2)
                g.DrawLine(p, 5, 5, 15, 15)
                g.DrawLine(p, 15, 5, 5, 15)
            End Using
        End Using
        Button5.Image = deleteButtonImage2

        RichTextBox1.ReadOnly = True
        RichTextBox1.Enabled = False

        RichTextBox2.ReadOnly = True
        RichTextBox2.Enabled = False

        Label1.Text = "KAYNAK DOSYA"
        Label1.Font = New Font("Arial", 14, FontStyle.Bold)
        Label1.ForeColor = Color.DarkCyan
        Label1.BackColor = Color.White

        Label3.Text = "HEDEF DOSYA"
        Label3.Font = New Font("Arial", 14, FontStyle.Bold)
        Label3.ForeColor = Color.DarkCyan
        Label3.BackColor = Color.White


        Label2.Text = "TİRSAN KARDAN"
        ' Font ayarları
        Label2.Font = New Font("Amaranth", 16, FontStyle.Bold)
        Label2.ForeColor = Color.White
        Label2.BackColor = Color.Transparent

        Button1.Text = "Gözat..."
        Button1.Font = New Font("Calibri", 11)
        Button1.BackColor = Color.White
        Button1.ForeColor = Color.Black

        Button4.Text = "Gözat..."
        Button4.Font = New Font("Calibri", 11)
        Button4.BackColor = Color.White
        Button4.ForeColor = Color.Black

        Button2.Text = "HEDEF DOSYAYA YANSIT"
        Button2.Font = New Font("Calibri", 17)
        Button2.BackColor = Color.GreenYellow
        Button2.ForeColor = Color.Black

        Me.BackgroundImage = My.Resources.background
        Me.BackgroundImageLayout = ImageLayout.Stretch
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim openFileDialog As New OpenFileDialog()

        openFileDialog.Filter = "Excel Dosyaları (*.xls;*.xlsx)|*.xls;*.xlsx|Tüm Dosyalar (*.*)|*.*"
        openFileDialog.FilterIndex = 1
        openFileDialog.RestoreDirectory = True

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            selectedFile = openFileDialog.FileName

            RichTextBox1.Text = System.IO.Path.GetFileName(selectedFile)
            ' Seçilen dosyanın adını saklayın.
        End If
    End Sub

    Private Sub Button4_Click(sender As Object, e As EventArgs) Handles Button4.Click
        Dim openFileDialog As New OpenFileDialog()

        openFileDialog.Filter = "Excel Dosyaları (*.xls;*.xlsx)|*.xls;*.xlsx|Tüm Dosyalar (*.*)|*.*"
        openFileDialog.FilterIndex = 1
        openFileDialog.RestoreDirectory = True

        If openFileDialog.ShowDialog() = DialogResult.OK Then
            If selectedFile <> openFileDialog.FileName Then
                ' Eğer yeni dosya seçildiyse, selectedFile2'yi güncelleyin.
                selectedFile2 = openFileDialog.FileName
                RichTextBox2.Text = System.IO.Path.GetFileName(selectedFile2)
            Else
                ' Eğer aynı dosya tekrar seçildiyse, aynı dosyayı tekrar seçildi olarak işaretleyebilirsiniz.
                RichTextBox2.Text = "Aynı dosya tekrar seçildi"
            End If
        End If
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        If Not String.IsNullOrEmpty(selectedFile) Then
            selectedFiles.Remove(selectedFile)
            ' Seçili dosyanın yolunu temizleyin.
            selectedFile = ""

            RichTextBox1.Clear() ' RichTextBox içeriğini temizleyin.
        End If
    End Sub

    Private Sub Button5_Click(sender As Object, e As EventArgs) Handles Button5.Click
        If Not String.IsNullOrEmpty(selectedFile2) Then
            selectedFiles.Remove(selectedFile2)
            ' Seçili dosyanın yolunu temizleyin.
            selectedFile2 = ""

            RichTextBox2.Clear() ' RichTextBox içeriğini temizleyin.
        End If
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        ' Önce hedef dosya yolunu RichTextBox2'den alın
        ExcelPackage.LicenseContext = LicenseContext.NonCommercial
        Dim newFilePath As String = Path.Combine("C:\Users\ardaa\Documents", RichTextBox2.Text)

        ' Kaynak dosyanın yolu
        Dim sourceFilePath As String = selectedFile ' Daha önce gözattan seçilmiş dosya yolu
        Dim targetFilePath As String = selectedFile2 ' Daha önce gözattan seçilmiş dosya yolu

        If Not String.IsNullOrEmpty(sourceFilePath) Then
            If File.Exists(newFilePath) Then
                ' Eğer hedef dosya varsa, onu açın ve güncelleyin
                Using existingPackage As New ExcelPackage(New FileInfo(newFilePath))
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                    Using sourcePackage As New ExcelPackage(New FileInfo(sourceFilePath))
                        Dim sourceWorksheet As ExcelWorksheet = sourcePackage.Workbook.Worksheets(0)
                        ' Kaynak dosyadan hücre değerini alın
                        Dim valueList As New List(Of Object)()
                        Dim cellvalueList As New List(Of Object)()

                        ' Value'ları listeye ekle
                        If ButtonNumber = 1 Then
                            For i As Integer = 1 To 19
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        If ButtonNumber = 2 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        If ButtonNumber = 3 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        If ButtonNumber = 4 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        ' Hedef dosyaya değerleri aktarın
                        Dim targetWorksheet As ExcelWorksheet = existingPackage.Workbook.Worksheets(0)
                        If ButtonNumber = 1 Then
                            For i As Integer = 0 To 18
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If

                        If ButtonNumber = 2 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If

                        If ButtonNumber = 3 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If

                        If ButtonNumber = 4 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If
                    End Using

                    ' Dosyayı kaydedin
                    existingPackage.Save()
                End Using
            Else
                ' Eğer hedef dosya yoksa, yeni bir dosya oluşturun ve kaynak dosyanın içeriğini kopyalayın
                Using targetPackage As New ExcelPackage(New FileInfo(targetFilePath))
                    Dim newWorksheet As ExcelWorksheet = targetPackage.Workbook.Worksheets.Add("Test 1")
                    ExcelPackage.LicenseContext = LicenseContext.NonCommercial
                    Using sourcePackage As New ExcelPackage(New FileInfo(sourceFilePath))
                        Dim sourceWorksheet As ExcelWorksheet = sourcePackage.Workbook.Worksheets(0)
                        ' Kaynak dosyadan hücre değerini alın
                        Dim valueList As New List(Of Object)()
                        Dim cellvalueList As New List(Of Object)()

                        ' Value'ları listeye ekle
                        If ButtonNumber = 1 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        If ButtonNumber = 2 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        If ButtonNumber = 3 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        If ButtonNumber = 4 Then
                            For i As Integer = 1 To 18
                                valueList.Add(sourceWorksheet.Cells("C" & i).Value)
                                cellvalueList.Add(sourceWorksheet.Cells("A" & i).Value)
                            Next
                        End If

                        ' Hedef dosyaya değerleri aktarın
                        Dim targetWorksheet As ExcelWorksheet = targetPackage.Workbook.Worksheets(0)

                        Dim rowCount As Integer = targetWorksheet.Dimension.Rows
                        Dim colCount As Integer = targetWorksheet.Dimension.Columns

                        ' Tüm hücreleri kopyalayın
                        For row As Integer = 1 To rowCount
                            For col As Integer = 1 To colCount
                                newWorksheet.Cells(row, col).Value = targetWorksheet.Cells(row, col).Value
                            Next
                        Next

                        If ButtonNumber = 1 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If

                        If ButtonNumber = 2 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If

                        If ButtonNumber = 3 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If

                        If ButtonNumber = 4 Then
                            For i As Integer = 0 To 17
                                targetWorksheet.Cells(cellvalueList(i)).Value = valueList(i)
                            Next
                        End If
                    End Using
                    targetPackage.SaveAs(New FileInfo(newFilePath))
                End Using
            End If
            MessageBox.Show("Değer kopyalandı ve Excel dosyası güncellendi veya oluşturuldu: " & newFilePath)
        End If
    End Sub

End Class
