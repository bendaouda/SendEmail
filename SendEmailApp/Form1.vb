Imports System.IO
Imports System.Net.Mail
Imports ExcelDataReader

Imports Excel = Microsoft.Office.Interop.Excel
Imports Office = Microsoft.Office.Core



Public Class Form1
    Public tables As DataTableCollection
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        ' Set the caption bar text of the form.   
        Me.Text = "send Multiple mail"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click


        Try
            Dim Smtp_Server As New SmtpClient
            Dim e_mail As New MailMessage()
            Smtp_Server.UseDefaultCredentials = False
            Smtp_Server.Credentials = New Net.NetworkCredential("27f9b23c2ba6ffa31ff3ba4e4816e927", "a289d3916eee28a2d54a9150d897cace")
            Smtp_Server.Port = 587 '25 ou 587 ou encore comme askia 465
            Smtp_Server.EnableSsl = True
            Smtp_Server.Host = "in-v3.mailjet.com"

            e_mail = New MailMessage()
            ' e_mail.From = New MailAddress(txtFrom.Text)
            e_mail.From = New MailAddress("Enquête sur le devenir des candidats au titre professionnel <ministere.travail@survey-lvdc.fr>")
            e_mail.To.Add(txtTo.Text)
            e_mail.Subject = "Send mail test"
            e_mail.IsBodyHtml = True
            e_mail.Body = txtMessage.Text
            Smtp_Server.Send(e_mail)
            MsgBox("Mail envoyé avec succès")

        Catch error_t As Exception
            MsgBox(error_t.ToString)
        End Try
    End Sub

    Private Sub getMails()


        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel 97-2003 Workbook|*.xlsx|*.xlsx|Excel Workbook"}
            If ofd.ShowDialog() = DialogResult.OK Then
                txtBA.Text = ofd.FileName
                txtMessage.Text = ""
                Using stream = File.Open(ofd.FileName, FileMode.Open, FileAccess.Read)
                    Using reader As IExcelDataReader = ExcelReaderFactory.CreateReader(stream)
                        Dim result As DataSet = reader.AsDataSet(New ExcelDataSetConfiguration() With {
                                                                 .ConfigureDataTable = Function(__) New ExcelDataTableConfiguration() With {
                                                                 .UseHeaderRow = True}})
                        tables = result.Tables
                        cboSheet.Items.Clear()
                        For Each table As DataTable In tables
                            cboSheet.Items.Add(table.TableName)
                        Next

                    End Using
                End Using
            End If
        End Using
    End Sub

    Private Sub DataGridView1_CellContentClick(sender As Object, e As DataGridViewCellEventArgs) Handles dataGr.CellContentClick

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        getMails()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSheet.SelectedIndexChanged
        Dim dt As DataTable = tables(cboSheet.SelectedItem.ToString())
        dataGr.DataSource = dt

        txtTo.Text = dt.Rows(0).Item(1) 'l'idée sera davoir les mail un par un - - avec une boucle 

        Dim line As String
        Dim fileReader As System.IO.StreamReader
        fileReader = My.Computer.FileSystem.OpenTextFileReader("C:\Users\BenD\Desktop\3078_DECPROV2.html") 'récupérer après tous les mails dans le dossier
        line = fileReader.ReadLine
        While (line IsNot Nothing)
            line = fileReader.ReadLine
            txtMessage.Text = txtMessage.Text + line
        End While

        'MsgBox(dt.Columns.Count)

        For j = 2 To dt.Columns.Count - 1
            Dim crochetGauche As Integer = txtMessage.Text.IndexOf("[")
            Dim crochetDroite As Integer = txtMessage.Text.IndexOf("]")

            Dim var1 = Mid(txtMessage.Text, crochetGauche + 1, crochetDroite - crochetGauche + 1)
            txtMessage.Text = txtMessage.Text.Replace(var1, dt.Rows(0).Item(j))

        Next


    End Sub


End Class
