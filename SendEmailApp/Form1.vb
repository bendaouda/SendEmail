Imports System.IO
Imports System.Net.Mail
Imports System.Runtime.CompilerServices
Imports ExcelDataReader
Public Class Form1
    Public tables As DataTableCollection
    Public pathname As String()
    Public nombre_ligne As Double
    Public dt As DataTable
    Public chemin As String

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Me.Text = "Multiple Email - LVDC"
    End Sub

    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim line As String
        Dim fileReader As System.IO.StreamReader
        Dim FromEmail As String = Nothing
        Dim SubjectLib As String = Nothing
        Dim EmailText As String = Nothing
        Dim ToEmail As String = Nothing

        For k = 0 To nombre_ligne - 1
            ToEmail = ""
            FromEmail = ""
            SubjectLib = ""
            EmailText = ""

            ToEmail = dt.Rows(k).Item(1) 'l'idée sera davoir les mail un par un - - avec une boucle 
            FromEmail = dt.Rows(k).Item(4) 'expéditeur
            SubjectLib = dt.Rows(k).Item(3) 'On récupere l'objet

            fileReader = Nothing

            fileReader = My.Computer.FileSystem.OpenTextFileReader(chemin + dt.Rows(k).Item(2)) 'récupérer après tous les mails dans le dossier. Le row change
            line = fileReader.ReadLine
            While (line IsNot Nothing)
                line = fileReader.ReadLine
                EmailText = EmailText + line
            End While


            'On renseigne les infos fichier et dans l'odre d'affichage dans le mail
            For j = 5 To dt.Columns.Count - 1
                Dim crochetGauche As Integer = EmailText.IndexOf("[")
                Dim crochetDroite As Integer = EmailText.IndexOf("]")

                Dim var1 = Mid(EmailText, crochetGauche + 1, crochetDroite - crochetGauche + 1)
                EmailText = EmailText.Replace(var1, dt.Rows(k).Item(j))
            Next



            Try
                Dim Smtp_Server As New SmtpClient
                Dim e_mail As New MailMessage()
                Smtp_Server.UseDefaultCredentials = False
                Smtp_Server.Credentials = New Net.NetworkCredential("27f9b23c2ba6ffa31ff3ba4e4816e927", "a289d3916eee28a2d54a9150d897cace")
                Smtp_Server.Port = 587 '25 ou 587 ou encore comme askia 465
                Smtp_Server.EnableSsl = True
                Smtp_Server.Host = "in-v3.mailjet.com"

                e_mail = New MailMessage()
                e_mail.From = New MailAddress(SubjectLib + " <" + FromEmail + ">")
                e_mail.To.Add(ToEmail)
                e_mail.Subject = SubjectLib
                e_mail.IsBodyHtml = True
                e_mail.Body = EmailText
                Smtp_Server.Send(e_mail)
                MsgBox("Mail envoyé avec succès")

            Catch error_t As Exception
                MsgBox(error_t.ToString)
            End Try
        Next



    End Sub

    Private Sub getMails()
        Using ofd As OpenFileDialog = New OpenFileDialog() With {.Filter = "Excel 97-2003 Workbook|*.xlsx|*.xlsx|Excel Workbook"}
            If ofd.ShowDialog() = DialogResult.OK Then
                txtBA.Text = ofd.FileName
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

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        getMails()
    End Sub

    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cboSheet.SelectedIndexChanged

        pathname = txtBA.Text.Split(New Char() {"\"c})

        chemin = ""
        For i = 0 To pathname.Length - 2
            chemin = chemin + pathname(i) + "\"
        Next

        dt = tables(cboSheet.SelectedItem.ToString())
        dataGr.DataSource = dt

        nombre_ligne = dt.Rows.Count





    End Sub

    Private Sub ToolTip1_Popup(sender As Object, e As PopupEventArgs)
        Me.Text = "Help"
    End Sub

    Private Sub Button3_Click(sender As Object, e As EventArgs) Handles Button3.Click
        Help.Show()

    End Sub
End Class
