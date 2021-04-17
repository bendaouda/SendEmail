<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.Label1 = New System.Windows.Forms.Label()
        Me.Label2 = New System.Windows.Forms.Label()
        Me.Label3 = New System.Windows.Forms.Label()
        Me.txtBA = New System.Windows.Forms.TextBox()
        Me.txtTo = New System.Windows.Forms.TextBox()
        Me.txtMessage = New System.Windows.Forms.TextBox()
        Me.Button1 = New System.Windows.Forms.Button()
        Me.dataGr = New System.Windows.Forms.DataGridView()
        Me.Button2 = New System.Windows.Forms.Button()
        Me.cboSheet = New System.Windows.Forms.ComboBox()
        Me.Label4 = New System.Windows.Forms.Label()
        Me.TextBox1 = New System.Windows.Forms.TextBox()
        Me.txtFrom = New System.Windows.Forms.Label()
        CType(Me.dataGr, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'Label1
        '
        Me.Label1.AutoSize = True
        Me.Label1.Location = New System.Drawing.Point(298, 49)
        Me.Label1.Name = "Label1"
        Me.Label1.Size = New System.Drawing.Size(66, 13)
        Me.Label1.TabIndex = 0
        Me.Label1.Text = "Choisir la BA"
        '
        'Label2
        '
        Me.Label2.AutoSize = True
        Me.Label2.Location = New System.Drawing.Point(367, 346)
        Me.Label2.Name = "Label2"
        Me.Label2.Size = New System.Drawing.Size(20, 13)
        Me.Label2.TabIndex = 1
        Me.Label2.Text = "To"
        '
        'Label3
        '
        Me.Label3.AutoSize = True
        Me.Label3.Location = New System.Drawing.Point(251, 474)
        Me.Label3.Name = "Label3"
        Me.Label3.Size = New System.Drawing.Size(44, 13)
        Me.Label3.TabIndex = 2
        Me.Label3.Text = "Messge"
        '
        'txtBA
        '
        Me.txtBA.Location = New System.Drawing.Point(370, 46)
        Me.txtBA.Name = "txtBA"
        Me.txtBA.Size = New System.Drawing.Size(539, 20)
        Me.txtBA.TabIndex = 3
        '
        'txtTo
        '
        Me.txtTo.Location = New System.Drawing.Point(420, 338)
        Me.txtTo.Name = "txtTo"
        Me.txtTo.Size = New System.Drawing.Size(376, 20)
        Me.txtTo.TabIndex = 4
        '
        'txtMessage
        '
        Me.txtMessage.Location = New System.Drawing.Point(301, 376)
        Me.txtMessage.Multiline = True
        Me.txtMessage.Name = "txtMessage"
        Me.txtMessage.Size = New System.Drawing.Size(628, 227)
        Me.txtMessage.TabIndex = 5
        '
        'Button1
        '
        Me.Button1.Location = New System.Drawing.Point(530, 609)
        Me.Button1.Name = "Button1"
        Me.Button1.Size = New System.Drawing.Size(75, 23)
        Me.Button1.TabIndex = 6
        Me.Button1.Text = "Send"
        Me.Button1.UseVisualStyleBackColor = True
        '
        'dataGr
        '
        Me.dataGr.AllowUserToOrderColumns = True
        Me.dataGr.ColumnHeadersHeightSizeMode = System.Windows.Forms.DataGridViewColumnHeadersHeightSizeMode.AutoSize
        Me.dataGr.Cursor = System.Windows.Forms.Cursors.Default
        Me.dataGr.Location = New System.Drawing.Point(304, 153)
        Me.dataGr.Name = "dataGr"
        Me.dataGr.Size = New System.Drawing.Size(625, 126)
        Me.dataGr.TabIndex = 7
        '
        'Button2
        '
        Me.Button2.Location = New System.Drawing.Point(924, 46)
        Me.Button2.Name = "Button2"
        Me.Button2.Size = New System.Drawing.Size(37, 20)
        Me.Button2.TabIndex = 8
        Me.Button2.Text = "..."
        Me.Button2.UseVisualStyleBackColor = True
        '
        'cboSheet
        '
        Me.cboSheet.FormattingEnabled = True
        Me.cboSheet.Location = New System.Drawing.Point(370, 107)
        Me.cboSheet.Name = "cboSheet"
        Me.cboSheet.Size = New System.Drawing.Size(183, 21)
        Me.cboSheet.TabIndex = 9
        '
        'Label4
        '
        Me.Label4.AutoSize = True
        Me.Label4.Location = New System.Drawing.Point(301, 114)
        Me.Label4.Name = "Label4"
        Me.Label4.Size = New System.Drawing.Size(37, 13)
        Me.Label4.TabIndex = 10
        Me.Label4.Text = "Feuille"
        '
        'TextBox1
        '
        Me.TextBox1.Location = New System.Drawing.Point(420, 298)
        Me.TextBox1.Name = "TextBox1"
        Me.TextBox1.Size = New System.Drawing.Size(376, 20)
        Me.TextBox1.TabIndex = 11
        '
        'txtFrom
        '
        Me.txtFrom.AutoSize = True
        Me.txtFrom.Location = New System.Drawing.Point(367, 298)
        Me.txtFrom.Name = "txtFrom"
        Me.txtFrom.Size = New System.Drawing.Size(30, 13)
        Me.txtFrom.TabIndex = 12
        Me.txtFrom.Text = "From"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.AutoSizeMode = System.Windows.Forms.AutoSizeMode.GrowAndShrink
        Me.ClientSize = New System.Drawing.Size(1202, 644)
        Me.Controls.Add(Me.txtFrom)
        Me.Controls.Add(Me.TextBox1)
        Me.Controls.Add(Me.Label4)
        Me.Controls.Add(Me.cboSheet)
        Me.Controls.Add(Me.Button2)
        Me.Controls.Add(Me.dataGr)
        Me.Controls.Add(Me.Button1)
        Me.Controls.Add(Me.txtMessage)
        Me.Controls.Add(Me.txtTo)
        Me.Controls.Add(Me.txtBA)
        Me.Controls.Add(Me.Label3)
        Me.Controls.Add(Me.Label2)
        Me.Controls.Add(Me.Label1)
        Me.Name = "Form1"
        Me.StartPosition = System.Windows.Forms.FormStartPosition.CenterScreen
        Me.Text = "Form1"
        CType(Me.dataGr, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents Label1 As Label
    Friend WithEvents Label2 As Label
    Friend WithEvents Label3 As Label
    Friend WithEvents txtBA As TextBox
    Friend WithEvents txtTo As TextBox
    Friend WithEvents txtMessage As TextBox
    Friend WithEvents Button1 As Button
    Friend WithEvents dataGr As DataGridView
    Friend WithEvents Button2 As Button
    Friend WithEvents cboSheet As ComboBox
    Friend WithEvents Label4 As Label
    Friend WithEvents TextBox1 As TextBox
    Friend WithEvents txtFrom As Label
End Class
