<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class Form1
    Inherits System.Windows.Forms.Form

    'Descartar substituições de formulário para limpar a lista de componentes.
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

    'Exigido pelo Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'OBSERVAÇÃO: o procedimento a seguir é exigido pelo Windows Form Designer
    'Pode ser modificado usando o Windows Form Designer.  
    'Não o modifique usando o editor de códigos.
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.BackgroundWorker1 = New System.ComponentModel.BackgroundWorker()
        Me.FolderBrowserDialog1 = New System.Windows.Forms.FolderBrowserDialog()
        Me.txtCaminho = New System.Windows.Forms.TextBox()
        Me.btnProcurar = New System.Windows.Forms.Button()
        Me.btnExecutar = New System.Windows.Forms.Button()
        Me.lblArquivo = New System.Windows.Forms.Label()
        Me.SuspendLayout()
        '
        'txtCaminho
        '
        Me.txtCaminho.Location = New System.Drawing.Point(51, 43)
        Me.txtCaminho.Name = "txtCaminho"
        Me.txtCaminho.Size = New System.Drawing.Size(253, 20)
        Me.txtCaminho.TabIndex = 0
        '
        'btnProcurar
        '
        Me.btnProcurar.Location = New System.Drawing.Point(310, 41)
        Me.btnProcurar.Name = "btnProcurar"
        Me.btnProcurar.Size = New System.Drawing.Size(75, 23)
        Me.btnProcurar.TabIndex = 1
        Me.btnProcurar.Text = "..."
        Me.btnProcurar.UseVisualStyleBackColor = True
        '
        'btnExecutar
        '
        Me.btnExecutar.Location = New System.Drawing.Point(169, 83)
        Me.btnExecutar.Name = "btnExecutar"
        Me.btnExecutar.Size = New System.Drawing.Size(75, 23)
        Me.btnExecutar.TabIndex = 2
        Me.btnExecutar.Text = "Executar"
        Me.btnExecutar.UseVisualStyleBackColor = True
        '
        'lblArquivo
        '
        Me.lblArquivo.AutoSize = True
        Me.lblArquivo.Location = New System.Drawing.Point(12, 126)
        Me.lblArquivo.Name = "lblArquivo"
        Me.lblArquivo.Size = New System.Drawing.Size(43, 13)
        Me.lblArquivo.TabIndex = 3
        Me.lblArquivo.Text = "Arquivo"
        '
        'Form1
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(431, 157)
        Me.Controls.Add(Me.lblArquivo)
        Me.Controls.Add(Me.btnExecutar)
        Me.Controls.Add(Me.btnProcurar)
        Me.Controls.Add(Me.txtCaminho)
        Me.Name = "Form1"
        Me.Text = "RenomearExcel"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub

    Friend WithEvents BackgroundWorker1 As System.ComponentModel.BackgroundWorker
    Friend WithEvents FolderBrowserDialog1 As FolderBrowserDialog
    Friend WithEvents txtCaminho As TextBox
    Friend WithEvents btnProcurar As Button
    Friend WithEvents btnExecutar As Button
    Friend WithEvents lblArquivo As Label
End Class
