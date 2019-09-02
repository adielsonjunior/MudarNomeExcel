Imports System.IO
Imports Microsoft.Office.Interop

Public Class Form1
    Dim dt As DataTable = New DataTable()
    Dim DiretorioSaida = "C:\dados\"
    Dim arrayArquivos As String()
    Private Sub btnProcurar_Click(sender As Object, e As EventArgs) Handles btnProcurar.Click
        If FolderBrowserDialog1.ShowDialog() = DialogResult.OK Then
            txtCaminho.Text = FolderBrowserDialog1.SelectedPath
        End If





        'CRIA DATATABLE

        dt.Columns.AddRange(New DataColumn() {New DataColumn("FileName"), New DataColumn("FilePath"), New DataColumn("Progress"), New DataColumn("UF"), New DataColumn("Categoria"), New DataColumn("Periodo"), New DataColumn("Saida")})

        'CARREGA LISTA DE ARQUIVOS TXT DO DIRETORIO EM UM ARRAY DE STRING
        arrayArquivos = Directory.GetFiles(FolderBrowserDialog1.SelectedPath, "*.xls")






        'CRIA DIRETORIO
        If Directory.Exists(FolderBrowserDialog1.SelectedPath & "\Sinapi_Renomeados") = False Then
            Directory.CreateDirectory(FolderBrowserDialog1.SelectedPath & "\Sinapi_Renomeados")
            DiretorioSaida = FolderBrowserDialog1.SelectedPath & "\Sinapi_Renomeados"
        Else
            DiretorioSaida = FolderBrowserDialog1.SelectedPath & "\Sinapi_Renomeados"
        End If

    End Sub

    Private Sub btnExecutar_Click(sender As Object, e As EventArgs) Handles btnExecutar.Click

        'VERIFICA SE NÃO ESTÁ VAZIO
        If IsNothing(arrayArquivos) = False Then
            If arrayArquivos.Count > 0 Then

                For Each valor As String In arrayArquivos

                    'APENAS PARA PEGAR O NOME DO ARQUIVO ENTRE A STRING
                    Dim arraynomeArquivo As String() = Split(valor, "\")

                    Dim linha As DataRow = dt.NewRow
                    linha("FileName") = arraynomeArquivo(arraynomeArquivo.Count - 1)
                    linha("FilePath") = valor
                    linha("UF") = Split(arraynomeArquivo(arraynomeArquivo.Count - 1), "_")(5)
                    linha("Periodo") = Split(arraynomeArquivo(arraynomeArquivo.Count - 1), "_")(6).Substring(0, 2) & "/" & Split(arraynomeArquivo(arraynomeArquivo.Count - 1), "_")(6).Substring(2, 4)
                    linha("Progress") = "0%"
                    linha("Categoria") = IIf(Split(Split(arraynomeArquivo(arraynomeArquivo.Count - 1), "_")(7), ".")(0) = "NaoDesonerado", "Onerado", "Desonerado")
                    linha("Saida") = DiretorioSaida

                    dt.Rows.Add(linha)
                Next
            End If
        End If
        For Each itens As DataRow In dt.Rows
            RenomearExcel(itens("FilePath"), itens("FileName"))
            lblArquivo.Text = itens("FileName")

        Next
    End Sub


    Public Sub RenomearExcel(caminho As String, nome As String)
        Try

            Dim xlApp As New Excel.Application
            Dim xlWorkBook As Excel.Workbook
            xlWorkBook = xlApp.Workbooks.Open(caminho, ReadOnly:=False)
            Dim xlWorkSheet As Excel.Worksheet = xlWorkBook.Sheets(1)



            Dim teste As String = xlWorkSheet.Name

            If Not xlWorkSheet.Name = "Sheet1" Then
                xlWorkSheet.Name = "Sheet1"
                xlWorkBook.SaveAs(DiretorioSaida & "/" & nome)
                xlWorkBook.Close()

            End If
        Catch ex As Exception

        End Try
    End Sub


End Class