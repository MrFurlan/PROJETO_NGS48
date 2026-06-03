Imports System.Collections
Imports System.IO
Imports System.Text
Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Importa
    Inherits BasePage

    Dim path As String
    Dim Index As Integer
    Dim NomeArq As String
    Dim arq As String
    Dim display As String
    Dim Mensagem As String
    Dim Arquivo As String

    Dim Ds As DataSet
    Dim dr As DataRow

    Dim Sql As String
    Dim SqlArray As New ArrayList

    Dim ConexaoOrigem As String = ""
    Dim ConexaoDestino As String = ""

    Private Sub Page_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Me.setMenu(eModulo.Gestao)
        If Not IsPostBack And IsConnect Then
            '    If Not Funcoes.VerificaPermissao("Importa", "ACESSAR") Then
            'MsgBox(Me.Page, "Usuário sem permisão para acessar essa página.", "~/Gestao.aspx")
            'Exit Sub
            '    End If
        End If
    End Sub

    Private Sub ImportaContasAPagar()

        Dim SqlArray As New ArrayList
        Dim dsaux As New Data.DataSet
        Dim i As Integer = 0

        Sql = ""

        path = "c:\resumo\cpimport.txt"
        'path = Arquivo

        If (File.Exists(path)) Then
            Dim enc As Encoding = Encoding.ASCII
            Dim sr As StreamReader = New StreamReader(path, enc)
            Dim line As String

            Try
                Do
                    line = sr.ReadLine()

                    If line Is Nothing Then
                        Exit Do
                    End If

                    '' Limpar caracteres especiais
                    If line <> Nothing Then
                        Sql &= Funcoes.EliminarCaracteresEspeciais(RTrim(line))
                        i = i + 1
                    End If

                    If i > 50 Then
                        SqlArray.Add(Sql)
                        If Banco.GravaBanco(SqlArray) = False Then
                            Mensagem = HttpContext.Current.Session("ssMessage")
                            txtMensagem.Text = Mensagem
                            ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
                        Else
                            i = 0
                            SqlArray.Clear()
                            Sql = ""
                        End If
                    End If

                Loop Until line Is Nothing

                sr.Close()

                If Banco.GravaBanco(SqlArray) = False Then
                    Mensagem = HttpContext.Current.Session("ssMessage")
                    txtMensagem.Text = Mensagem
                    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)

                Else
                    Mensagem = "Importado com sucesso."
                    txtMensagem.Text = Mensagem
                    ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
                End If
            Catch E As Exception
                sr.Close()
                Mensagem = E.Message
                txtMensagem.Text = Mensagem
                ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
            End Try
        Else
            Mensagem = "Titulos.txt não encontrado."
            txtMensagem.Text = Mensagem
            ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
        End If


        'Dim SqlArray As New ArrayList
        'Dim dsaux As New Data.DataSet
        'Dim i As Integer = 0

        'Sql = ""

        ''Dim strArquivos() As String = Directory.GetFiles("c:\resumo\", "CPIMPOR.txt")
        'Dim strArquivos As String = "c:\resumo\CPIMPOR.txt"
        'path = strArquivos
        ''For Each strArquivo As String In strArquivos

        'Dim enc As Encoding = Encoding.ASCII
        ''Dim sr As StreamReader = New StreamReader(strArquivo, enc)
        'Dim sr As StreamReader = New StreamReader(path)

        ''StreamReader(path)
        'Dim line As String

        'Try
        '    Do
        '        line = sr.ReadLine()

        '        If line Is Nothing Then
        '            Exit Do
        '        End If

        '        If line <> Nothing Then
        '            Sql &= Funcoes.EliminarCaracteresEspeciais(RTrim(line))
        '        End If

        '        SqlArray.Add(Sql)

        '        If Banco.GravaBanco(SqlArray) = False Then
        '            Mensagem = Globais.GMsg
        '            ScriptManager.RegisterClientScriptBlock(Me, Me.TabContainer1.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
        '        End If

        '    Loop Until line Is Nothing

        '    sr.Close()

        'Catch E As Exception
        '    sr.Close()
        '    Mensagem = E.Message
        '    ScriptManager.RegisterClientScriptBlock(Me, Me.TabContainer1.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
        'End Try

        ''Next

        'Mensagem = "Arquivo Importado com Sucesso... "
        'ScriptManager.RegisterClientScriptBlock(Me, Me.TabContainer1.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)

    End Sub

    Private Sub ExportaContasAPagar()
        Dim arquivo As String = "c:\resumo\cpexpor.txt"
        Dim linha As String
        Dim strm As StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)

        Dim Array As New ArrayList

        Sql = "SELECT  Registro_Id, UsuarioLiberacao, UsuarioLiberacaoData" & vbCrLf & _
              " FROM ContasAPagar" & vbCrLf & _
              " WHERE UsuarioLiberacao <> ''" & vbCrLf


        Sql = "SELECT  Registro_Id, ISNULL(UsuarioLiberacao, '') AS UsuarioLiberacao, UsuarioLiberacaoData" & vbCrLf & _
              " FROM   ContasAPagar" & vbCrLf & _
              " WHERE  UsuarioLiberacao <> ''" & vbCrLf

        Ds = Banco.ConsultaDataSet(Sql, "ContasAPagar")

        If Ds.Tables(0).Rows.Count > 0 Then
            For Each row In Ds.Tables(0).Rows
                linha = Format(row("Registro_Id"), "000000") & vbCrLf & _
                Funcoes.AlinharEsquerda(row("UsuarioLiberacao"), 15, " ") & vbCrLf & _
                Format(row("UsuarioLiberacaoData"), "ddMMyyyy") & vbCrLf

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            Next
        End If

        txtMensagem.Text = "Exportado com Sucesso"

    End Sub

    Protected Sub cmdImportar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdImportar.Click
        Try
            ImportaContasAPagar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdExportar_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles cmdExportar.Click
        Try
            ExportaContasAPagar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdImportaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Valida As Boolean

        ConexaoOrigem = HttpContext.Current.Session("ssEnderecoLocal")
        ConexaoDestino = "Data Source=sql_serv; Initial Catalog= Insol2010; User Id=sa; Password=INSOL_!"

        Sql = " SELECT * FROM Clientes " & vbCrLf & _
              " WHERE Nome <> ''"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Valida = False

            If Len(Dr("CLiente_Id")) = 11 Then
                Valida = Funcoes.ValidaCPF(Dr("CLiente_Id"))
            End If
            If Len(Dr("CLiente_Id")) = 14 Then
                Valida = Funcoes.ValidaCNPJ(Dr("CLiente_Id"))
            End If

            If Valida = True Then
                Sql = "DELETE FROM Insol2010..Clientes Where Cliente_Id = '" & Dr("Cliente_Id") & "' And Endereco_Id = " & Dr("Endereco_Id")
                SqlArray.Add(Sql)
                Sql = "DELETE FROM Insol2010..ClientesXTipos Where Tipo_Id = 4 And Cliente_Id = '" & Dr("Cliente_Id") & "' And Endereco_Id = " & Dr("Endereco_Id")
                SqlArray.Add(Sql)
                Sql = "DELETE FROM Insol2010..ClientesXTipos Where Tipo_Id = 8 And Cliente_Id = '" & Dr("Cliente_Id") & "' And Endereco_Id = " & Dr("Endereco_Id")
                SqlArray.Add(Sql)

                Sql = "INSERT INTO Insol2010..Clientes " & vbCrLf & _
                      "(Cliente_Id, Endereco_Id, Regiao, Categoria, Estado, Nome, " & vbCrLf & _
                      "Fantasia, Endereco, Bairro, Cep, Cidade, Inscricao, Telefone, " & vbCrLf & _
                      "Fax, Email, Imagem, Reduzido, Numero, Complemento, CodigoDoMunicipio, Pais) " & vbCrLf & _
                      "VALUES ('" & vbCrLf

                Sql &= Dr("Cliente_Id") & "'," & vbCrLf & _
                       Dr("Endereco_Id") & "," & vbCrLf & _
                       Dr("Regiao") & "," & vbCrLf & _
                       Dr("Categoria") & ", '" & vbCrLf & _
                       Dr("Estado") & "', '" & vbCrLf & _
                       Dr("Nome") & "', '" & vbCrLf & _
                       Dr("Fantasia") & "', '" & vbCrLf & _
                       Dr("Endereco") & "', '" & vbCrLf & _
                       Dr("Bairro") & "', '" & vbCrLf & _
                       Dr("Cep") & "', '" & vbCrLf & _
                       Dr("Cidade") & "', '" & vbCrLf & _
                       Dr("Inscricao") & "', '" & vbCrLf & _
                       Dr("Telefone") & "', '" & vbCrLf & _
                       Dr("Fax") & "', '" & vbCrLf & _
                       Dr("Email") & "', '" & vbCrLf & _
                       Dr("Imagem") & "', '" & vbCrLf & _
                       Dr("Reduzido") & "', " & vbCrLf & _
                       "0, '" & vbCrLf & _
                       Dr("Complemento") & "', " & vbCrLf & _
                       "0, 1058)" & vbCrLf

                SqlArray.Add(Sql)

                Sql = " INSERT INTO Insol2010..ClientesXTipos" & vbCrLf & _
                        " (Tipo_Id" & vbCrLf & _
                        ", Cliente_Id" & vbCrLf & _
                        ", Endereco_Id)" & vbCrLf & _
                        " VALUES(4, '" & Dr("Cliente_Id") & "', " & Dr("Endereco_Id") & ")" & vbCrLf
                SqlArray.Add(Sql)

                Sql = " INSERT INTO Insol2010..ClientesXTipos" & vbCrLf & _
                        " (Tipo_Id" & vbCrLf & _
                        ", Cliente_Id" & vbCrLf & _
                        ", Endereco_Id)" & vbCrLf & _
                        " VALUES(8, '" & Dr("Cliente_Id") & "', " & Dr("Endereco_Id") & ")" & vbCrLf
                SqlArray.Add(Sql)

            End If
        Next
        If Banco.GravaBanco(SqlArray) = False Then
            HttpContext.Current.Session("ssCampo") = "Lancamentos_Credito"
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage").ToString)
        Else
            MsgBox(Me.Page, "Processo concluído com Sucesso.", eTitulo.Sucess)
        End If
    End Sub

    Protected Sub cmdImportaSinco_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim SqlArray As New ArrayList
        Dim dsaux As New Data.DataSet
        Dim i As Integer = 0
        Dim Data As String
        Dim Conta As String
        'Dim Custo As String
        'Dim ContraPartida As String
        Dim Valor As String
        Dim DC As String

        Try
            'Data = Format(CDate("31" & "/" & "07" & "/" & "20" & "06"), "yyyy/MM/dd")

            path = "c:\Sinco2010\razao.txt"

            If (File.Exists(path)) Then
                Dim enc As Encoding = Encoding.ASCII
                Dim sr As StreamReader = New StreamReader(path, enc)
                Dim line As String
                Try
                    Do
                        line = sr.ReadLine()

                        If line Is Nothing Then
                            Exit Do
                        End If

                        Data = Mid(line, 1, 2) & "/" & Mid(line, 3, 2) & "/" & Mid(line, 5, 4)
                        Conta = Mid(line, 9, 20)
                        Valor = Mid(line, 93, 15) & "." & Mid(line, 108, 2)
                        DC = Mid(line, 110, 1)
                        i += 1
                        Sql = "Insert into RazaoSinco (Data, Conta, Ordem, Valor, DC)" & vbCrLf & _
                              " Values ('" & Format(CDate(Data), "yyyy/MM/dd") & "'" & vbCrLf & _
                              ", '" & RTrim(Conta) & "'" & vbCrLf & _
                              ", " & i & vbCrLf & _
                              ", " & Valor & vbCrLf & _
                              ",'" & DC & "')" & vbCrLf
                        SqlArray.Add(Sql)

                    Loop Until line Is Nothing

                    sr.Close()

                    If Banco.GravaBanco(SqlArray) = False Then
                        sr.Close()
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Else
                        HttpContext.Current.Session("ssMessage") = "Arquivo Importado com sucesso."
                        sr.Close()
                        MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    End If
                Catch ex As Exception
                    HttpContext.Current.Session("ssMessage") = ex.Message
                    sr.Close()
                End Try
            Else
                'Return "Cli491.txt não encontrado."
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim arquivo As String = "C:\Sinco2010\Saldo.txt"
        Dim linha As String = ""
        Dim strm As StreamWriter
        Dim Inicial As Decimal
        Dim Saldo As Decimal


        If Dir(arquivo).Length > 0 Then Kill(arquivo)

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
              "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
              "  FROM       RazaoSinco AS RID" & vbCrLf & _
              "  WHERE     (DC = 'D') AND (Data = '2010/01/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
              " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
              "   FROM      RazaoSinco AS RIC" & vbCrLf & _
              "   WHERE     (DC = 'C') AND (Data = '2010/01/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
              " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
              "   FROM       RazaoSinco AS RMD" & vbCrLf & _
              "   WHERE     (DC = 'D') AND (Data BETWEEN '2010/01/02' AND '2010/01/31') AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
              " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
              "   FROM       RazaoSinco AS RMC" & vbCrLf & _
              "   WHERE     (DC = 'C') AND (Data BETWEEN '2010/01/02' AND '2010/01/31') AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
              " FROM RazaoSinco" & vbCrLf & _
              " GROUP BY Conta" & vbCrLf & _
              " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else

                With Dr
                    linha = "01012010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0")
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0")

                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '---Fevereiro-------------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/02/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/02/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 2) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 2) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01022010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf

                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next


        '----Marco----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/03/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/03/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 3) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 3) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01032010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf

                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Abril----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/04/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/04/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 4) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 4) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01042010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Maio----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/05/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/05/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 5) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 5) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01052010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Junho----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/06/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/06/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 6) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 6) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01062010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Julho----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/07/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/07/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 7) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 7) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01072010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Agosto----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/08/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/08/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 8) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 8) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01082010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Setembro----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/09/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/09/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 9) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 9) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01092010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Outubro----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/10/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/10/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 10) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 10) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01102010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Novembro----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/11/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/11/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 11) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 11) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01112010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        '----Dezembro----------------------------------------------------------------

        Sql = " SELECT     Conta, ISNULL" & vbCrLf & _
                  "((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "  FROM       RazaoSinco AS RID" & vbCrLf & _
                  "  WHERE     (DC = 'D') AND (Data < '2010/12/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialDebito, ISNULL" & vbCrLf & _
                  " ((SELECT    SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM      RazaoSinco AS RIC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Data < '2010/12/01') AND (RazaoSinco.Conta = Conta)), 0) AS InicialCredito, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMD" & vbCrLf & _
                  "   WHERE     (DC = 'D') AND (Month(Data) = 12) AND (RazaoSinco.Conta = Conta)), 0) AS Debitos, ISNULL" & vbCrLf & _
                  " ((SELECT     SUM(Valor) AS Inicial" & vbCrLf & _
                  "   FROM       RazaoSinco AS RMC" & vbCrLf & _
                  "   WHERE     (DC = 'C') AND (Month(Data) = 12) AND (RazaoSinco.Conta = Conta)), 0) AS Creditos" & vbCrLf & _
                  " FROM RazaoSinco" & vbCrLf & _
                  " GROUP BY Conta" & vbCrLf & _
                  " ORDER BY Conta" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            If Dr("InicialDebito") = 0 And Dr("InicialCredito") = 0 And Dr("Debitos") = 0 And Dr("Creditos") = 0 Then
            Else
                With Dr
                    linha = "01122010"                                                                                                'Tipo
                    linha &= Funcoes.AlinharEsquerda(.Item("Conta"), 28, " ")
                    Inicial = .Item("InicialDebito") - .Item("InicialCredito")
                    Saldo = Inicial
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                    linha &= Funcoes.AlinharDireita(Replace((.Item("Debitos")), ",", ""), 17, "0") & vbCrLf & _
                    Funcoes.AlinharDireita(Replace((.Item("Creditos")), ",", ""), 17, "0") & vbCrLf
                    Saldo += .Item("Debitos") - .Item("Creditos")
                    Inicial = Saldo
                    If Inicial < 0 Then
                        Inicial = Inicial * -1
                    End If

                    linha &= Funcoes.AlinharDireita(Replace((Inicial.ToString), ",", ""), 17, "0")
                    If Saldo > 0 Then
                        linha &= "D"
                    Else
                        linha &= "C"
                    End If
                End With

                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
            End If
        Next

        MsgBox(Me.Page, "Processo concluído com Sucesso.", eTitulo.Sucess)
        ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "mensagemServidor = '" & Mensagem & "';", True)
    End Sub

End Class