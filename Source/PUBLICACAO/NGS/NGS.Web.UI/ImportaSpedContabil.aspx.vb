Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ImportaSpedContabil
    Inherits BasePage
    Dim path As String

    Dim dsArquivo As New DataSet
    Dim dsEmpresa As New DataSet

    Dim dr As DataRow
    Dim SEmpresa() As String

    Dim Sql As String
    Dim SqlArray As New ArrayList
    Dim Ds As DataSet

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then

            If Funcoes.VerificaPermissao("SpedContabil", "ACESSAR") Then
                Dim Codigo As String
                Dim Descricao As String
                Dim Nome As String
                Dim Cidade As String
                Dim Cnpj As String

                DdlEmpresa.Items.Clear()
                Sql = "  SELECT Clientes.Cliente_Id AS Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf & _
                      " FROM   Clientes INNER JOIN" & vbCrLf & _
                      "        ClientesXEmpresas ON " & vbCrLf & _
                      "        Clientes.Cliente_Id  = ClientesXEmpresas.Empresa_Id AND " & vbCrLf & _
                      "        Clientes.Endereco_Id = ClientesXEmpresas.EndEmpresa_Id" & vbCrLf & _
                      " Where	ClientesXEmpresas.Matriz = 'S'" & vbCrLf

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                    Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
                    Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
                    Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
                    Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
                    Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
                    Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
                    DdlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
                Next

                DdlEmpresa.Items.Insert(0, "")
                DdlEmpresa.SelectedIndex = 0
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub cmdPlanoDeContas_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dsaux As New Data.DataSet
        Dim Codigo, Empresa, Titulo, Cliente, CentroDeCusto, Produto As String
        Dim PrimeiraImportacao As Boolean = True
        Dim count As Integer = 0

        Try
            SEmpresa = DdlEmpresa.SelectedValue.Split("-")

            path = txtPlanoDeContas.Text

            If txtPlanoDeContas.Text = Nothing Then
                MsgBox(Me.Page, "Selecione o arquivo de origem.")
                Exit Sub
            End If

            Empresa = SEmpresa(0)

            If (File.Exists(path)) Then
                Dim enc As Encoding = Encoding.ASCII
                Dim sr As StreamReader = New StreamReader(path, enc)
                Dim line As String

                Try
                    Do
                        line = sr.ReadLine()
                        If line Is Nothing Or Mid(line, 7, 9) = Nothing Then
                            Exit Do
                        End If
                        Codigo = Mid(line, 7, 9)
                        If Mid(Codigo, 2, 9) = "00000000" Then
                            Codigo = Mid(Codigo, 1, 1)
                        End If
                        If Mid(Codigo, 3, 9) = "0000000" Then
                            Codigo = Mid(Codigo, 1, 2)
                        End If
                        If Mid(Codigo, 4, 9) = "000000" Then
                            Codigo = Mid(Codigo, 1, 3)
                        End If
                        If Mid(Codigo, 6, 9) = "0000" Then
                            Codigo = Mid(Codigo, 1, 5)
                        End If

                        Titulo = RTrim(Mid(line, 17, 40)).Replace("'", "")
                        Produto = Mid(line, 58, 1)
                        CentroDeCusto = Mid(line, 60, 1)
                        Cliente = Mid(line, 57, 1)


                        Sql = "select * from PlanoDeContas "
                        Sql &= " WHERE Empresa_Id = '" & Empresa & "' And Conta_Id = '" & Codigo & "'"
                        dsEmpresa = Banco.ConsultaDataSet(Sql, "Clientes")

                        If dsEmpresa.Tables(0).Rows.Count = 0 Then
                            Sql = "  INSERT INTO PlanoDeContas (Empresa_Id, EndEmpresa_Id, Conta_Id, " & vbCrLf & _
                                  " Titulo, Cliente, Produto, CentroDeCusto, TipoDeCliente, " & vbCrLf & _
                                  " Responsabilidade, ContaOrcamentaria, " & vbCrLf & _
                                  " ContaAnterior)" & vbCrLf & _
                                  " VALUES(" & vbCrLf & _
                                  "'" & Empresa & "'," & vbCrLf & _
                                  "" & 0 & "," & vbCrLf & _
                                  " '" & Codigo & "'" & vbCrLf & _
                                  ", '" & Titulo & "'" & vbCrLf & _
                                  ", '" & Cliente & "'" & vbCrLf & _
                                  ", '" & Produto & "'" & vbCrLf & _
                                  ",'" & CentroDeCusto & "' " & vbCrLf & _
                                  ", 0" & vbCrLf & _
                                  ", ''" & vbCrLf & _
                                  ", ''" & vbCrLf & _
                                  ", '')" & vbCrLf

                            SqlArray.Add(Sql)

                            If Banco.GravaBanco(SqlArray) = False Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                SqlArray.Clear()
                            Else
                                SqlArray.Clear()
                            End If
                        End If
                    Loop Until line Is Nothing
                    sr.Close()
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
                MsgBox(Me.Page, "Arquivo importado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Arquivo não encontrado.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdCustos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dsaux As New Data.DataSet
        Dim Codigo, Descricao As String
        Dim count As Integer = 0

        Try
            path = txtCentroDeCusto.Text

            If txtCentroDeCusto.Text = Nothing Then
                MsgBox(Me.Page, "Selecione o arquivo de origem.")
                Exit Sub
            End If

            If (File.Exists(path)) Then
                Dim enc As Encoding = Encoding.ASCII
                Dim sr As StreamReader = New StreamReader(path, enc)
                Dim line As String
                Try
                    Do
                        line = sr.ReadLine()

                        If line Is Nothing Or Mid(line, 7, 9) = Nothing Then
                            Exit Do
                        End If

                        Codigo = Mid(line, 6, 3)
                        Descricao = RTrim(Mid(line, 9, 40)).Replace("'", "")

                        Sql = "select * from CentrosDeCustos " & vbCrLf & _
                              "  where Ativo = 1 and CentroDeCusto_Id = '" & Codigo & "'" & vbCrLf
                        dsEmpresa = Banco.ConsultaDataSet(Sql, "Custos")

                        If dsEmpresa.Tables(0).Rows.Count = 0 Then
                            Sql = "  INSERT INTO CentrosDeCustos (CentroDeCusto_Id, Descricao)" & vbCrLf & _
                                  " VALUES(" & vbCrLf & _
                                  "'" & Codigo & "'" & vbCrLf & _
                                  ", '" & Descricao & "')" & vbCrLf
                            SqlArray.Add(Sql)

                            If Banco.GravaBanco(SqlArray) = False Then
                                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                                SqlArray.Clear()
                            Else
                                SqlArray.Clear()
                            End If
                        End If
                    Loop Until line Is Nothing
                    sr.Close()
                Catch ex As Exception
                    MsgBox(Me.Page, ex.Message, eTitulo.Erro)
                End Try
                MsgBox(Me.Page, "Arquivo importado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, "Arquivo não encontrado.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim dsaux As New Data.DataSet
        Dim Reduzido, Nome, Endereco, Numero, Complemento, Bairro, Cidade, Cep, Estado, Cliente_Id, Inscricao, FisJur, Pais As String
        Dim count As Integer = 0

        Try
            path = txtClientes.Text

            If txtClientes.Text = Nothing Then
                MsgBox(Me.Page, "Selecione o arquivo de origem.")
                Exit Sub
            End If

            If (File.Exists(path)) Then
                Dim enc As Encoding = Encoding.ASCII
                Dim sr As StreamReader = New StreamReader(path, enc)
                Dim line As String
                Try
                    Do
                        line = sr.ReadLine()
                        If line Is Nothing Or Mid(line, 12, 5) = Nothing Then
                            Exit Do
                        End If
                        Reduzido = Mid(line, 12, 5)
                        Nome = RTrim(Mid(line, 27, 40)).Replace("'", "")
                        Endereco = RTrim(Mid(line, 67, 30)).Replace("'", "")
                        Numero = RTrim(Mid(line, 97, 5)).Replace("?", "")
                        Complemento = RTrim(Mid(line, 102, 15)).Replace("'", "")
                        Bairro = RTrim(Mid(line, 117, 30)).Replace("'", "")
                        Cidade = RTrim(Mid(line, 155, 30)).Replace("'", "")
                        Cep = RTrim(Mid(line, 185, 8))
                        Estado = Mid(line, 193, 2)
                        Cliente_Id = Mid(line, 7, 3) & Mid(line, 10, 2) & Mid(line, 12, 5)
                        Inscricao = Trim(".")
                        FisJur = Mid(line, 195, 1)
                        Pais = "1058"

                        Sql = "select * from Clientes " & vbCrLf & _
                              " WHERE Cliente_ID = '" & Cliente_Id & "'" & vbCrLf
                        dsEmpresa = Banco.ConsultaDataSet(Sql, "Clientes")

                        If dsEmpresa.Tables(0).Rows.Count = 0 Then
                            Sql = " INSERT INTO Clientes (Cliente_Id, Endereco_Id, Regiao, Categoria, Estado, " & vbCrLf & _
                                  " Nome,Fantasia, Endereco, Bairro, Cep, Cidade, Inscricao,  Reduzido, Numero, Complemento, Situacao, Pais)" & vbCrLf & _
                                  " VALUES('" & Cliente_Id & "'," & vbCrLf & _
                                  " " & 0 & "," & vbCrLf & _
                                  " " & 1 & "," & vbCrLf & _
                                  " " & 4 & "," & vbCrLf & _
                                  " '" & Estado & "'," & vbCrLf & _
                                  " '" & Nome & "'," & vbCrLf & _
                                  " '" & Nome & "'," & vbCrLf & _
                                  " '" & Endereco & "'," & vbCrLf & _
                                  " '" & Bairro & "'," & vbCrLf & _
                                  " '" & Cep & "'," & vbCrLf & _
                                  " '" & Cidade & "'," & vbCrLf & _
                                  " '" & Inscricao & "'," & vbCrLf & _
                                  " '" & Reduzido & "'," & vbCrLf & _
                                  " " & Numero & "," & vbCrLf & _
                                  " '" & Complemento & "'," & vbCrLf & _
                                  " " & 1 & "," & vbCrLf & _
                                  " '" & Pais & "')" & vbCrLf
                            SqlArray.Add(Sql)
                        End If
                        If Banco.GravaBanco(SqlArray) = False Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            SqlArray.Clear()
                        Else
                            SqlArray.Clear()
                            'sr.Close()
                        End If
                    Loop Until line Is Nothing
                    sr.Close()
                Catch ex As Exception
                    MsgBox(Me.Page, "Problema na importação.")
                    sr.Close()
                Finally
                    MsgBox(Me.Page, "Arquivo importado com Sucesso.", eTitulo.Sucess)
                    sr.Close()
                End Try
            Else
                MsgBox(Me.Page, "Arquivo não encontrado.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdRazao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim SqlArray As New ArrayList
        Dim dsaux As New Data.DataSet
        Dim sql, path As String
        Dim count As Integer = 0
        Dim Sequencia As Integer = 0
        Dim Filial As Integer = 0

        Try
            path = txtRazao.Text
            SEmpresa = DdlEmpresa.SelectedValue.Split("-")

            If txtRazao.Text = Nothing Then
                MsgBox(Me.Page, "Selecione o arquivo de origem.")
                Exit Sub
            End If
            If (File.Exists(path)) Then
                Dim enc As Encoding = Encoding.ASCII
                Dim sr As StreamReader = New StreamReader(path, enc)
                Dim line As String
                Try
                    Do
                        Dim Codigo, Empresa, Cliente, Movimento_Id, Lote, Titulo, TipoLancamento, CentroCusto, DebCre, Historico As String
                        Dim ValorDebito As Double
                        Dim ValorCredito As Double

                        line = sr.ReadLine()
                        If line Is Nothing Or Mid(line, 7, 5) = Nothing Then
                            Exit Do
                        End If
                        Empresa = SEmpresa(0)
                        Codigo = Mid(line, 12, 9)
                        If Mid(Codigo, 2, 9) = "00000000" Then
                            Codigo = Mid(Codigo, 1, 1)
                        End If
                        If Mid(Codigo, 3, 9) = "0000000" Then
                            Codigo = Mid(Codigo, 1, 2)
                        End If
                        If Mid(Codigo, 4, 9) = "000000" Then
                            Codigo = Mid(Codigo, 1, 3)
                        End If
                        If Mid(Codigo, 6, 9) = "0000" Then
                            Codigo = Mid(Codigo, 1, 5)
                        End If
                        If Mid(line, 21, 5) <> "00000" Then
                            Cliente = Mid(line, 7, 5) & Mid(line, 21, 5)
                        Else
                            Cliente = ""
                        End If
                        If Mid(line, 34, 2) = "01" And Mid(line, 32, 2) = "01" Then
                            Movimento_Id = "31" & "/" & "12" & "/" & "2009"
                        Else
                            Movimento_Id = Mid(line, 34, 2) & "/" & Mid(line, 32, 2) & "/" & "2010"
                        End If

                        Lote = Mid(line, 36, 4)
                        Sequencia += 1
                        Titulo = Mid(line, 258, 6)
                        TipoLancamento = Mid(line, 58, 3)
                        CentroCusto = Mid(line, 384, 3)
                        DebCre = Mid(line, 95, 1)
                        Filial = Mid(line, 7, 5)

                        If DebCre = "D" Then
                            ValorDebito = (Mid(line, 61, 17) / 100)
                            ValorCredito = 0
                        Else
                            ValorCredito = (Mid(line, 61, 17) / 100)
                            ValorDebito = 0
                        End If

                        Historico = RTrim(Mid(line, 99, 45)) & " " & RTrim(Mid(line, 144, 45))

                        If Mid(line, 34, 2) <> "00" And Lote < 9000 And Mid(Codigo, 1, 1) < 5 And Mid(Codigo, 1, 1) <> "0" And (ValorDebito <> 0 Or ValorCredito <> 0) Then

                            count += 1

                            sql = "  INSERT INTO Razao (Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, " & vbCrLf & _
                                  " EndCliente_Id, Movimento_Id, Lote_Id, Sequencia_Id, Titulo, Custo, Indexador, " & vbCrLf & _
                                  " DataMoeda, DebitoOficial, CreditoOficial,DebitoMoeda,CreditoMoeda, Historico, PrevistoRealizado, Numero_Nf)" & vbCrLf & _
                                  "VALUES('" & Empresa & "'" & vbCrLf & _
                                  ", " & "0" & " " & vbCrLf & _
                                  ",'" & Codigo & "' " & vbCrLf
                            If Cliente = Nothing Then
                                sql &= ", '" & "" & "'"
                            Else
                                sql &= ", '" & Cliente & "'"
                            End If

                            sql &= ", " & "0" & "" & vbCrLf & _
                                   ", '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf & _
                                   ", " & Lote & "" & vbCrLf & _
                                   ", " & Sequencia & vbCrLf & _
                                   ", " & Titulo & "" & vbCrLf & _
                                   ", " & CentroCusto & "" & vbCrLf & _
                                   ", " & "1" & "" & vbCrLf & _
                                   ", '" & Movimento_Id.ToSqlDate() & "'" & vbCrLf & _
                                   ", " & Replace(ValorDebito, ",", ".") & " " & vbCrLf & _
                                   ", " & Replace(ValorCredito, ",", ".") & " " & vbCrLf & _
                                   ", " & 0 & " " & vbCrLf & _
                                   ", " & 0 & " " & vbCrLf & _
                                   ", '" & Historico.Replace("'", " ") & "' " & vbCrLf & _
                                   ",'R'" & vbCrLf & _
                                   ", " & Filial & ")" & vbCrLf
                            SqlArray.Add(sql)
                        End If

                        If Banco.GravaBanco(SqlArray) Then
                            MsgBox(Me.Page, "Sucesso na inclusão.", eTitulo.Info)
                            SqlArray.Clear()
                        End If

                    Loop Until line Is Nothing
                    sr.Close()
                    MsgBox(Me.Page, "Arquivo importado com Sucesso.", eTitulo.Sucess)
                Catch ex As Exception
                    HttpContext.Current.Session("ssMessage") = ex.Message
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"), eTitulo.Erro)
                    sr.Close()
                End Try
            Else
                MsgBox(Me.Page, "Arquivo não encontrado.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ImportaSpedContabil")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class