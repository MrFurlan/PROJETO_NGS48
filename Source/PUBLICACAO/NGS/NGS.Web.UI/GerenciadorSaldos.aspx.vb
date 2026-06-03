Imports NGS.Lib.Negocio

Public Class GerenciadorSaldos
    Inherits BasePage

    Private Structure RazaoExcel
        Public Empresa As String
        Public EndEmpresa As Integer
        Public Conta As String
        Public Cliente As String
        Public EndCliente As String
        Public UnidadeDeNegocio As String
        Public TipoDeLancamento As Integer
        Public Produto As String
        Public Custo As String
        Public Lote As Integer
        Public ValorOficial As String
        Public ValorMoeda As String
        Public Quantidade As String
        Public Historico As String
        Public DebitoCredito As String
        Public Movimento As Date
        Public Sequencia As Long
    End Structure

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not String.IsNullOrWhiteSpace(Request.QueryString("mnu")) AndAlso Request.QueryString.AllKeys.Contains("mnu") Then Me.setMenu(eModulo.Gerencial) Else Me.setMenu(eModulo.Contabil)
        'If Not IsPostBack And IsConnect Then
        '    If Funcoes.VerificaPermissao("GerenciadorSaldos", "ACESSAR") Then
        '        CargaUnidadeDeNegocioEmpresa()
        '        Limpar()
        '    Else
        MsgBox(Me.Page, "Usuário sem permissão para acessar essa página", "~/Contabil.aspx")
        '        Exit Sub
        '    End If
        'End If
    End Sub

    Protected Sub ddlUnidadeDeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeDeNegocio.SelectedIndexChanged
        Try
            ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeDeNegocio.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CargaUnidadeDeNegocioEmpresa()
        ddl.Carregar(ddlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeDeNegocio, ddlEmpresa, True)
    End Sub

    Private Sub Limpar()
        txtDataInicial.Text = Format(Today, "dd/MM/yyyy")
        txtDataFinal.Text = Format(Today, "dd/MM/yyyy")
    End Sub

    Protected Sub lnkGerarLote_Click(sender As Object, e As EventArgs) Handles lnkGerarLote.Click
        Try
            Dim sqls As New ArrayList
            Dim SequenciaPrimeiro As Long
            Dim SequenciaUltimo As Long

            Dim sql As String = " DELETE RAZAO " & vbCrLf & _
                                "   WHERE   left(Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                "       AND Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "' And Lote_Id in (6000, 6001)"
            If Banco.GravaBanco(sql) Then

                sql = "  SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, " & vbCrLf & _
                                "           SUM(CreditoOficial - DebitoOficial) AS Saldo   " & vbCrLf & _
                                "       FROM    Razao" & vbCrLf & _
                                "       WHERE   Conta_Id in ('1010201', '1010202')   " & vbCrLf & _
                                "           AND left(Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                "           AND Movimento_Id BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                                "       GROUP BY    Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_ID" & vbCrLf & _
                                "                   Having(SUM(CreditoOficial - DebitoOficial) > 0)" & vbCrLf & _
                                "" & vbCrLf & _
                                "       Union" & vbCrLf & _
                                "" & vbCrLf & _
                                "   SELECT  Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_Id, " & vbCrLf & _
                                "           SUM(DebitoOficial - CreditoOficial) AS Saldo   " & vbCrLf & _
                                "       FROM    Razao" & vbCrLf & _
                                "       WHERE   Conta_Id in ('2010101', '2010102', '2010104', '2010107', '2010108', '2010109', '2010111')   " & vbCrLf & _
                                "           AND left(Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                "           AND Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf & _
                                "       GROUP BY    Empresa_Id, EndEmpresa_Id, Conta_Id, Cliente_Id, EndCliente_ID" & vbCrLf & _
                                "                   Having(SUM(DebitoOficial - CreditoOficial) > 0)" & vbCrLf

                Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Razao")

                If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then

                    For Each row As DataRow In ds.Tables(0).Rows
                        Dim objRazao As New RazaoExcel
                        objRazao.Cliente = row("Empresa_Id")
                        objRazao.EndEmpresa = row("EndEmpresa_Id")
                        objRazao.Conta = row("Conta_Id")
                        objRazao.Cliente = row("Cliente_Id")
                        objRazao.EndCliente = row("EndCliente_Id")
                        objRazao.UnidadeDeNegocio = "01"
                        objRazao.Produto = ""
                        objRazao.Custo = ""
                        objRazao.Movimento = txtDataInicial.Text
                        objRazao.Lote = 6000
                        SequenciaUltimo = SequenciaUltimo + 1
                        objRazao.Sequencia = SequenciaUltimo
                        objRazao.ValorOficial = Replace(row("Saldo"), ".", "")
                        objRazao.ValorMoeda = 0
                        objRazao.Quantidade = 0
                        objRazao.Historico = "Valor Transferido para Conta de Adiantamento"
                        objRazao.TipoDeLancamento = 1

                        If Left(row("Conta_Id"), 1) = 1 Then
                            objRazao.DebitoCredito = "D"
                        Else
                            objRazao.DebitoCredito = "C"
                        End If

                        LanctosContabeis(objRazao, sqls)

                        SequenciaUltimo = SequenciaUltimo + 1
                        objRazao.Sequencia = SequenciaUltimo

                        Select Case row("Conta_Id")
                            Case "1010201"
                                objRazao.Conta = "2010601"
                            Case "1010202"
                                objRazao.Conta = "2010601"
                            Case "2010101"
                                objRazao.Conta = "1010303"
                            Case "2010102"
                                objRazao.Conta = "1010303"
                            Case "2010104"
                                objRazao.Conta = "1010305"
                            Case "2010107"
                                objRazao.Conta = "1010302"
                            Case "2010108"
                                objRazao.Conta = "1010304"
                            Case "2010109"
                                objRazao.Conta = "1010302"
                            Case "2010111"
                                objRazao.Conta = "1010305"
                        End Select

                        If Left(row("Conta_Id"), 1) = 1 Then
                            objRazao.DebitoCredito = "C"
                        Else
                            objRazao.DebitoCredito = "D"
                        End If

                        LanctosContabeis(objRazao, sqls)

                        '------Dia Primeiro
                        objRazao.Empresa = row("Empresa_Id")
                        objRazao.EndEmpresa = row("EndEmpresa_Id")
                        objRazao.Conta = row("Conta_Id")
                        objRazao.Cliente = row("Cliente_Id")
                        objRazao.EndCliente = row("EndCliente_Id")
                        objRazao.UnidadeDeNegocio = "01"
                        objRazao.Produto = ""
                        objRazao.Custo = ""
                        objRazao.Movimento = txtDataFinal.Text
                        objRazao.Lote = 6001
                        SequenciaPrimeiro = SequenciaPrimeiro + 1
                        objRazao.Sequencia = SequenciaPrimeiro
                        objRazao.ValorOficial = Replace(row("Saldo"), ".", "")
                        objRazao.ValorMoeda = 0
                        objRazao.Quantidade = 0
                        objRazao.Historico = "Valor Transferido da Conta de Adiantamento"
                        objRazao.TipoDeLancamento = 1

                        If Left(row("Conta_Id"), 1) = 1 Then
                            objRazao.DebitoCredito = "C"
                        Else
                            objRazao.DebitoCredito = "D"
                        End If

                        LanctosContabeis(objRazao, sqls)

                        SequenciaPrimeiro = SequenciaPrimeiro + 1
                        objRazao.Sequencia = SequenciaPrimeiro

                        Select Case row("Conta_Id")
                            Case "1010201"
                                objRazao.Conta = "2010601"
                            Case "1010202"
                                objRazao.Conta = "2010601"
                            Case "2010101"
                                objRazao.Conta = "1010303"
                            Case "2010102"
                                objRazao.Conta = "1010303"
                            Case "2010104"
                                objRazao.Conta = "1010305"
                            Case "2010107"
                                objRazao.Conta = "1010302"
                            Case "2010108"
                                objRazao.Conta = "1010304"
                            Case "2010109"
                                objRazao.Conta = "1010302"
                            Case "2010111"
                                objRazao.Conta = "1010305"
                        End Select

                        If Left(row("Conta_Id"), 1) = 1 Then
                            objRazao.DebitoCredito = "D"
                        Else
                            objRazao.DebitoCredito = "C"
                        End If

                        LanctosContabeis(objRazao, sqls)
                    Next

                    If Banco.GravaBanco(sqls) Then
                        MsgBox(Me.Page, "Dados registrado com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, Session("ssMessage"))
                    End If
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LanctosContabeis(ByVal objRazao As RazaoExcel, ByRef sqls As ArrayList)
        Try
            Dim sql As String = "INSERT INTO Razao " & vbCrLf & _
                            "                   (Empresa_Id, " & vbCrLf & _
                            "                   EndEmpresa_Id, " & vbCrLf & _
                            "                   Conta_Id, " & vbCrLf & _
                            "                   Cliente_Id, " & vbCrLf & _
                            "                   EndCliente_Id, " & vbCrLf & _
                            "                   Movimento_Id, " & vbCrLf & _
                            "                   Lote_Id, " & vbCrLf & _
                            "                   Sequencia_Id, " & vbCrLf & _
                            "                   Titulo, " & vbCrLf & _
                            "                   UnidadeDeNegocio, " & vbCrLf & _
                            "                   Produto, " & vbCrLf & _
                            "                   Custo, " & vbCrLf & _
                            "                   Serie_Nf, " & vbCrLf & _
                            "                   Indexador, " & vbCrLf & _
                            "                   DataMoeda, " & vbCrLf & _
                            "                   DebitoOficial, " & vbCrLf & _
                            "                   CreditoOficial, " & vbCrLf & _
                            "                   DebitoMoeda, " & vbCrLf & _
                            "                   CreditoMoeda, " & vbCrLf & _
                            "                   DebitoQuantidade, " & vbCrLf & _
                            "                   CreditoQuantidade, " & vbCrLf & _
                            "                   Historico, " & vbCrLf & _
                            "                   PrevistoRealizado)" & vbCrLf & _
                            "           VALUES ('" & objRazao.Empresa & "'" & vbCrLf & _
                            ",              " & objRazao.EndEmpresa & vbCrLf & _
                            ",              '" & objRazao.Conta & "'" & vbCrLf

            If Len(objRazao.Conta) = 7 Then
                sql &= ",       '" & objRazao.Cliente & "'" & vbCrLf & _
                       ",        " & objRazao.EndCliente & vbCrLf
            Else
                sql &= ", ''" & vbCrLf & _
                       ", 0" & vbCrLf
            End If

            sql &= ", '" & objRazao.Movimento.ToSqlDate() & "'" & vbCrLf & _
                   ", " & objRazao.Lote & vbCrLf & _
                   ", " & objRazao.Sequencia & vbCrLf & _
                   ", 0" & vbCrLf & _
                   ", '" & objRazao.UnidadeDeNegocio & "'" & vbCrLf & _
                   ", '" & objRazao.Produto & "'" & vbCrLf & _
                   ", '" & objRazao.Custo & "'" & vbCrLf & _
                   ", '" & objRazao.TipoDeLancamento & "'" & vbCrLf & _
                   ", 3" & vbCrLf & _
                   ", '" & objRazao.Movimento.ToSqlDate() & "'" & vbCrLf

            If objRazao.DebitoCredito = "D" Then
                sql &= ", " & Replace(objRazao.ValorOficial, ",", ".") & vbCrLf & _
                            ", 0.0" & vbCrLf & _
                            ", " & Replace(objRazao.ValorMoeda, ",", ".") & vbCrLf & _
                            ", 0.0" & vbCrLf & _
                            ", " & Replace(objRazao.Quantidade, ",", ".") & vbCrLf & _
                            ", 0.0" & vbCrLf
            Else
                sql &= ", 0.0" & vbCrLf & _
                            ", " & Replace(objRazao.ValorOficial, ",", ".") & vbCrLf & _
                            ", 0.0" & vbCrLf & _
                            ", " & Replace(objRazao.ValorMoeda, ",", ".") & vbCrLf & _
                            ", 0.0" & vbCrLf & _
                            ", " & Replace(objRazao.Quantidade, ",", ".") & vbCrLf
            End If

            sql &= ", '" & objRazao.Historico & "'" & vbCrLf & _
                   ", 'P')" & vbCrLf

            sqls.Add(sql)

        Catch ex As Exception
            Throw New Exception(ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcluirLote_Click(sender As Object, e As EventArgs) Handles lnkExcluirLote.Click
        Try
            If validaExclusao() Then
                Dim Sql As String = " DELETE RAZAO " & vbCrLf & _
                                    "   WHERE   left(Empresa_Id, 8) = '" & Left(ddlEmpresa.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf & _
                                    "       AND Movimento_Id between '" & txtDataInicial.Text.ToSqlDate() & "' And '" & txtDataFinal.Text.ToSqlDate() & "' And Lote_Id in (6000, 6001)"

                If Banco.GravaBanco(Sql) Then
                    MsgBox(Me.Page, "Informações excluídas com Sucesso.", eTitulo.Sucess)
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function validaExclusao()
        If String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf String.IsNullOrWhiteSpace("") Then
            MsgBox(Me.Page, "Informe as datas para exclusão.")
            Return False
        End If

        Return True
    End Function

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "GerenciadorSaldos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class