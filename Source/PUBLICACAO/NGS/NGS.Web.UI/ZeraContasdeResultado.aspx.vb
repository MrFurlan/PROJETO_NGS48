Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ZeraContasdeResultado
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim scriptManager As ScriptManager = scriptManager.GetCurrent(Me)
        scriptManager.AsyncPostBackTimeout = 3600 * 5
        Me.setMenu(eModulo.Contabil)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/Contabil.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("ZeraContasdeResultado", "ACESSAR") Then
                VerificaCheck()

                ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.EmpresasConsolidadas, "")

                ddlEmpresa.SelectedValue = Left(Session("ssEmpresa"), 8)

                LiberaEmpresa()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Contabil.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub txtConta0_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            VerificaContas(txtConta0.Text)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkZerarContas_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lnkZerarContas.Click
        Try
            If ValidaCampos() Then
                If chkContasDeEstoque.Checked Then
                    ZerarContasEstoque()
                    MsgBox(Me.Page, "Contas de Estoque Zeradas")
                Else
                    ZerarContas()
                    MsgBox(Me.Page, "Contas de Resultado Zeradas")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub chkContasDeEstoque_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            VerificaCheck()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Function VerificaContas(ByVal conta_id As String) As Boolean
        Dim sql As String
        Dim dsContas As New DataSet
        Dim dr As DataRow

        sql = "select Conta_Id, Titulo" & vbCrLf & _
              "  From PlanoDeContas" & vbCrLf & _
              " where Conta_Id      ='" & conta_id & "' " & vbCrLf
        '"   and Empresa_Id    ='" & txtEmpresa.Text & "' " & vbCrLf & _
        '"   and EndEmpresa_Id = " & txtEndEmpresa.Text & "  " & vbCrLf

        dsContas = Banco.ConsultaDataSet(sql, "PlanoDeContas")

        If dsContas.Tables(0).Rows.Count = 0 Then
            txtConta0.Text = ""
            txtConta1.Text = ""
            txtConta2.Text = ""
            Return False
        Else
            For Each dr In dsContas.Tables(0).Rows
                txtConta0.Text = Mid(dr("conta_id"), 1, 7)
                txtConta1.Text = Mid(dr("conta_id"), 8, 2)
                txtConta2.Text = dr("Titulo")
            Next
        End If
        Return True
    End Function

    Public Sub VerificaCheck()
        If chkContasDeEstoque.Checked Then
            If Left(ddlEmpresa.SelectedValue.ToString, 8) = "03189063" Then
                VerificaContas("101060101")
            ElseIf Left(ddlEmpresa.SelectedValue.ToString, 8) = "05272759" Then
                VerificaContas("101060101")
            Else
                VerificaContas("101040101")
            End If

            lblLote.Text = "Lote: 7600"
            txtConta0.Enabled = False
            txtConta1.Enabled = False
            txtConta2.Enabled = False
        Else
            VerificaContas("307010101")
            lblLote.Text = "Lote: 7500"
            txtConta0.Enabled = True
            txtConta1.Enabled = True
            txtConta2.Enabled = True
        End If
    End Sub

    Public Function ValidaCampos() As Boolean
        If ddlEmpresa.SelectedIndex = 0 Then
            MsgBox(Me.Page, "Informe a empresa.")
            Return False
        ElseIf (String.IsNullOrEmpty(txtDataInicio.Text) Or String.IsNullOrEmpty(txtDataFim.Text)) Then
            MsgBox(Me.Page, "Informe a período.")
            Return False
        ElseIf txtDataInicio.Text > txtDataFim.Text Then
            MsgBox(Me.Page, "Período inicial e maior que o período final.")
            Return False
        ElseIf txtConta2.Text.Length = 0 And chkContasDeEstoque.Checked = False Then
            MsgBox(Me.Page, "Se não informar a conta só pode continuar se for Zerar Contas de Estoque.")
            Return False
        ElseIf Not CDate(txtDataInicio.Text).Year = CDate(txtDataFim.Text).Year Then
            MsgBox(Me.Page, "Ano do período inicial é diferente do período final.")
            Return False
        End If
        Return True
    End Function

#Region "Zerar Contas"

    Private Sub ZerarContas()
        Dim sql As String
        Dim dscontas As DataSet
        Dim tabela As String = "Razao"

        ExcluirLote()

        sql = "SELECT Empresa_Id, EndEmpresa_Id " & vbCrLf & _
              "  FROM ClientesXEmpresas" & vbCrLf & _
              " Where left(Empresa_Id,8) = '" & ddlEmpresa.SelectedValue.ToString & "'" & vbCrLf

        Dim dsEmpresas As DataSet = Banco.ConsultaDataSet(sql, "Empresas")

        For Each Row As DataRow In dsEmpresas.Tables(0).Rows
            sql = " SELECT Empresa_Id, EndEmpresa_Id," & vbCrLf & _
                  "        Conta_Id, " & vbCrLf & _
                  "        ISNULL(Produto, '') as Produto," & vbCrLf & _
                  "        case" & vbCrLf & _
                  "          when isnull(custo,'0') = '0'" & vbCrLf & _
                  "             then ''" & vbCrLf & _
                  "             else custo" & vbCrLf & _
                  "        end Custo, " & vbCrLf & _
                  "        SUM(DebitoOficial - CreditoOficial) AS Oficial, " & vbCrLf & _
                  "        SUM(DebitoMoeda - CreditoMoeda) AS Moeda " & vbCrLf & _
                  "   FROM Razao " & vbCrLf & _
                  "  WHERE (Conta_Id > '2999999999999') " & vbCrLf & _
                  "    AND (LEN(Conta_Id) > 7) " & vbCrLf & _
                  "    AND Empresa_Id    ='" & Row("Empresa_Id") & "'" & vbCrLf & _
                  "    AND EndEmpresa_Id = " & Row("EndEmpresa_id")

            sql &= "   AND Movimento_Id BETWEEN '" & CDate(txtDataInicio.Text).ToString("yyyy-MM-dd") & "'" & " AND '" & CDate(txtDataFim.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                   " GROUP BY Empresa_Id, EndEmpresa_Id, Conta_Id, ISNULL(Produto, ''), case when isnull(custo,'0') = '0' then '' else custo end  "

            dscontas = Banco.ConsultaDataSet(sql, tabela)
            ExecutarLancamentos(dscontas, tabela)
        Next
    End Sub

    Private Sub ExcluirLote()
        Dim sql As String
        Dim arrsql As New ArrayList

        sql = "DELETE Razao" & vbCrLf & _
              " WHERE Lote_Id       = 7500 " & vbCrLf & _
              "   AND Movimento_Id BETWEEN '" & CDate(txtDataInicio.Text).ToString("yyyy-MM-dd") & "'" & " AND '" & CDate(txtDataFim.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              "   AND left(Empresa_Id, 8) = '" & ddlEmpresa.SelectedValue.ToString & "'" & vbCrLf & vbCrLf & _
              "   AND left(Conta_Id, 1) > 1"

        arrsql.Add(sql)

        If Not Banco.GravaBanco(arrsql) Then
            MsgBox(Me.Page, "Não foi possível excluir os últimos registros.")
        End If
    End Sub

    Private Sub ExecutarLancamentos(ByRef dscontas As DataSet, ByVal tabela As String)
        Dim drcontas As DataRow = Nothing
        Dim sql As String = ""
        Dim sequencia As Integer
        Dim arrsql As New ArrayList
        Dim oficial, moeda, creditooficial, debitooficial, creditomoeda, debitomoeda As Decimal

        sequencia = 0
        creditooficial = 0
        debitooficial = 0
        creditomoeda = 0
        debitomoeda = 0

        For Each drcontas In dscontas.Tables(tabela).Rows

            oficial = CDec(FormatNumber(drcontas("Oficial"), 2))
            moeda = CDec(FormatNumber(drcontas("Moeda"), 2))

            If oficial <> 0 Then
                sequencia += 1

                ' Moeda
                If drcontas("Moeda") < 0 Then
                    debitomoeda += Math.Abs(drcontas("Moeda"))
                Else
                    creditomoeda += Math.Abs(drcontas("Moeda"))
                End If

                ' Oficial
                If drcontas("Oficial") < 0 Then
                    debitooficial += Math.Abs(drcontas("Oficial"))
                Else
                    creditooficial += Math.Abs(drcontas("Oficial"))
                End If

                sql = sqlLancamento(drcontas("Empresa_Id"), drcontas("EndEmpresa_Id"), drcontas("Conta_Id"), drcontas("Produto"), drcontas("custo"), drcontas("Oficial"), drcontas("Moeda"), sequencia)

                If sql.Length > 0 Then
                    arrsql.Add(sql)
                End If
            End If
        Next

        oficial = debitooficial - creditooficial
        moeda = debitomoeda - creditomoeda

        ' Lanca a diferenca na conta especificada
        If oficial <> 0 Then
            sequencia += 1
            sql = sqlLancamento(drcontas("Empresa_Id"), drcontas("EndEmpresa_Id"), txtConta0.Text & txtConta1.Text, drcontas("Produto"), drcontas("custo"), oficial, moeda, sequencia)
            arrsql.Add(sql)
        End If

        If arrsql.Count > 0 Then
            If Not Banco.GravaBanco(arrsql) Then
                MsgBox(Me.Page, Session("ssMessage"))
            End If
        End If
    End Sub

    Private Function sqlLancamento(ByVal Empresa_Id As String, ByVal EndEmpresa_Id As String, ByVal Conta_Id As String, ByVal Produto_Id As String, ByVal CentroDeCustos As String, ByVal Oficial As Decimal, ByVal Moeda As Decimal, ByVal Sequencia_Id As Integer) As String
        Dim sql As String
        Dim debitooficial, debitomoeda, creditooficial, creditomoeda As Decimal

        debitooficial = 0
        debitomoeda = 0
        creditooficial = 0
        creditomoeda = 0

        If Moeda < 0 Then
            debitomoeda = Math.Abs(Moeda)
        Else
            creditomoeda = Moeda
        End If
        If Oficial < 0 Then
            debitooficial = Math.Abs(Oficial)
        Else
            creditooficial = Oficial
        End If

        sql = "  INSERT INTO Razao " & vbCrLf & _
              " (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
              " Conta_Id, " & vbCrLf & _
              " Cliente_Id, EndCliente_Id, " & vbCrLf & _
              " Movimento_Id, " & vbCrLf & _
              " Lote_Id, Sequencia_Id, " & vbCrLf & _
              " Produto, " & vbCrLf & _
              " Custo, " & vbCrLf & _
              " Indexador, " & vbCrLf & _
              " DataMoeda, " & vbCrLf & _
              " DebitoOficial, DebitoMoeda, " & vbCrLf & _
              " CreditoOficial, CreditoMoeda, " & vbCrLf & _
              " Historico, " & vbCrLf & _
              " PrevistoRealizado)" & vbCrLf & _
              " VALUES " & vbCrLf & _
              " ('" & Empresa_Id & "', " & EndEmpresa_Id & vbCrLf & _
              ", '" & Conta_Id & "'" & vbCrLf & _
              ", '', 0" & vbCrLf & _
              ", '" & CDate(txtDataFim.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf & _
              ", 7500, " & Sequencia_Id & vbCrLf & _
              ", '" & Produto_Id & "'" & vbCrLf & _
              ", '" & CentroDeCustos & "'" & vbCrLf & _
              ", 3" & vbCrLf & _
              ", '" & Format(Today(), "yyyy/MM/dd") & "'" & vbCrLf & _
              ", " & Str(debitooficial) & ", " & Str(debitomoeda) & vbCrLf & _
              ", " & Str(creditooficial) & ", " & Str(creditomoeda) & vbCrLf & _
              ", 'TRANSFERENCIA PARA O RESULTADO DO EXERCICIO'" & vbCrLf & _
              ", 'D')"

        Return sql
    End Function

#End Region

#Region "Zerar Contas Estoque"

    Private Sub ZerarContasEstoque()
        Dim sql As String
        Dim dscontas As DataSet
        Dim tabela As String = "Razao"
        Dim Ano As Integer = CDate(txtDataFim.Text).Year

        ExcluirLoteEstoque(Ano)

        'sql = " SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, ISNULL(Produto, 0) as Produto, Custo, " & vbCrLf & _
        '             "        SUM(DebitoOficial - CreditoOficial) AS Oficial, " & vbCrLf & _
        '             "        SUM(DebitoMoeda - CreditoMoeda) AS Moeda " & vbCrLf & _
        '             "   FROM Razao " & vbCrLf & _
        '             "  where ((conta_Id > '101060101' and conta_Id < '101060199') or (conta_Id > '101050200' and conta_Id < '101050299') )" & vbCrLf & _
        '             "    AND left(Empresa_Id,8) = '" & ddlEmpresa.SelectedValue.ToString & "'" & vbCrLf & _
        '             "    AND Movimento_Id BETWEEN '" & Ano & "-01-01' AND '" & Ano & "-12-31" & "' " & _
        '             "    And Lote_Id <> 5" & _
        '             "  GROUP BY Empresa_Id, EndEmpresa_Id, Conta_Id, ISNULL(Produto, 0), Custo "

        sql = " SELECT Empresa_Id, EndEmpresa_Id, Conta_Id, ISNULL(Produto, 0) as Produto, 0 As Custo, " & vbCrLf & _
              "        SUM(DebitoOficial - CreditoOficial) AS Oficial, " & vbCrLf & _
              "        0 AS Moeda " & vbCrLf & _
              "   FROM Razao " & vbCrLf

        If Left(ddlEmpresa.SelectedValue.ToString, 8) = "03189063" Then
            sql &= "  where ((conta_Id > '101060101' and conta_Id <= '101060199')  )" & vbCrLf
        ElseIf Left(ddlEmpresa.SelectedValue.ToString, 8) = "05272759" Or Left(ddlEmpresa.SelectedValue.ToString, 8) = "44979506" Then
            sql &= "  where ((conta_Id > '101060101' and conta_Id <= '101060199')  )" & vbCrLf
        ElseIf Left(ddlEmpresa.SelectedValue.ToString, 8) = "05366261" OrElse Left(ddlEmpresa.SelectedValue.ToString, 8) = "38198213" OrElse Left(ddlEmpresa.SelectedValue.ToString, 8) = "40938762" OrElse Left(ddlEmpresa.SelectedValue.ToString, 8) = "44005444" OrElse Left(ddlEmpresa.SelectedValue.ToString, 8) = "62747840" OrElse Left(ddlEmpresa.SelectedValue.ToString, 8) = "62780383" OrElse Left(ddlEmpresa.SelectedValue.ToString, 8) = "63358210" Then
            sql &= "  where ((conta_Id > '101060101' and conta_Id <= '101060199')  )" & vbCrLf
        Else
            sql &= "  where ((conta_Id >= '101040101' and conta_Id <= '101040199')  )" & vbCrLf & _
                   "    AND Conta_Id not in ('101040104','101040130','101040131') " & vbCrLf
        End If

        sql &= "    AND left(Empresa_Id,8) = '" & ddlEmpresa.SelectedValue.ToString & "'" & vbCrLf & _
               "    AND Movimento_Id <= '" & Ano & "-12-31'" & _
               "  GROUP BY Empresa_Id, EndEmpresa_Id, Conta_Id, ISNULL(Produto, 0), Custo " & vbCrLf & _
               "  Having(SUM(DebitoOficial - CreditoOficial) <> 0)" & vbCrLf & _
               "  Order by Produto" & vbCrLf

        dscontas = Banco.ConsultaDataSet(sql, tabela)
        ExecutarLancamentosEstoque(dscontas, tabela)
    End Sub

    Private Sub ExcluirLoteEstoque(ByVal ano As Integer)
        Dim sql As String
        Dim arrSql As New ArrayList

        sql = "DELETE Razao" & vbCrLf & _
              " WHERE Lote_Id       = 7600 " & vbCrLf & _
              "   AND Movimento_Id  = '" & (ano + 1) & "/01/01'" & vbCrLf & _
              "   AND left(Empresa_Id,8) = '" & ddlEmpresa.SelectedValue.ToString & "'" & vbCrLf & _
              "   AND left(Conta_Id, 1) = 1"

        arrSql.Add(sql)

        If Not Banco.GravaBanco(arrSql) Then
            MsgBox(Me.Page, "Não foi possível excluir os últimos registros.")
        End If
    End Sub

    Private Sub ExecutarLancamentosEstoque(ByRef dscontas As DataSet, ByVal tabela As String)
        Dim drcontas As DataRow
        Dim sql As String
        Dim sequencia As Integer
        Dim arrsql As New ArrayList
        Dim oficial, moeda, creditooficial, debitooficial, creditomoeda, debitomoeda As Decimal

        sequencia = 0
        creditooficial = 0
        debitooficial = 0
        creditomoeda = 0
        debitomoeda = 0

        For Each drcontas In dscontas.Tables(tabela).Rows

            oficial = CDec(FormatNumber(drcontas("Oficial"), 2))
            moeda = CDec(FormatNumber(drcontas("Moeda"), 2))


            If oficial <> 0 Then
                sequencia += 1

                ' MOEDA
                If drcontas("Moeda") < 0 Then
                    debitomoeda = Math.Abs(drcontas("Moeda"))
                Else
                    creditomoeda = Math.Abs(drcontas("Moeda"))
                End If
                ' OFICIAL
                If drcontas("Oficial") < 0 Then
                    debitooficial = Math.Abs(drcontas("Oficial"))
                Else
                    creditooficial = Math.Abs(drcontas("Oficial"))
                End If

                sql = sqlLancamentoEstoque(drcontas("Empresa_Id"), drcontas("EndEmpresa_Id"), drcontas("Conta_Id"), drcontas("Produto"), drcontas("custo"), drcontas("Oficial"), drcontas("Moeda"), sequencia)

                If sql.Length > 0 Then
                    arrsql.Add(sql)
                End If
            End If

            If oficial <> 0 Then
                sequencia += 1
                sql = sqlLancamentoEstoque2(drcontas("Empresa_Id"), drcontas("EndEmpresa_Id"), Left(drcontas("Conta_Id"), 7) & "01", drcontas("Produto"), drcontas("custo"), drcontas("Oficial"), drcontas("Moeda"), sequencia)
                arrsql.Add(sql)
            End If
        Next

        If arrsql.Count > 0 Then
            If Not Banco.GravaBanco(arrsql) Then
                MsgBox(Me.Page, "O processo não foi executado, tente novamente.")
            End If
        End If
    End Sub

    Private Function sqlLancamentoEstoque(ByVal Empresa_Id As String, ByVal EndEmpresa_Id As String, ByVal Conta_Id As String, ByVal Produto_Id As String, ByVal CentroDeCustos As String, ByVal Oficial As Decimal, ByVal Moeda As Decimal, ByVal Sequencia_Id As Integer) As String
        Dim sql As String
        Dim debitooficial, debitomoeda, creditooficial, creditomoeda As Decimal

        debitooficial = 0
        debitomoeda = 0
        creditooficial = 0
        creditomoeda = 0

        If Moeda < 0 Then
            debitomoeda = Math.Abs(Moeda)
        Else
            creditomoeda = Moeda
        End If
        If Oficial < 0 Then
            debitooficial = Math.Abs(Oficial)
        Else
            creditooficial = Oficial
        End If

        sql = "  INSERT INTO Razao " & vbCrLf & _
              " (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
              " Conta_Id, " & vbCrLf & _
              " Cliente_Id, EndCliente_Id, " & vbCrLf & _
              " Movimento_Id, " & vbCrLf & _
              " Lote_Id, Sequencia_Id, " & vbCrLf & _
              " Produto, " & vbCrLf & _
              " Custo, " & vbCrLf & _
              " Indexador, " & vbCrLf & _
              " DataMoeda, " & vbCrLf & _
              " DebitoOficial, DebitoMoeda, " & vbCrLf & _
              " CreditoOficial, CreditoMoeda, " & vbCrLf & _
              " Historico, " & vbCrLf & _
              " PrevistoRealizado)" & vbCrLf & _
              " VALUES " & vbCrLf & _
              " ('" & Empresa_Id & "', " & EndEmpresa_Id & vbCrLf & _
              ",'" & Conta_Id & "'" & vbCrLf & _
              ",'', 0" & vbCrLf & _
              ",'" & CDate(txtDataFim.Text).Year + 1 & "/01/01'" & vbCrLf & _
              ",7600, " & Sequencia_Id & vbCrLf & _
              ",'" & Produto_Id & "'" & vbCrLf & _
              ",'" & CentroDeCustos & "'" & vbCrLf & _
              ",3" & vbCrLf & _
              ",'" & CDate(txtDataFim.Text).Year + 1 & "/01/01'" & vbCrLf & _
              "," & debitooficial.ToString.Replace(".", "").Replace(",", ".") & ", " & debitomoeda.ToString.Replace(".", "").Replace(",", ".") & vbCrLf & _
              "," & creditooficial.ToString.Replace(".", "").Replace(",", ".") & ", " & creditomoeda.ToString.Replace(".", "").Replace(",", ".") & vbCrLf & _
              ",'SALDO INICIAL APURADO NO PERIODO DE " & txtDataInicio.Text & " A " & txtDataFim.Text & "'" & vbCrLf & _
              ",'D')" & vbCrLf

        Return sql
    End Function

    Private Function sqlLancamentoEstoque2(ByVal Empresa_Id As String, ByVal EndEmpresa_Id As String, ByVal Conta_Id As String, ByVal Produto_Id As String, ByVal CentroDeCustos As String, ByVal Oficial As Decimal, ByVal Moeda As Decimal, ByVal Sequencia_Id As Integer) As String
        Dim sql As String
        Dim debitooficial, debitomoeda, creditooficial, creditomoeda As Decimal

        debitooficial = 0
        debitomoeda = 0
        creditooficial = 0
        creditomoeda = 0

        If Moeda > 0 Then
            debitomoeda = Math.Abs(Moeda)
        Else
            creditomoeda = Math.Abs(Moeda)
        End If
        If Oficial > 0 Then
            debitooficial = Math.Abs(Oficial)
        Else
            creditooficial = Math.Abs(Oficial)
        End If

        sql = "  INSERT INTO Razao " & vbCrLf & _
              " (Empresa_Id, EndEmpresa_Id, " & vbCrLf & _
              " Conta_Id, " & vbCrLf & _
              " Cliente_Id, EndCliente_Id, " & vbCrLf & _
              " Movimento_Id, " & vbCrLf & _
              " Lote_Id, Sequencia_Id, " & vbCrLf & _
              " Produto, " & vbCrLf & _
              " Custo, " & vbCrLf & _
              " Indexador, " & vbCrLf & _
              " DataMoeda, " & vbCrLf & _
              " DebitoOficial, DebitoMoeda, " & vbCrLf & _
              " CreditoOficial, CreditoMoeda, " & vbCrLf & _
              " Historico, " & vbCrLf & _
              " PrevistoRealizado)" & vbCrLf & _
              " VALUES " & vbCrLf & _
              " ('" & Empresa_Id & "', " & EndEmpresa_Id & vbCrLf & _
              ", '" & Conta_Id & "'" & vbCrLf & _
              ", '', 0" & vbCrLf & _
              ", '" & CDate(txtDataFim.Text).Year + 1 & "/01/01'" & vbCrLf & _
              ", 7600, " & Sequencia_Id & vbCrLf & _
              ", '" & Produto_Id & "'" & vbCrLf & _
              ", '" & CentroDeCustos & "'" & vbCrLf & _
              ", 3" & vbCrLf & _
              ", '" & CDate(txtDataFim.Text).Year + 1 & "/01/01'" & vbCrLf & _
              ", " & debitooficial.ToString.Replace(".", "").Replace(",", ".") & ", " & debitomoeda.ToString.Replace(".", "").Replace(",", ".") & vbCrLf & _
              ", " & creditooficial.ToString.Replace(".", "").Replace(",", ".") & ", " & creditomoeda.ToString.Replace(".", "").Replace(",", ".") & vbCrLf & _
              ",'SALDO INICIAL APURADO NO PERIODO DE " & txtDataInicio.Text & " A " & txtDataFim.Text & "'" & vbCrLf & _
              ", 'D')" & vbCrLf

        Return sql
    End Function

#End Region
    Protected Sub lnkExcluir_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("ZeraContasdeResultado", "EXCLUIR") Then
                If ValidaCampos() Then
                    If chkContasDeEstoque.Checked Then
                        Dim Ano As Integer = CDate(txtDataFim.Text).Year

                        ExcluirLoteEstoque(Ano)
                        MsgBox(Me.Page, "Contas de estoque excluídas")
                    Else
                        ExcluirLote()
                        MsgBox(Me.Page, "Contas de Resultado excluídas")
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ZeraContasdeResultado")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class