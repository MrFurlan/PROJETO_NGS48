Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Partial Class NotaFiscalXClassificacao
    Inherits BasePage

    Dim sql As String = String.Empty
    Dim objNotaFiscal As [Lib].Negocio.NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack AndAlso IsConnect Then
            If Not Request.QueryString("produto") Is Nothing Then txtProduto.Value = Request.QueryString("produto")
            If Not Request.QueryString("entsai") Is Nothing Then txtEntradaSaida.Value = Request.QueryString("entsai")
            If Not Request.QueryString("HID") Is Nothing Then HID.Value = Request.QueryString("HID")

            If txtEntradaSaida.Value.Length > 0 Then lblEntSai.Text = txtEntradaSaida.Value.ToUpper()

            If Not Session("objNotaFiscalxOP" & Session.SessionID) Is Nothing Then
                RecuperaNotaFiscal()
            Else
                SessaoRecuperaNotaFiscal()
            End If

            DdlClassificacao.Items.Add(New ListItem(objNotaFiscal.Pedido.Itens(0).Classificacao.Descricao, objNotaFiscal.Pedido.Itens(0).Classificacao.Codigo))
            DdlClassificacao.SelectedIndex = 0
            Carregar_Classificacao(txtProduto.Value)
            LimparCampos()
        End If
    End Sub

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" + HID.Value), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Sub SessaoSalvaNotaFiscal()
        Session("objNotaFiscal" + HID.Value) = objNotaFiscal
    End Sub

    Private Sub SalvaNotaFiscal()
        Session("objNotaFiscalxOP" & Session.SessionID) = objNotaFiscal
    End Sub

    Private Sub RecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscalxOP" & Session.SessionID), [Lib].Negocio.NotaFiscal)
    End Sub

    Protected Sub cmdCalcular_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If txtPrimeiraPesagem.Text.Length = 0 Then
            MsgBox(Me.Page, "Pesagem deve ser maior que zero")
        ElseIf txtProduto.Value.Length = 0 Then
            MsgBox(Me.Page, "Produto para Classificação não foi selecionado")
        ElseIf txtEntradaSaida.Value.Length = 0 Then
            MsgBox(Me.Page, "Operação para Classificação não foi selecionado")
        Else
            HttpContext.Current.Session("ssCalcular") = "S"

            If GridDescontos.Rows.Count > 0 Then
                ConsultaClassificacao(txtProduto.Value)
            Else
                txtPesoBruto.Text = txtPrimeiraPesagem.Text
                txtDesconto.Text = 0
                txtLiquido.Text = txtPrimeiraPesagem.Text
            End If
        End If
    End Sub

    Private Sub Carregar_Classificacao(ByVal Produto As String)
        Dim ds As New DataSet
        Dim i As Integer

        sql = "SELECT Analises.Analise_Id as Codigo, Analises.Descricao, 0.00 as Percentual, 0 as Indice, 0 as Desconto " & vbCrLf & _
                  "FROM Classificacoes " & vbCrLf & _
                  "    INNER JOIN Analises " & vbCrLf & _
                  "        ON Classificacoes.Analise_Id = Analises.Analise_Id " & vbCrLf & _
                  " WHERE Classificacoes.Tabela_Id = " & DdlClassificacao.SelectedValue & vbCrLf & _
                  "   AND Classificacoes.Produto_Id = '" & Produto & "'" & vbCrLf & _
                  "   AND Classificacoes.Sequencia_Id = 1 " & vbCrLf & _
                  "   AND Analises.Analise_Id < 100 " & vbCrLf & _
                  " ORDER BY Analises.Analise_Id"

        ds = Banco.ConsultaDataSet(sql, "Consulta")

        GridDescontos.DataSource = ds
        GridDescontos.DataBind()

        For Each Dr As DataRow In ds.Tables(0).Rows
            If GridDescontos.Rows(i).Cells(0).Text = 11 Then
                CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = "0,25"
                CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = "0"
                CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = "0"

                CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = False
            Else
                CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = Dr("Percentual")
                CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = Dr("Indice")
                CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = Dr("Desconto")

                CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Enabled = True
            End If
            CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Enabled = False
            CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Enabled = False
            i += 1
        Next
    End Sub

    Sub ConsultaClassificacao(ByVal Produto As String)
        Dim ds As New DataSet
        Dim i As Integer
        Dim tolerancia, indice, decIndice As Decimal
        Dim Analise As Integer
        Dim Desconto As Integer = 0
        Dim Percentual As String
        Dim dra As DataRow

        txtPesoBruto.Text = CInt(txtPrimeiraPesagem.Text)
        txtLiquido.Text = CInt(txtPrimeiraPesagem.Text)

        txtDesconto.Text = "0"

        For i = 0 To GridDescontos.Rows.Count - 1
            Analise = GridDescontos.Rows(i).Cells(0).Text()

            If CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text <> "" And Analise <> 6 Then
                Percentual = CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
                Percentual = Percentual.Replace(",", ".")

                If Analise = "1" Then
                    sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, 0.00 As IndiceDesconto, 0 AS Desconto, FaixaValida "
                    sql &= "FROM Classificacoes "
                    sql &= "Where Tabela_Id = " & DdlClassificacao.SelectedValue & " And Produto_Id = '" & Produto & "' And Analise_Id = " & Analise & " "
                    sql &= "And " & Percentual & " >= FaixaInicial and FaixaValida = 'S' Order By Analise_Id"
                ElseIf Analise = "11" Then
                    sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, "
                    sql &= "Indice As IndiceDesconto, "
                    sql &= "round(Indice * " & txtPesoBruto.Text & " / 100,0) AS Desconto, FaixaValida "
                    sql &= "FROM Classificacoes "
                    sql &= "WHERE (Tabela_Id = " & DdlClassificacao.SelectedValue & ") AND (Produto_Id = '" & Produto & "') AND (Analise_Id = " & Analise & ") "
                    sql &= "AND " & Percentual & " BETWEEN FaixaInicial AND FaixaFinal Order By Analise_Id"
                Else
                    sql = "SELECT FaixaInicial, FaixaFinal, Indice, Tolerancia, "
                    sql &= "Indice * (" & Percentual & " - Tolerancia) As IndiceDesconto, "
                    sql &= "round(Indice * (" & Percentual & " - Tolerancia) * " & txtPesoBruto.Text & " / 100,0) AS Desconto, FaixaValida "
                    sql &= "FROM Classificacoes "
                    sql &= "WHERE (Tabela_Id = " & DdlClassificacao.SelectedValue & ") AND (Produto_Id = '" & Produto & "') AND (Analise_Id = " & Analise & ") "
                    sql &= "AND " & Percentual & " BETWEEN FaixaInicial AND FaixaFinal Order By Analise_Id"
                End If

                ds = Banco.ConsultaDataSet(sql, "Consulta")

                If lblEntSai.Text = "ENTRADA" And ds.Tables(0).Rows.Count = 0 Then
                    HttpContext.Current.Session("ssCalcular") = "N"
                    MsgBox(Me.Page, "Classificação informada não foi encontrada, verifique se informou corretamente ou entre em contato com o Suporte")
                    Exit Sub
                ElseIf lblEntSai.Text = "ENTRADA" And Not ds.Tables(0).Rows(0).Item("FaixaValida") = "S" Then
                    HttpContext.Current.Session("ssCalcular") = "N"
                    MsgBox(Me.Page, "Faixa digitada da Análise " & Analise & " é inválida")
                    Exit Sub
                End If

                If Analise = "1" Then
                    Percentual = Percentual.Replace(".", ",")
                    decIndice = 0
                    Dim j As Integer = 0

                    For Each dra In ds.Tables(0).Rows
                        tolerancia = FormatNumber(dra("FaixaInicial"), 2)
                        indice = dra("Indice")

                        If Percentual > dra("FaixaFinal") Then
                            decIndice = decIndice + (FormatNumber(dra("FaixaFinal"), 2) - FormatNumber(dra("FaixaInicial"), 2)) * indice
                        Else
                            'decIndice = decIndice + (Percentual - FormatNumber(dr("FaixaInicial"), 2)) * indice
                            'Arrumei a Fórmula junto com o Roberto, não estava correta - Furlan - 08/01/2009
                            decIndice = (Percentual - FormatNumber(dra("Tolerancia"), 2)) * indice
                        End If

                        j += 1

                        If j > 1 Then
                            dra.Delete()
                        End If
                    Next

                    If j > 1 Then
                        ds.AcceptChanges()
                    End If

                    ds.Tables(0).Rows(0).Item("IndiceDesconto") = FormatNumber(decIndice, 2)
                    ds.Tables(0).Rows(0).Item("Desconto") = CStr(CInt((CDec(txtPesoBruto.Text) * (ds.Tables(0).Rows(0).Item("IndiceDesconto") / 100)) + 0.0000000000001))
                End If

                If lblEntSai.Text = "ENTRADA" Then
                    CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = Convert.ToDecimal(ds.Tables(0).Rows(0).Item("IndiceDesconto")).ToString("N2")
                    CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = CInt(ds.Tables(0).Rows(0).Item("Desconto"))
                    Desconto = Desconto + ds.Tables(0).Rows(0).Item("Desconto")
                Else
                    CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = FormatNumber(0, 2)
                    CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = 0
                End If
            End If
        Next i

        txtDesconto.Text = Desconto
        txtLiquido.Text = CInt(txtPesoBruto.Text) - Desconto

        For i = 0 To GridDescontos.Rows.Count - 1
            Analise = GridDescontos.Rows(i).Cells(0).Text()

            If Analise = 6 Then
                If CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = "" Then
                    Percentual = "1.00"
                    CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = "1,00"
                Else
                    Percentual = CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text
                    Percentual = Percentual.Replace(",", ".")

                    If Percentual = "0.00" Then
                        Percentual = "1.00"
                        CType(GridDescontos.Rows(i).FindControl("txtPercentual"), TextBox).Text = "1,00"
                    End If
                End If

                sql = "SELECT Indice AS IndiceDesconto, round(Indice * " & (CInt(txtPesoBruto.Text) - Desconto) & " / 100,0) AS Desconto, "
                sql &= "FaixaValida FROM Classificacoes WHERE (Tabela_Id = " & DdlClassificacao.SelectedValue & ") AND "
                sql &= "(Produto_Id = '" & Produto & "') AND (Analise_Id = " & Analise & ") AND (Sequencia_Id = " & Percentual & ") "
                sql &= "Order By Analise_Id"

                ds = Banco.ConsultaDataSet(sql, "Consulta")

                If lblEntSai.Text = "ENTRADA" Then
                    CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = Convert.ToDecimal(ds.Tables(0).Rows(0).Item("IndiceDesconto")).ToString("N2")
                    CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = CInt(ds.Tables(0).Rows(0).Item("Desconto"))
                Else
                    CType(GridDescontos.Rows(i).FindControl("txtIndice"), TextBox).Text = FormatNumber(0, 2)
                    CType(GridDescontos.Rows(i).FindControl("txtDesconto"), TextBox).Text = 0
                End If
            End If
        Next i
    End Sub

    Sub LimparCampos()
        HttpContext.Current.Session("ssCalcular") = "N"
        DdlClassificacao.SelectedIndex = 0
        txtPrimeiraPesagem.Text = ""
        txtPesoBruto.Text = ""
        txtDesconto.Text = ""
        txtLiquido.Text = ""
        txtPrimeiraPesagem.Focus()
    End Sub

    Protected Sub btnOK_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        If HttpContext.Current.Session("ssCalcular") = "N" Then
            MsgBox(Me.Page, "Classificação não foi informada")
        Else
            cmdCalcular_Click(Nothing, Nothing)

            If Not Session("objNotaFiscalxOP" & Session.SessionID) Is Nothing Then
                RecuperaNotaFiscal()
            Else
                SessaoRecuperaNotaFiscal()
            End If

            Dim n As New [Lib].Negocio.Numerador(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, 110)
            Dim Rom As New [Lib].Negocio.Romaneio(objNotaFiscal.CodigoEmpresa, objNotaFiscal.EnderecoEmpresa, objNotaFiscal.CodigoRomaneio)
            If objNotaFiscal.CodigoRomaneio > 0 Then
                Rom.Codigo = objNotaFiscal.CodigoRomaneio
                Rom.DescontosAnalises = New [Lib].Negocio.ListRomaneioXDesconto()
                Rom.DescontosAnalises.Parent = Rom
            Else
                Rom.Codigo = n.Sequencia + 1
                Rom.DescontosAnalises = New [Lib].Negocio.ListRomaneioXDesconto(Rom)
            End If
            Rom.PesoBruto = CDbl(txtPesoBruto.Text)
            Rom.Desconto = CDbl(txtDesconto.Text)
            Rom.PesoLiquido = CDbl(txtLiquido.Text)
            Rom.Observacoes = txtObservacao.Text
            Rom.Fisico = False

            Dim i As Integer
            For i = 0 To GridDescontos.Rows.Count - 1
                Dim des As New [Lib].Negocio.RomaneioXDesconto(Rom)
                des.CodigoAnalise = GridDescontos.Rows(i).Cells(0).Text
                des.Percentual = CType(GridDescontos.Rows(i).Cells(2).FindControl("txtPercentual"), TextBox).Text
                des.Indice = CType(GridDescontos.Rows(i).Cells(3).FindControl("txtIndice"), TextBox).Text
                des.Desconto = CType(GridDescontos.Rows(i).Cells(4).FindControl("txtDesconto"), TextBox).Text
                Rom.DescontosAnalises.Add(des)
            Next

            objNotaFiscal.Romaneio = Rom
            objNotaFiscal.CodigoRomaneio = n.Sequencia + 1

            If Not Session("objNotaFiscalxOP" & Session.SessionID) Is Nothing Then
                SalvaNotaFiscal()
            Else
                SessaoSalvaNotaFiscal()
            End If

            Session("Classificacao" & Request.QueryString("tipo")) = "S"

            ScriptManager.RegisterClientScriptBlock(Me, Me.UpdatePanel1.GetType(), Guid.NewGuid().ToString(), "fecharJanela();", True)
        End If
    End Sub

    Protected Sub btnLimpar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        LimparCampos()
    End Sub

End Class