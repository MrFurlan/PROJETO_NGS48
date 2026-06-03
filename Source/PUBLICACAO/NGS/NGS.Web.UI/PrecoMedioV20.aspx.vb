Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PrecoMedioV20
    Inherits BasePage

    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            If Not IsPostBack And IsConnect Then
                Me.setMenu(eModulo.Comercial)
                If Funcoes.VerificaPermissao("PrecoMedioV20", "ACESSAR") Then
                    BuscaGrupo()
                    CarregaEmpresa()
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscaProduto()
        Dim sql As String
        Dim Codigo As String
        Dim Descricao As String

        sql = "SELECT Produtos.Produto_Id as Codigo, Produtos.Descricao  as DescricaoProd, Produtos.Agrupar " & vbCrLf &
              " FROM Produtos " & vbCrLf &
              " INNER Join" & vbCrLf &
              " GruposDeEstoques ON Produtos.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf
        If lstGrupoProduto.SelectedValue <> Nothing Then
            sql &= " where GruposDeEstoques.Grupo_Id =" & lstGrupoProduto.SelectedValue & " "
        End If
        sql &= " Order by Produtos.Descricao asc "

        lstProduto.Items.Clear()

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "Produtos").Tables(0).Rows
            Codigo = Dr("Codigo")
            Descricao = Dr("DescricaoProd")
            lstProduto.Items.Add(New ListItem(Descricao, Codigo))

        Next

        lstProduto.Items.Insert(0, "")
        lstProduto.SelectedIndex = 0
    End Sub

    Private Sub BuscaGrupo()
        Dim sql As String
        Dim Codigo As String
        Dim Descricao As String

        'sql = "SELECT Grupo_Id as Codigo,Descricao  as DescricaoGrupo "
        'sql &= " FROM GruposDeEstoques  where len(Grupo_Id)=5 Order by Descricao asc "

        sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao" & vbCrLf &
              " FROM         GruposDeEstoques INNER JOIN" & vbCrLf &
              "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & vbCrLf &
              " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) and Produtos.Agrupar = 'N' order by GruposDeEstoques.Descricao" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(sql, "GruposDeEstoques").Tables(0).Rows
            Codigo = Dr("Grupo_Id")

            'Nome = Funcoes.AlinharEsquerda(Dr("DescricaoGrupo"), 50, ".")

            Descricao = Dr("Descricao")

            lstGrupoProduto.Items.Add(New ListItem(Descricao, Codigo))
        Next

        lstGrupoProduto.Items.Insert(0, "")
        lstGrupoProduto.SelectedIndex = 0
    End Sub

    Private Sub BuscaSafra()
        Dim sql As String

        sql = "SELECT Safra_Id  " & vbCrLf &
              " FROM Safras " & vbCrLf

        lstProduto.DataSource = Banco.ConsultaDataSet(sql, "Safras")
        lstProduto.DataTextField = "Safra_Id"
        lstProduto.DataValueField = "Safra_Id"
        lstProduto.DataBind()
        lstProduto.Items.Insert(0, "")
        lstProduto.SelectedIndex = 0
    End Sub

    Public Sub CarregaEmpresa()
        ddl.Carregar(lstEmpresa, CarregarDDL.Tabela.Empresas, "")
        Funcoes.VerificaEmpresa(lstEmpresa)
    End Sub

    Protected Sub lstGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstGrupoProduto.SelectedIndexChanged
        Try
            BuscaProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub viewReport1(ByVal dsSoma As DataSet, ByVal Moeda As String)
        Dim crpt As New ReportDocument()

        Try
            crpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/CrPrecoMedioGraneis.rpt")
            crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
            crpt.SetDataSource(dsSoma)

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As ParameterFieldDefinitions
            Dim crparameterfielddefinition As ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssCidadeEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Estado")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssEstadoEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Preço Medio de Graneis em " & Moeda & "  - Periodo de " & txtDataInicio.Text & " a " & txtDataFim.Text
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Relatorio")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Graneis"
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("SubRelatorio")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Graneis"
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Private Sub viewReport2(ByVal dsOperacao As DataSet, ByVal Moeda As String)
        Dim crpt As New ReportDocument()

        Try
            crpt.FileName = HttpContext.Current.Server.MapPath("~/Reports/CrPrecoMedioGraneisOperacao.rpt")
            crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)

            crpt.SetDataSource(dsOperacao)

            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As ParameterFieldDefinitions
            Dim crparameterfielddefinition As ParameterFieldDefinition

            crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssCidadeEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Estado")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssEstadoEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Preço Medio de Graneis em " & Moeda & " - Periodo de " & txtDataInicio.Text & " a " & txtDataFim.Text
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Relatorio")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Graneis"
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Private Sub Limpar()
        txtDataInicio.Text = "01/01/" & Format(Now, "yyyy")
        txtDataFim.Text = Format(Now, "dd/MM/yyyy")
        lstGrupoProduto.SelectedIndex = "0"
        lstProduto.SelectedIndex = "0"
        DropOperacao.SelectedIndex = "1"
        dpRelPor.SelectedIndex = "0"
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            lstEmpresa.Enabled = False
        End If
    End Sub

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtDataInicio.Text) Then
            MsgBox(Me.Page, "Data não informada.")
            Return False
        ElseIf String.IsNullOrWhiteSpace(txtDataFim.Text) Then
            MsgBox(Me.Page, "Data não informada.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            'If Funcoes.VerificaPermissao("PrecoMedioV20", "RELATORIO") Then
            If ValidarCampos() Then
                Dim sql As String
                Dim dsCalculo As New DataSet
                Dim dsSomaAux As New DataSet
                Dim drCalculo As DataRow
                Dim drSoma As DataRow
                Dim NovaLinha As DataRow
                Dim NovaLinhaOperacao As DataRow
                Dim drCalculoAux() As DataRow
                Dim drSomaAux() As DataRow
                Dim ValorCompras As Double
                Dim PesoCompras As Integer
                Dim PrecoMedio As Double
                Dim PrecoMedioSaida As Double
                Dim strExpr As String
                Dim Empresa As String
                Dim Ciclo As Boolean = True
                Dim drAux As DataRow
                Dim int As Integer = 0
                Dim Moeda As String = String.Empty

                If RadioDolar.Checked Then
                    Moeda = "Dolar"
                ElseIf RadioReais.Checked Then
                    Moeda = "Reais"
                End If

                sql = " SELECT  Empresa_Id, EndEmpresa_Id, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf,Peso as PesoFrete, 0.00 AS Peso," & vbCrLf &
                      " (SELECT  case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else SUM(ValorOficial) end  AS ValorEncargos" & vbCrLf &
                      " FROM ConhecimentosDeTransportesXEncargos" & vbCrLf &
                      " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                      " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'OUTROSCREDITOS')) +" & vbCrLf &
                      " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else  SUM(ValorOficial) end AS ValorEncargos" & vbCrLf &
                      " FROM ConhecimentosDeTransportesXEncargos" & vbCrLf &
                      " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                      " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'ESTADIAS')) +" & vbCrLf &
                      " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else SUM(ValorOficial) end AS ValorEncargos" & vbCrLf &
                      " FROM ConhecimentosDeTransportesXEncargos" & vbCrLf &
                      " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                      " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'SOBRAS')) -" & vbCrLf &
                      " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else SUM(ValorOficial)  end AS ValorEncargos" & vbCrLf &
                      " FROM ConhecimentosDeTransportesXEncargos" & vbCrLf &
                      " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                      " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'QUEBRAS')) -" & vbCrLf &
                      " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else  SUM(ValorOficial) end AS ValorEncargos" & vbCrLf &
                      " FROM ConhecimentosDeTransportesXEncargos" & vbCrLf &
                      " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                      " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'OUTROSDESCONTOS ')) AS ValorEncargos, Conhecimento_Id, " & vbCrLf &
                      " 0.00 AS ValorMoeda, case when '" & Moeda & "'= 'Dolar' then ValorOficial / (select Indice from Cotacoes where Indexador_Id = 3 and Data_Id =Movimento) else  ValorOficial end AS VlrFrete, Serie_Nf, Numero_Nf" & vbCrLf &
                      " INTO [#TempAuxFretes]" & vbCrLf &
                      " FROM ConhecimentosDeTransportes" & vbCrLf &
                      " Where ConhecimentosDeTransportes.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'" & vbCrLf
                If lstEmpresa.SelectedValue <> Nothing Then
                    Dim strEmpresa As String()
                    strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                    sql &= " AND SubString(ConhecimentosDeTransportes.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

                End If
                sql &= ";"

                ''Busca Conhecimento de transporte com fixacoes
                sql &= " SELECT NotasFiscais.Empresa_Id," & vbCrLf &
                       "        NotasFiscais.EndEmpresa_Id," & vbCrLf &
                       "        Cliente_Nf," & vbCrLf &
                       "        EndCliente_Nf," & vbCrLf &
                       "        EntradaSaida_Nf," & vbCrLf &
                       "        PesoFrete," & vbCrLf &
                       "        Peso," & vbCrLf &
                       "        ISNULL(ValorEncargos, 0) AS ValorEncargos," & vbCrLf &
                       "        0.00 as ValorMoeda," & vbCrLf &
                       "        VlrFrete + ISNULL(ValorEncargos, 0) AS VlrFrete," & vbCrLf &
                       "        Serie_Nf," & vbCrLf &
                       "        Numero_Nf," & vbCrLf &
                       "        NotasFiscais.Operacao," & vbCrLf &
                       "        NotasFiscais.SubOperacao," & vbCrLf &
                       "        SubOperacoes.Classe," & vbCrLf &
                       "        SubOperacoes.Descricao AS DescricaoOperacao," & vbCrLf &
                       "       (SELECT TOP (1) Produto_Id" & vbCrLf &
                       "          FROM NotasFiscaisXItens" & vbCrLf &
                       "         WHERE NotasFiscais.Empresa_Id      = Empresa_Id" & vbCrLf &
                       "           AND NotasFiscais.EndEmpresa_Id   = EndEmpresa_Id" & vbCrLf &
                       "           AND NotasFiscais.Cliente_Id      = Cliente_Id" & vbCrLf &
                       "           AND NotasFiscais.EndCliente_Id   = EndCliente_Id" & vbCrLf &
                       "           AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id" & vbCrLf &
                       "           AND NotasFiscais.Serie_Id        = Serie_Id" & vbCrLf &
                       "           AND NotasFiscais.Nota_Id         = Nota_Id) AS Produto" & vbCrLf &
                       "  INTO [#TempAuxFretes1]" & vbCrLf &
                       "  FROM [#TempAuxFretes] " & vbCrLf &
                       " INNER JOIN NotasFiscais" & vbCrLf &
                       "    ON [#TempAuxFretes].Empresa_Id      = NotasFiscais.Empresa_Id" & vbCrLf &
                       "   AND [#TempAuxFretes].EndEmpresa_Id   = NotasFiscais.EndEmpresa_Id" & vbCrLf &
                       "   AND [#TempAuxFretes].Cliente_Nf      = NotasFiscais.Cliente_Id" & vbCrLf &
                       "   AND [#TempAuxFretes].EndCliente_Nf   = NotasFiscais.EndCliente_Id" & vbCrLf &
                       "   AND [#TempAuxFretes].EntradaSaida_Nf = NotasFiscais.EntradaSaida_Id" & vbCrLf &
                       "   AND [#TempAuxFretes].Serie_Nf        = NotasFiscais.Serie_Id" & vbCrLf &
                       "   AND [#TempAuxFretes].Numero_Nf       = NotasFiscais.Nota_Id " & vbCrLf &
                       " INNER JOIN SubOperacoes" & vbCrLf &
                       "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                       "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                       " INNER Join VW_PedidosXItensXFixacoes pfi" & vbCrLf &
                       "    ON NotasFiscais.Pedido = pfi.Pedido_Id" & vbCrLf &
                       " Where NotasFiscais.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'" & vbCrLf

                If lstEmpresa.SelectedValue <> Nothing Then
                    Dim strEmpresa As String()
                    strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                    sql &= " AND SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

                End If
                sql &= ";"

                ''Busca Conhecimento de transporte sem fixacoes apenas transferencias
                sql &= " insert into [#TempAuxFretes1] SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf,PesoFrete, Peso," & vbCrLf &
                       " ISNULL(ValorEncargos, 0) AS ValorEncargos,0.00 as ValorMoeda ,VlrFrete + ISNULL(ValorEncargos, 0) AS VlrFrete, Serie_Nf, Numero_Nf, " & vbCrLf &
                       " NotasFiscais.Operacao, NotasFiscais.SubOperacao, SubOperacoes.Classe, SubOperacoes.Descricao AS DescricaoOperacao," & vbCrLf &
                       " (SELECT  TOP (1) Produto_Id" & vbCrLf &
                       " FROM  NotasFiscaisXItens" & vbCrLf &
                       " WHERE  (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) " & vbCrLf &
                       " AND (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND " & vbCrLf &
                       " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) " & vbCrLf &
                       " AS Produto" & vbCrLf
                '' sql &= " INTO  [#TempAuxFretes1]"
                sql &= " FROM  [#TempAuxFretes] INNER JOIN" & vbCrLf &
                       " NotasFiscais ON [#TempAuxFretes].Empresa_Id = NotasFiscais.Empresa_Id AND " & vbCrLf &
                       " [#TempAuxFretes].EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND [#TempAuxFretes].Cliente_Nf = NotasFiscais.Cliente_Id AND " & vbCrLf &
                       " [#TempAuxFretes].EndCliente_Nf = NotasFiscais.EndCliente_Id AND [#TempAuxFretes].EntradaSaida_Nf = NotasFiscais.EntradaSaida_Id AND " & vbCrLf &
                       " [#TempAuxFretes].Serie_Nf = NotasFiscais.Serie_Id AND [#TempAuxFretes].Numero_Nf = NotasFiscais.Nota_Id INNER JOIN" & vbCrLf &
                       " SubOperacoes ON  NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf &
                       " NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf
                ' sql &= " INNER Join PedidosXItensXFixacoes ON NotasFiscais.Pedido = PedidosXItensXFixacoes.Pedido_Id"
                sql &= " Where SubOperacoes.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and NotasFiscais.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'"
                If lstEmpresa.SelectedValue <> Nothing Then
                    Dim strEmpresa As String()
                    strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                    sql &= " AND SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

                End If
                sql &= ";"

                ''Busca notas com fixaçoes 
                sql &= " SELECT NotasFiscais.Empresa_Id," & vbCrLf &
                       "        NotasFiscais.EndEmpresa_Id," & vbCrLf &
                       "        NotasFiscais.EntradaSaida_Id," & vbCrLf &
                       "        Sum(0.00) as PesoFrete," & vbCrLf &
                       "        NotasFiscais.Operacao," & vbCrLf &
                       "        NotasFiscais.SubOperacao," & vbCrLf &
                       "        SubOperacoes.Classe," & vbCrLf &
                       "        NotasFiscais.Cliente_Id," & vbCrLf &
                       "        NotasFiscais.EndCliente_Id," & vbCrLf &
                       "       (SELECT TOP (1) Produto_Id" & vbCrLf &
                       "          FROM NotasFiscaisXItens" & vbCrLf &
                       "         WHERE NotasFiscais.Empresa_Id      = Empresa_Id" & vbCrLf &
                       "           AND NotasFiscais.EndEmpresa_Id   = EndEmpresa_Id" & vbCrLf &
                       "           AND NotasFiscais.Cliente_Id      = Cliente_Id" & vbCrLf &
                       "           AND NotasFiscais.EndCliente_Id   = EndCliente_Id" & vbCrLf &
                       "           AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id" & vbCrLf &
                       "           AND NotasFiscais.Serie_Id        = Serie_Id" & vbCrLf &
                       "           AND NotasFiscais.Nota_Id         = Nota_Id) AS Produto," & vbCrLf &
                       "       (SELECT TOP (1) PesoFiscal" & vbCrLf &
                       "          FROM NotasFiscaisXItens" & vbCrLf &
                       "         WHERE NotasFiscais.Empresa_Id      = Empresa_Id" & vbCrLf &
                       "           AND NotasFiscais.EndEmpresa_Id   = EndEmpresa_Id" & vbCrLf &
                       "           AND NotasFiscais.Cliente_Id      = Cliente_Id" & vbCrLf &
                       "           AND NotasFiscais.EndCliente_Id   = EndCliente_Id" & vbCrLf &
                       "           AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id" & vbCrLf &
                       "           AND NotasFiscais.Serie_Id        = Serie_Id" & vbCrLf &
                       "           AND NotasFiscais.Nota_Id         = Nota_Id) AS Peso," & vbCrLf &
                       "       Case when '" & Moeda & "'= 'Dolar' " & vbCrLf &
                       "              then ((Sum(pif.TotalMoeda) / Sum(pif.Quantidade))" & vbCrLf &
                       "                   * (SELECT TOP (1) PesoFiscal" & vbCrLf &
                       "                        FROM NotasFiscaisXItens" & vbCrLf &
                       "                       WHERE NotasFiscais.Empresa_Id      = Empresa_Id" & vbCrLf &
                       "                         AND NotasFiscais.EndEmpresa_Id   = EndEmpresa_Id" & vbCrLf &
                       "                         AND NotasFiscais.Cliente_Id      = Cliente_Id" & vbCrLf &
                       "                         AND NotasFiscais.EndCliente_Id   = EndCliente_Id" & vbCrLf &
                       "                         AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id" & vbCrLf &
                       "                         AND NotasFiscais.Serie_Id        = Serie_Id" & vbCrLf &
                       "                         AND NotasFiscais.Nota_Id         = Nota_Id)) " & vbCrLf &
                       "              else ((Sum(pif.TotalOficial) / Sum(pif.Quantidade)) " & vbCrLf &
                       "                   * (SELECT TOP (1) PesoFiscal" & vbCrLf &
                       "                        FROM NotasFiscaisXItens" & vbCrLf &
                       "                       WHERE NotasFiscais.Empresa_Id      = Empresa_Id" & vbCrLf &
                       "                         AND NotasFiscais.EndEmpresa_Id   = EndEmpresa_Id" & vbCrLf &
                       "                         AND NotasFiscais.Cliente_Id      = Cliente_Id" & vbCrLf &
                       "                         AND NotasFiscais.EndCliente_Id   = EndCliente_Id" & vbCrLf &
                       "                         AND NotasFiscais.EntradaSaida_Id = EntradaSaida_Id" & vbCrLf &
                       "                         AND NotasFiscais.Serie_Id        = Serie_Id" & vbCrLf &
                       "                         AND NotasFiscais.Nota_Id         = Nota_Id)) " & vbCrLf &
                       "        end  as Valor," & vbCrLf &
                       "        0.00 AS VlrFrete," & vbCrLf &
                       "        SubOperacoes.Descricao AS DescricaoOperacao" & vbCrLf &
                       "  INTO [#TempNotasFiscais]" & vbCrLf &
                       "  FROM NotasFiscais" & vbCrLf &
                       " INNER JOIN SubOperacoes" & vbCrLf &
                       "    ON NotasFiscais.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                       "   AND NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                       " INNER JOIN NotasFiscaisXItens" & vbCrLf &
                       "    ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
                       "   AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
                       "   And NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
                       "   And NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
                       "   And NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
                       "   And NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
                       "   AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                       " INNER Join VW_PedidosXItensXFixacoes pif" & vbCrLf &
                       "    ON NotasFiscais.Pedido = pif.Pedido_Id" & vbCrLf &
                       "   AND pif.quantidade > 0 " & vbCrLf

                If lstEmpresa.SelectedValue <> Nothing Then
                    Dim strEmpresa As String()
                    strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                    sql &= " Where SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' and (SubOperacoes.Classe <> '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "') and  (SubOperacoes.Classe <> '" & eClassesOperacoes.REMESSAS.ToString & "') "

                End If
                sql &= " group by NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Operacao, NotasFiscais.SubOperacao, "
                sql &= " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,NotasFiscais.Serie_Id,NotasFiscais.Nota_Id,SubOperacoes.Descricao;"

                ''Busca notas sem fixaçoes apenas transferencia
                sql &= " insert into [#TempNotasFiscais] SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id,Sum(0.00) as PesoFrete, NotasFiscais.Operacao, NotasFiscais.SubOperacao," & vbCrLf &
                       " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id," & vbCrLf &
                       "  (SELECT TOP (1) Produto_Id" & vbCrLf &
                       "  FROM NotasFiscaisXItens" & vbCrLf &
                       "  WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                       "  (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND " & vbCrLf &
                       "  (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) " & vbCrLf &
                       "  AS Produto," & vbCrLf &
                       " (SELECT     TOP (1) PesoFiscal" & vbCrLf &
                       " FROM NotasFiscaisXItens" & vbCrLf &
                       " WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND " & vbCrLf &
                       " (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND " & vbCrLf &
                       " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) " & vbCrLf &
                       " AS Peso, case when '" & Moeda & "'= 'Dolar' then Sum(NotasFiscaisXItens.Valor) else Sum(NotasFiscaisXItens.Valor) end  as Valor, 0.00 AS VlrFrete, SubOperacoes.Descricao AS DescricaoOperacao" & vbCrLf
                'sql &= " INTO [#TempNotasFiscais]"
                sql &= " FROM NotasFiscais INNER JOIN" & vbCrLf &
                       " SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND " & vbCrLf &
                       " NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN" & vbCrLf &
                       " NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND " & vbCrLf &
                       " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id And NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id And " & vbCrLf &
                       " NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id And NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id And " & vbCrLf &
                       " NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id" & vbCrLf &
                       " where SubOperacoes.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'" & vbCrLf
                If lstEmpresa.SelectedValue <> Nothing Then
                    Dim strEmpresa As String()
                    strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                    sql &= " and SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

                End If
                sql &= " group by NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Operacao, NotasFiscais.SubOperacao, "
                sql &= " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,NotasFiscais.Serie_Id,NotasFiscais.Nota_Id,SubOperacoes.Descricao;"

                ''Busca fixacoes
                sql &= " SELECT pif.Empresa_Id," & vbCrLf &
                       "        pif.EndEmpresa_Id," & vbCrLf &
                       "        pif.Produto_Id," & vbCrLf &
                       "        SubOperacoes.EntradaSaida,0.00 as PesoFrete," & vbCrLf &
                       "        pif.Operacao," & vbCrLf &
                       "        pif.SubOperacao," & vbCrLf &
                       "        SubOperacoes.Classe," & vbCrLf &
                       "        Pedidos.Cliente," & vbCrLf &
                       "        Pedidos.EndCliente, " & vbCrLf &
                       "        Produtos.Nome AS NomeProduto," & vbCrLf &
                       "        pif.Quantidade AS Peso," & vbCrLf &
                       "        case when '" & Moeda & "'= 'Dolar' then  pif.TotalMoeda else pif.TotalOficial end AS Valor, " & vbCrLf &
                       "        isnull((SELECT case when '" & Moeda & "'= 'Dolar' then SUM(pife.ValorMoeda) else SUM(pife.ValorOficial) end AS ValorMoedaEncargos" & vbCrLf &
                       "                  FROM VW_PedidosXItensXFixacoesXEncargos pife " & vbCrLf &
                       "                 WHERE pife.Empresa_Id    = pif.Empresa_Id" & vbCrLf &
                       "                   AND pife.EndEmpresa_Id = pif.EndEmpresa_Id" & vbCrLf &
                       "                   AND pife.Pedido_Id     = pif.Pedido_Id" & vbCrLf &
                       "                   AND pife.Produto_Id    = pif.Produto_Id" & vbCrLf &
                       "                   AND pife.Fixacao_Id    = pif.Fixacao_Id" & vbCrLf &
                       "                   AND pife.Encargo_Id    = 'Frete'),0) as VlrFrete, " & vbCrLf &
                       "        Clientes.Nome AS NomeCliente, " & vbCrLf &
                       "        Clientes.Cidade," & vbCrLf &
                       "        Clientes.Estado," & vbCrLf &
                       "        Produtos.Grupo," & vbCrLf &
                       "        SubOperacoes.Descricao AS NomeOperacao, " & vbCrLf &
                       "        ISNULL((SELECT case when '" & Moeda & "'= 'Dolar' then SUM(pife.ValorMoeda) else SUM(pife.ValorOficial) end AS ValorMoedaEncargos" & vbCrLf &
                       "                  FROM VW_PedidosXItensXFixacoesXEncargos pife" & vbCrLf &
                       "                 WHERE pife.Empresa_Id    = pif.Empresa_Id" & vbCrLf &
                       "                   AND pife.EndEmpresa_Id = pif.EndEmpresa_Id" & vbCrLf &
                       "                   AND pife.Pedido_Id     = pif.Pedido_Id" & vbCrLf &
                       "                   AND pife.Produto_Id    = pif.Produto_Id" & vbCrLf &
                       "                   AND pife.Fixacao_Id    = pif.Fixacao_Id" & vbCrLf &
                       "                   AND pife.Encargo_Id    = 'PIS'),0)" & vbCrLf &
                       "        + ISNULL((SELECT case when '" & Moeda & "'= 'Dolar' then SUM(pife.ValorMoeda) else SUM(pife.ValorOficial) end AS ValorMoedaEncargos" & vbCrLf &
                       "                    FROM VW_PedidosXItensXFixacoesXEncargos pife" & vbCrLf &
                       "                   WHERE pife.Empresa_Id    = pif.Empresa_Id" & vbCrLf &
                       "                     AND pife.EndEmpresa_Id = pif.EndEmpresa_Id" & vbCrLf &
                       "                     AND pife.Pedido_Id     = pif.Pedido_Id" & vbCrLf &
                       "                     AND pife.Produto_Id    = pif.Produto_Id" & vbCrLf &
                       "                     AND pife.Fixacao_Id    = pif.Fixacao_Id" & vbCrLf &
                       "                     AND (pife.Encargo_Id    = 'FUNRURAL' or pife.Encargo_Id    = 'FUNRURAL JUDICIAL' or pife.Encargo_Id    = 'SENAR')),0)" & vbCrLf &
                       "        + ISNULL((SELECT case when '" & Moeda & "'= 'Dolar' then SUM(pife.ValorMoeda) else SUM(pife.ValorOficial) end AS ValorMoedaEncargos" & vbCrLf &
                       "                    FROM VW_PedidosXItensXFixacoesXEncargos pife" & vbCrLf &
                       "                   WHERE pife.Empresa_Id    = pif.Empresa_Id" & vbCrLf &
                       "                     AND pife.EndEmpresa_Id = pif.EndEmpresa_Id" & vbCrLf &
                       "                     AND pife.Pedido_Id     = pif.Pedido_Id" & vbCrLf &
                       "                     AND pife.Produto_Id    = pif.Produto_Id" & vbCrLf &
                       "                     AND pife.Fixacao_Id    = pif.Fixacao_Id" & vbCrLf &
                       "                     AND pife.Encargo_Id    = 'FETHAB'),0) AS ValorMoedaEncargos" & vbCrLf &
                       "   INTO [#TempFixacaoDePreco]" & vbCrLf &
                       "   FROM VW_PedidosXItensXFixacoes pif" & vbCrLf &
                       "   INNER JOIN Pedidos" & vbCrLf &
                       "      ON pif.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf &
                       "     AND pif.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf &
                       "     AND pif.Pedido_Id     = Pedidos.Pedido_Id" & vbCrLf &
                       "   INNER JOIN SubOperacoes" & vbCrLf &
                       "      ON pif.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                       "     AND pif.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf &
                       "   INNER JOIN Produtos" & vbCrLf &
                       "      ON pif.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                       "   INNER JOIN Clientes" & vbCrLf &
                       "      ON Pedidos.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf &
                       "     AND Pedidos.EndEmpresa_Id = Clientes.Endereco_Id " & vbCrLf &
                       "   Where pif.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'" & vbCrLf

                If lstEmpresa.SelectedValue <> Nothing Then
                    Dim strEmpresa As String()
                    strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                    sql &= " AND SubString(pif.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

                End If
                sql &= ";"

                ''--Cria temp Final
                sql &= " SELECT   [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto AS Produto_Id, EntradaSaida_Nf AS EntradaSaida, Sum(PesoFrete) as PesoFrete," & vbCrLf &
                       " Operacao AS Operacao_Id, SubOperacao AS SubOperacoes_Id, Classe, Cliente_Nf AS Cliente_Id, EndCliente_Nf AS EndCliente_Id, Produtos.Nome, " & vbCrLf &
                       " SUM(0.00) AS Peso, SUM(ValorMoeda) AS Valor, SUM(VlrFrete) AS VlrFrete, Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado, " & vbCrLf &
                       " Produtos.Grupo, DescricaoOperacao AS DescSubOperacao, SUM(ValorEncargos) AS ValorEncargos,Empresa.Cidade as CidadeEmpresa" & vbCrLf &
                       " INTO [#TempFinal]" & vbCrLf &
                       " FROM [#TempAuxFretes1] INNER JOIN" & vbCrLf &
                       " Produtos ON [#TempAuxFretes1].Produto = Produtos.Produto_Id INNER JOIN" & vbCrLf &
                       " Clientes ON [#TempAuxFretes1].Empresa_Id = Clientes.Cliente_Id AND [#TempAuxFretes1].EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf &
                       " INNER JOIN Clientes as Empresa ON Cliente_Nf = Empresa.Cliente_Id AND EndCliente_Nf = Empresa.Endereco_Id  " & vbCrLf &
                       " WHERE Classe Is Not NULL and  [#TempAuxFretes1].Empresa_Id = '" & "-1" & "' " & vbCrLf &
                       " GROUP BY [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto, EntradaSaida_Nf, Operacao, " & vbCrLf &
                       " SubOperacao, Classe, Cliente_Nf, EndCliente_Nf, Produtos.Nome, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, " & vbCrLf &
                       " DescricaoOperacao, Empresa.Cidade " & vbCrLf


                If dpRelPor.SelectedValue = "Movimento" Then
                    ''--Buscas Temporarios Recibo Fretes
                    sql &= " insert Into #TempFinal SELECT  [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto AS Produto_Id, EntradaSaida_Nf AS EntradaSaida, Sum(PesoFrete) as PesoFrete ," & vbCrLf &
                           " Operacao AS Operacao_Id, SubOperacao AS SubOperacoes_Id, Classe, Cliente_Nf AS Cliente_Id, EndCliente_Nf AS EndCliente_Id, Produtos.Nome, " & vbCrLf &
                           " SUM(0.00) AS Peso, SUM(ValorMoeda) AS Valor, SUM(VlrFrete) AS VlrFrete, Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado, " & vbCrLf &
                           " Produtos.Grupo, DescricaoOperacao AS DescSubOperacao, SUM(ValorEncargos) AS ValorEncargos,Empresa.Cidade as CidadeEmpresa" & vbCrLf &
                           " FROM [#TempAuxFretes1] INNER JOIN" & vbCrLf &
                           " Produtos ON [#TempAuxFretes1].Produto = Produtos.Produto_Id INNER JOIN" & vbCrLf &
                           " Clientes ON [#TempAuxFretes1].Empresa_Id = Clientes.Cliente_Id AND [#TempAuxFretes1].EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf &
                           " INNER JOIN Clientes as Empresa ON Cliente_Nf = Empresa.Cliente_Id AND EndCliente_Nf = Empresa.Endereco_Id  " & vbCrLf &
                           " WHERE  (Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' or Classe = '" & eClassesOperacoes.VENDAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "' or Classe = '" & eClassesOperacoes.REAJUSTES.ToString & "' or Classe = '" & eClassesOperacoes.REMESSAS.ToString & "') " & vbCrLf &
                           " GROUP BY [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto, EntradaSaida_Nf, Operacao, " & vbCrLf &
                           " SubOperacao, Classe, Cliente_Nf, EndCliente_Nf, Produtos.Nome, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, " & vbCrLf &
                           " DescricaoOperacao, Empresa.Cidade" & vbCrLf

                    ''--Buscas Temporarios Notas Fiscais
                    sql &= " insert Into #TempFinal SELECT  Empresa_Id, EndEmpresa_id, Produto AS Produto_Id, EntradaSaida_Id,Sum(PesoFrete) as PesoFrete , Operacao AS Operacao_Id, " & vbCrLf &
                           " SubOperacao AS SubOperacoes_Id, Classe, [#TempNotasFiscais].Cliente_Id, EndCliente_Id, Produtos.Nome, SUM(ISNULL(Peso, 0)) AS Peso, " & vbCrLf &
                           " SUM(Valor) AS Valor, SUM(VlrFrete) AS VlrFrete, Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, " & vbCrLf &
                           " DescricaoOperacao AS DescSubOperacao, SUM(0.00) AS ValorEncargos,Empresa.Cidade as CidadeEmpresa " & vbCrLf &
                           " FROM [#TempNotasFiscais] INNER JOIN " & vbCrLf &
                           " Produtos ON [#TempNotasFiscais].Produto = Produtos.Produto_Id INNER JOIN " & vbCrLf &
                           " Clientes ON [#TempNotasFiscais].Empresa_Id = Clientes.Cliente_Id AND [#TempNotasFiscais].EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf &
                           " INNER JOIN Clientes as Empresa ON [#TempNotasFiscais].Cliente_Id = Empresa.Cliente_Id AND EndCliente_Id = Empresa.Endereco_Id " & vbCrLf &
                           " WHERE     (Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' or Classe = '" & eClassesOperacoes.VENDAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "' or Classe = '" & eClassesOperacoes.REAJUSTES.ToString & "' or Classe = '" & eClassesOperacoes.REMESSAS.ToString & "')" & vbCrLf &
                           " GROUP BY Empresa_Id, EndEmpresa_id, Produto, EntradaSaida_Id, Operacao, SubOperacao, Classe, [#TempNotasFiscais].Cliente_Id, EndCliente_Id, " & vbCrLf &
                           " Produtos.Nome, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, DescricaoOperacao, Empresa.Cidade " & vbCrLf
                End If

                If dpRelPor.SelectedValue = "Fixações" Then
                    ''--Buscas Temporarios Fixacao de Precos
                    sql &= " insert Into #TempFinal SELECT      Empresa_Id, EndEmpresa_Id, Produto_Id, EntradaSaida,Sum(PesoFrete) as PesoFrete , Operacao as Operacao_Id, SubOperacao AS SubOperacoes_Id, Classe, " & vbCrLf &
                           " Cliente AS Cliente_Id, EndCliente AS EndCliente_Id, NomeProduto AS Nome, SUM(Peso) AS Peso, SUM(Valor - ValorMoedaEncargos) AS Valor, " & vbCrLf &
                           " SUM(VlrFrete) AS VlrFrete, NomeCliente, [#TempFixacaoDePreco].Cidade, [#TempFixacaoDePreco].Estado, Grupo, NomeOperacao AS DescSubOperacao, SUM(ValorMoedaEncargos) " & vbCrLf &
                           " AS ValorEncargos,Empresa.Cidade as CidadeEmpresa " & vbCrLf &
                           " FROM [#TempFixacaoDePreco] " & vbCrLf &
                           " INNER JOIN Clientes as Empresa ON [#TempFixacaoDePreco].Cliente = Empresa.Cliente_Id AND [#TempFixacaoDePreco].EndCliente = Empresa.Endereco_Id " & vbCrLf &
                           " WHERE(Classe Is Not NULL) " & vbCrLf &
                           " GROUP BY Empresa_Id, EndEmpresa_Id, Produto_Id, EntradaSaida, Operacao, SubOperacao, Classe, Cliente, EndCliente, NomeProduto, NomeCliente, " & vbCrLf &
                           " [#TempFixacaoDePreco].Cidade, [#TempFixacaoDePreco].Estado, Grupo, NomeOperacao,Empresa.Cidade " & vbCrLf
                End If

                'sql &= " SELECT SaldoInicialPrecoMedio.Empresa_Id, SaldoInicialPrecoMedio.EndEmpresa_Id, case 'N' when 'S' then SaldoInicialPrecoMedio.Produto_Id else Produtos.Grupo_Id end as Produto_Id, "
                'sql &= " SaldoInicialPrecoMedio.EntradaSaida_Id as EntradaSaida, SaldoInicialPrecoMedio.Operacao_Id, SaldoInicialPrecoMedio.SubOperacao_Id as SubOperacoes_Id, SubOperacoes.Classe, "
                'sql &= " '' AS Cliente_Id, 0 AS EndCliente_Id, case 'N' when 'S' then Produtos.Nome else GruposDeEstoques.Descricao end as Nome, Clientes.Nome AS NomeCliente, Sum(SaldoInicialPrecoMedio.Quantidade) AS Peso, "
                'sql &= " Sum(SaldoInicialPrecoMedio.ValorMoeda) AS Valor, 0.00 AS VlrFrete,Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado_Id, Produtos.Grupo_Id,SubOperacoes.Descricao as DescSubOperacao, '' as CidadeEmpresa"
                'sql &= " FROM  SaldoInicialPrecoMedio INNER JOIN"
                'sql &= " SubOperacoes ON SaldoInicialPrecoMedio.Operacao_Id = SubOperacoes.Operacao_Id AND "
                'sql &= " SaldoInicialPrecoMedio.SubOperacao_Id = SubOperacoes.SubOperacoes_Id INNER JOIN"
                'sql &= " Produtos ON SaldoInicialPrecoMedio.Produto_Id = Produtos.Produto_Id INNER JOIN"
                'sql &= " Clientes ON SaldoInicialPrecoMedio.Empresa_Id = Clientes.Cliente_Id AND SaldoInicialPrecoMedio.EndEmpresa_Id = Clientes.Endereco_Id"
                'sql &= " INNER JOIN GruposDeEstoques ON  Produtos.Grupo_Id = GruposDeEstoques.Grupo_Id"
                ' ''-- Where SaldoInicialPrecoMedio.Movimento_Id BETWEEN '" & Format(CDate(txtInicio.Text), "yyyy/MM/dd") & "' AND '" & Format(CDate(txtFim.Text), "yyyy/MM/dd") & "' ""
                ' ''--If ProdutoSafra <> Nothing Then"
                ' ''--    sql &= " and SaldoInicialPrecoMedio.Produto_Id='" & ProdutoSafra & "' ""
                ' ''--End If"

                'sql &= " Group By SaldoInicialPrecoMedio.Empresa_Id, SaldoInicialPrecoMedio.EndEmpresa_Id, case 'N' when 'S' then SaldoInicialPrecoMedio.Produto_Id else Produtos.Grupo_Id end, "
                'sql &= " SaldoInicialPrecoMedio.EntradaSaida_Id, SaldoInicialPrecoMedio.Operacao_Id, """
                'sql &= " SaldoInicialPrecoMedio.SubOperacao_Id, SubOperacoes.Classe, case 'N' when 'S' then Produtos.Nome else GruposDeEstoques.Descricao end, "
                'sql &= " Clientes.Nome, Clientes.Cidade, Clientes.Estado_Id, Produtos.Grupo_Id, SubOperacoes.Descricao"

                'sql &= " Union"

                sql &= " select Empresa_Id,EndEmpresa_Id,case 'N' when 'S' then #TempFinal.Produto_Id else #TempFinal.Grupo end as Produto_Id,EntradaSaida,Sum(PesoFrete) as PesoFrete,Operacao_Id,SubOperacoes_Id,Classe,CASE #TempFinal.Classe WHEN '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' then Cliente_Id else '' end as Cliente_Id ,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' then  EndCliente_Id else 0 end as EndCliente_Id,case 'N' when 'S' then #TempFinal.Nome else GruposDeEstoques.Descricao end as Nome,NomeCliente,Sum(Peso) as Peso,isnull(Sum(Valor),0) as Valor,Sum(VlrFrete) as VlrFrete ,NomeCliente,Cidade,Estado,#TempFinal.Grupo,DescSubOperacao,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' THEN #TempFinal.CidadeEmpresa else '' end as CidadeEmpresa from #TempFinal" & vbCrLf &
                       " inner join Produtos On #TempFinal.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                       " INNER JOIN GruposDeEstoques ON  #TempFinal.Grupo = GruposDeEstoques.Grupo_Id" & vbCrLf

                sql &= " Where #TempFinal.Produto_Id is not null  "
                If lstProduto.SelectedValue <> Nothing Then
                    sql &= " and #TempFinal.Produto_Id='" & lstProduto.SelectedValue & "' "
                End If
                If lstGrupoProduto.SelectedValue <> Nothing Then
                    sql &= " and GruposDeEstoques.Grupo_Id='" & lstGrupoProduto.SelectedValue & "' "
                End If

                sql &= " group by Empresa_Id,EndEmpresa_Id,case 'N' when 'S' then #TempFinal.Produto_Id else #TempFinal.Grupo end,EntradaSaida,Operacao_Id,SubOperacoes_Id,Classe,CASE #TempFinal.Classe WHEN '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' then Cliente_Id else '' end  ,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' then  EndCliente_Id else 0 end,case 'N' when 'S' then #TempFinal.Nome else GruposDeEstoques.Descricao end,NomeCliente,NomeCliente,Cidade,Estado,#TempFinal.Grupo,DescSubOperacao,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' THEN #TempFinal.CidadeEmpresa else '' end" & vbCrLf &
                       " order by Empresa_Id,EndEmpresa_Id,Produto_Id,EntradaSaida,Operacao_Id,SubOperacoes_Id,Cliente_Id,EndCliente_Id" & vbCrLf
                dsCalculo = Banco.ConsultaDataSet(sql, "ComercialDeInsumos")

                dsCalculo.Tables(0).Columns.Add("ValEntTrans", GetType(Double))
                dsCalculo.Tables(0).Columns.Add("PesoEntTrans", GetType(Double))
                dsCalculo.Tables(0).Columns.Add("ValSaidaTrans", GetType(Double))
                dsCalculo.Tables(0).Columns.Add("PesoSaidaTrans", GetType(Double))

                For Each drAux In dsCalculo.Tables(0).Rows
                    drAux("ValEntTrans") = 0
                    drAux("PesoEntTrans") = 0
                    drAux("ValSaidaTrans") = 0
                    drAux("PesoSaidaTrans") = 0
                Next

                While Ciclo = True
                    Ciclo = False
                    Empresa = ""
                    For Each drCalculo In dsCalculo.Tables(0).Rows
                        If drCalculo("Peso") <> 0 Then
                            If Empresa <> drCalculo("Empresa_Id") & drCalculo("EndEmpresa_Id") & drCalculo("Produto_Id") Then
                                ValorCompras = 0
                                PesoCompras = 0
                                PrecoMedio = 0
                                PrecoMedioSaida = 0
                            End If
                            If drCalculo("EntradaSaida") = "E" Then
                                ValorCompras = ValorCompras + (drCalculo("Valor") + drCalculo("VlrFrete"))
                                PesoCompras = PesoCompras + drCalculo("Peso")

                            ElseIf drCalculo("EntradaSaida") = "S" And drCalculo("Classe") = "TRANSFERENCIAS" Then
                                PrecoMedio = FormatNumber(((ValorCompras / PesoCompras) * 1000), 2)
                                If PrecoMedio > 0 Then
                                    If FormatNumber(drCalculo("Valor"), 2) <> FormatNumber((PrecoMedio * drCalculo("Peso")) / 1000, 2) Then
                                        Ciclo = True
                                    End If
                                    drCalculo("Valor") = FormatNumber((PrecoMedio * drCalculo("Peso")) / 1000, 2)
                                    'drCalculo("PesoSaidaTrans") = drCalculo("Peso")
                                    PrecoMedioSaida = FormatNumber(((drCalculo("Valor") + drCalculo("VlrFrete")) / drCalculo("Peso") * 1000), 2)

                                    ''Filial destino
                                    strExpr = "EntradaSaida = 'E' and Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' AND Empresa_Id = '" & drCalculo("Cliente_Id") & "' And EndEmpresa_Id = " & drCalculo("EndCliente_Id") & " and Produto_Id = '" & drCalculo("Produto_Id") & "' AND Cliente_Id = '" & drCalculo("Empresa_Id") & "' And EndCliente_Id = " & drCalculo("EndEmpresa_Id") & ""
                                    drCalculoAux = dsCalculo.Tables(0).Select(strExpr)
                                    If drCalculoAux.Length > 0 And PrecoMedioSaida > 0 Then
                                        drCalculoAux(0)("Valor") = drCalculo("Valor")
                                        drCalculoAux(0)("Peso") = drCalculo("Peso")
                                        drCalculoAux(0)("PesoFrete") = drCalculo("PesoFrete")
                                        drCalculoAux(0)("VlrFrete") = drCalculo("VlrFrete")
                                    End If
                                End If
                            End If
                            Empresa = drCalculo("Empresa_Id") & drCalculo("EndEmpresa_Id") & drCalculo("Produto_Id")
                        End If
                    Next
                End While

                If DropOperacao.Text = "N" Then
                    Dim tbSoma As DataTable = New DataTable("ComercialDeInsumos")
                    Dim dsSoma As New DataSet

                    tbSoma.Columns.Add("Produto_Id", GetType(String))
                    tbSoma.Columns.Add("Nome", GetType(String))
                    tbSoma.Columns.Add("PesoCompras", GetType(Integer)).DefaultValue = 0
                    tbSoma.Columns.Add("ValorU$Compras", GetType(Decimal)).DefaultValue = 0
                    tbSoma.Columns.Add("PesoTransf", GetType(Integer)).DefaultValue = 0
                    tbSoma.Columns.Add("ValorU$Transf", GetType(Decimal)).DefaultValue = 0
                    tbSoma.Columns.Add("PesoVendas", GetType(Integer)).DefaultValue = 0
                    tbSoma.Columns.Add("ValorU$Vendas", GetType(Decimal)).DefaultValue = 0
                    tbSoma.Columns.Add("Empresa_Id", GetType(String))
                    tbSoma.Columns.Add("EndEmpresa_Id", GetType(String))
                    tbSoma.Columns.Add("ValSaidaTrans", GetType(Decimal)).DefaultValue = 0
                    tbSoma.Columns.Add("PesoSaidaTrans", GetType(Integer)).DefaultValue = 0
                    tbSoma.Columns.Add("ValEntTrans", GetType(Decimal)).DefaultValue = 0
                    tbSoma.Columns.Add("PesoEntTrans", GetType(Integer)).DefaultValue = 0
                    tbSoma.Columns.Add("NomeCliente", GetType(String))
                    tbSoma.Columns.Add("Estado_Id", GetType(String))
                    tbSoma.Columns.Add("Cidade", GetType(String))
                    tbSoma.Columns.Add("Grupo_Id", GetType(String))
                    tbSoma.Columns.Add("Descricao", GetType(String))
                    tbSoma.Columns.Add("VlrFreteCompras", GetType(Double))
                    tbSoma.Columns.Add("VlrFreteVendas", GetType(Double))
                    tbSoma.Columns.Add("VlrFreteSaiTrans", GetType(Double))
                    tbSoma.Columns.Add("VlrFreteEntTrans", GetType(Double))

                    dsSoma.Tables.Add(tbSoma)
                    dsSoma.AcceptChanges()

                    For Each drSoma In dsCalculo.Tables(0).Rows
                        strExpr = "Produto_Id = '" & drSoma("Produto_Id") & "' and Empresa_Id = '" & drSoma("Empresa_Id") & "' And EndEmpresa_Id = " & drSoma("EndEmpresa_Id") & ""
                        drSomaAux = dsSoma.Tables(0).Select(strExpr)
                        If drSomaAux.Length = 0 Then
                            NovaLinha = dsSoma.Tables(0).NewRow()
                            NovaLinha("Produto_Id") = drSoma("Produto_Id")
                            NovaLinha("Nome") = drSoma("Nome")
                            If drSoma("EntradaSaida") = "E" And (drSoma("Classe") = "COMPRAS" Or drSoma("Classe") = "COMPLEMENTACOES" Or drSoma("Classe") = "REAJUSTES" Or drSoma("Classe") = "REMESSAS" Or drSoma("Classe") = "INICIAL" Or drSoma("Classe") = "IMPORTACOES") Then
                                NovaLinha("PesoCompras") = drSoma("Peso")
                                NovaLinha("ValorU$Compras") = drSoma("Valor")
                                NovaLinha("VlrFreteCompras") = drSoma("VlrFrete")
                            Else
                                NovaLinha("PesoCompras") = 0
                                NovaLinha("ValorU$Compras") = 0
                                NovaLinha("VlrFreteCompras") = 0
                            End If

                            If drSoma("Classe") = "TRANSFERENCIAS" Then
                                NovaLinha("PesoTransf") = drSoma("Peso")
                                NovaLinha("ValorU$Transf") = drSoma("Valor")
                            Else
                                NovaLinha("PesoTransf") = 0
                                NovaLinha("ValorU$Transf") = 0
                            End If
                            If drSoma("EntradaSaida") = "S" And (drSoma("Classe") = "VENDAS" Or drSoma("Classe") = "COMPLEMENTACOES" Or drSoma("Classe") = "REAJUSTES" Or drSoma("Classe") = "REMESSAS" Or drSoma("Classe") = "EXPORTACOES") Then
                                NovaLinha("PesoVendas") = drSoma("Peso")
                                NovaLinha("ValorU$Vendas") = drSoma("Valor")
                                NovaLinha("VlrFreteVendas") = drSoma("VlrFrete")
                            Else
                                NovaLinha("PesoVendas") = 0
                                NovaLinha("ValorU$Vendas") = 0
                                NovaLinha("VlrFreteVendas") = 0

                            End If
                            NovaLinha("Empresa_Id") = drSoma("Empresa_Id")
                            NovaLinha("EndEmpresa_Id") = drSoma("EndEmpresa_Id")

                            If drSoma("EntradaSaida") = "S" And drSoma("Classe") = "TRANSFERENCIAS" Then
                                NovaLinha("ValSaidaTrans") = drSoma("Valor")
                                NovaLinha("PesoSaidaTrans") = drSoma("Peso")
                                NovaLinha("VlrFreteSaiTrans") = drSoma("VlrFrete")
                            Else
                                NovaLinha("ValSaidaTrans") = 0
                                NovaLinha("PesoSaidaTrans") = 0
                                NovaLinha("VlrFreteSaiTrans") = 0
                            End If
                            If drSoma("EntradaSaida") = "E" And drSoma("Classe") = "TRANSFERENCIAS" Then
                                NovaLinha("ValEntTrans") = drSoma("Valor")
                                NovaLinha("PesoEntTrans") = drSoma("Peso")
                                NovaLinha("VlrFreteEntTrans") = drSoma("VlrFrete")
                            Else
                                NovaLinha("ValEntTrans") = 0
                                NovaLinha("PesoEntTrans") = 0
                                NovaLinha("VlrFreteEntTrans") = 0
                            End If

                            NovaLinha("NomeCliente") = drSoma("NomeCliente")
                            NovaLinha("Estado_Id") = drSoma("Estado")
                            NovaLinha("Cidade") = drSoma("Cidade")
                            NovaLinha("Grupo_Id") = drSoma("Grupo")
                            NovaLinha("Descricao") = drSoma("DescSubOperacao")

                            dsSoma.Tables(0).Rows.Add(NovaLinha)
                        ElseIf drSomaAux.Length > 0 Then
                            If drSoma("EntradaSaida") = "E" And (drSoma("Classe") = "COMPRAS" Or drSoma("Classe") = "COMPLEMENTACOES" Or drSoma("Classe") = "REAJUSTES" Or drSoma("Classe") = "REMESSAS" Or drSoma("Classe") = "INICIAL" Or drSoma("Classe") = "IMPORTACOES") Then
                                drSomaAux(0)("PesoCompras") = drSomaAux(0)("PesoCompras") + drSoma("Peso")
                                drSomaAux(0)("ValorU$Compras") = drSomaAux(0)("ValorU$Compras") + drSoma("Valor")
                                drSomaAux(0)("VlrFreteCompras") = drSomaAux(0)("VlrFreteCompras") + drSoma("VlrFrete")
                            Else
                                drSomaAux(0)("PesoCompras") = drSomaAux(0)("PesoCompras") + 0
                                drSomaAux(0)("ValorU$Compras") = drSomaAux(0)("ValorU$Compras") + 0
                                drSomaAux(0)("VlrFreteCompras") = drSomaAux(0)("VlrFreteCompras") + 0
                            End If
                            If drSoma("Classe") = "TRANSFERENCIAS" Then
                                drSomaAux(0)("PesoTransf") = drSomaAux(0)("PesoTransf") + drSoma("Peso")
                                drSomaAux(0)("ValorU$Transf") = drSomaAux(0)("ValorU$Transf") + drSoma("Valor")

                            Else
                                drSomaAux(0)("PesoTransf") = drSomaAux(0)("PesoTransf") + 0
                                drSomaAux(0)("ValorU$Transf") = drSomaAux(0)("ValorU$Transf") + 0
                            End If
                            If drSoma("EntradaSaida") = "S" And (drSoma("Classe") = "VENDAS" Or drSoma("Classe") = "COMPLEMENTACOES" Or drSoma("Classe") = "REAJUSTES" Or drSoma("Classe") = "REMESSAS" Or drSoma("Classe") = "EXPORTACOES") Then
                                drSomaAux(0)("PesoVendas") = drSomaAux(0)("PesoVendas") + drSoma("Peso")
                                drSomaAux(0)("ValorU$Vendas") = drSomaAux(0)("ValorU$Vendas") + drSoma("Valor")
                                drSomaAux(0)("VlrFreteVendas") = drSomaAux(0)("VlrFreteVendas") + drSoma("VlrFrete")

                            Else
                                drSomaAux(0)("PesoVendas") = drSomaAux(0)("PesoVendas") + 0
                                drSomaAux(0)("ValorU$Vendas") = drSomaAux(0)("ValorU$Vendas") + 0
                                drSomaAux(0)("VlrFreteVendas") = drSomaAux(0)("VlrFreteVendas") + 0
                            End If
                            If drSoma("EntradaSaida") = "S" And drSoma("Classe") = "TRANSFERENCIAS" Then
                                drSomaAux(0)("ValSaidaTrans") = drSomaAux(0)("ValSaidaTrans") + drSoma("Valor")
                                drSomaAux(0)("PesoSaidaTrans") = drSomaAux(0)("PesoSaidaTrans") + drSoma("Peso")
                                drSomaAux(0)("VlrFreteSaiTrans") = drSomaAux(0)("VlrFreteSaiTrans") + drSoma("VlrFrete")
                            Else
                                drSomaAux(0)("ValSaidaTrans") = drSomaAux(0)("ValSaidaTrans") + 0
                                drSomaAux(0)("PesoSaidaTrans") = drSomaAux(0)("PesoSaidaTrans") + 0
                                drSomaAux(0)("VlrFreteSaiTrans") = drSomaAux(0)("VlrFreteSaiTrans") + 0
                            End If
                            If drSoma("EntradaSaida") = "E" And drSoma("Classe") = "TRANSFERENCIAS" Then
                                drSomaAux(0)("ValEntTrans") = drSomaAux(0)("ValEntTrans") + drSoma("Valor")
                                drSomaAux(0)("PesoEntTrans") = drSomaAux(0)("PesoEntTrans") + drSoma("Peso")
                                drSomaAux(0)("VlrFreteEntTrans") = drSomaAux(0)("VlrFreteEntTrans") + drSoma("VlrFrete")
                            Else
                                drSomaAux(0)("ValEntTrans") = drSomaAux(0)("ValEntTrans") + 0
                                drSomaAux(0)("PesoEntTrans") = drSomaAux(0)("PesoEntTrans") + 0
                                drSomaAux(0)("VlrFreteEntTrans") = drSomaAux(0)("VlrFreteEntTrans") + 0
                            End If
                        End If
                    Next
                    viewReport1(dsSoma, Moeda)
                Else
                    Dim tbOperacao As DataTable = New DataTable("ComercialPorOperacao")
                    Dim dsOperacao As New DataSet
                    Dim drSomaOperacao() As DataRow

                    tbOperacao.Columns.Add("Produto_Id", GetType(String))
                    tbOperacao.Columns.Add("Nome", GetType(String))
                    tbOperacao.Columns.Add("Peso", GetType(Integer)).DefaultValue = 0
                    tbOperacao.Columns.Add("Valor", GetType(Decimal)).DefaultValue = 0
                    tbOperacao.Columns.Add("Empresa_Id", GetType(String))
                    tbOperacao.Columns.Add("EndEmpresa_Id", GetType(String))
                    tbOperacao.Columns.Add("EntradaSaida", GetType(String))
                    tbOperacao.Columns.Add("PesoFrete", GetType(Decimal))
                    tbOperacao.Columns.Add("Operacao_Id", GetType(String))
                    tbOperacao.Columns.Add("SubOperacoes_Id", GetType(String))
                    tbOperacao.Columns.Add("NomeCliente", GetType(String))
                    tbOperacao.Columns.Add("Cidade", GetType(String))
                    tbOperacao.Columns.Add("Estado_Id", GetType(String))
                    tbOperacao.Columns.Add("DescSubOperacao", GetType(String))
                    tbOperacao.Columns.Add("VlrFrete", GetType(Decimal)).DefaultValue = 0
                    tbOperacao.Columns.Add("CidadeEmpresa", GetType(String))
                    dsOperacao.Tables.Add(tbOperacao)

                    For Each drSoma In dsCalculo.Tables(0).Rows

                        strExpr = "Produto_Id = '" & drSoma("Produto_Id") & "' and Empresa_Id = '" & drSoma("Empresa_Id") & "' And EndEmpresa_Id = " & drSoma("EndEmpresa_Id") & " and CidadeEmpresa = '" & drSoma("CidadeEmpresa") & "' And Operacao_Id = " & drSoma("Operacao_Id") & " And SubOperacoes_Id = " & drSoma("SubOperacoes_Id") & " And EntradaSaida = '" & drSoma("EntradaSaida") & "'"
                        drSomaOperacao = dsOperacao.Tables(0).Select(strExpr)
                        If drSomaOperacao.Length = 0 Then
                            NovaLinhaOperacao = dsOperacao.Tables(0).NewRow()
                            NovaLinhaOperacao("Produto_Id") = drSoma("Produto_Id")
                            NovaLinhaOperacao("Nome") = drSoma("Nome")
                            NovaLinhaOperacao("NomeCliente") = drSoma("NomeCliente")
                            NovaLinhaOperacao("Cidade") = drSoma("Cidade")
                            NovaLinhaOperacao("Estado_Id") = drSoma("Estado")
                            NovaLinhaOperacao("DescSubOperacao") = drSoma("DescSubOperacao")
                            NovaLinhaOperacao("Peso") = drSoma("Peso")
                            NovaLinhaOperacao("Valor") = drSoma("Valor")
                            NovaLinhaOperacao("Empresa_Id") = drSoma("Empresa_Id")
                            NovaLinhaOperacao("EndEmpresa_Id") = drSoma("EndEmpresa_Id")
                            NovaLinhaOperacao("EntradaSaida") = drSoma("EntradaSaida")
                            NovaLinhaOperacao("PesoFrete") = drSoma("PesoFrete")
                            NovaLinhaOperacao("Operacao_Id") = drSoma("Operacao_Id")
                            NovaLinhaOperacao("SubOperacoes_Id") = drSoma("SubOperacoes_Id")
                            NovaLinhaOperacao("SubOperacoes_Id") = drSoma("SubOperacoes_Id")
                            NovaLinhaOperacao("VlrFrete") = drSoma("VlrFrete")
                            NovaLinhaOperacao("CidadeEmpresa") = drSoma("CidadeEmpresa")
                            dsOperacao.Tables(0).Rows.Add(NovaLinhaOperacao)

                        ElseIf drSomaOperacao.Length > 0 Then
                            If drSoma("Classe") = "TRANSFERENCIAS" Then
                                If drSoma("EntradaSaida") = "S" Then
                                    'drSomaOperacao(0)("Peso") = drSomaOperacao(0)("Peso") + drSoma("PesoSaidaTrans")
                                    drSomaOperacao(0)("Valor") = drSomaOperacao(0)("Valor") + drSoma("Valor")
                                Else
                                    'drSomaOperacao(0)("Peso") = drSomaOperacao(0)("Peso") + drSoma("PesoEntTrans")
                                    drSomaOperacao(0)("Valor") = drSomaOperacao(0)("Valor") + drSoma("ValEntTrans")
                                End If
                            Else
                                'drSomaOperacao(0)("Peso") = drSomaOperacao(0)("Peso") + drSoma("Peso")
                                drSomaOperacao(0)("Valor") = drSomaOperacao(0)("Valor") + drSoma("Valor")
                                drSomaOperacao(0)("VlrFrete") = drSomaOperacao(0)("VlrFrete") + drSoma("VlrFrete")
                            End If
                        End If
                    Next

                    viewReport2(dsOperacao, Moeda)
                End If
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkHTML_Click(sender As Object, e As EventArgs) Handles lnkHTML.Click
        Try
            'If Funcoes.VerificaPermissao("PrecoMedioV20", "RELATORIO") Then
            Dim linha As String
            Dim sql As String
            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
            Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
            Dim ds As New DataSet
            Dim dr As DataRow
            Dim strm As StreamWriter = Nothing

            If Dir(arquivo).Length > 0 Then Kill(arquivo)

            Dim Moeda As String = String.Empty

            If RadioDolar.Checked Then
                Moeda = "Dolar"
            ElseIf RadioReais.Checked Then
                Moeda = "Reais"
            End If

            sql = " SELECT PxIF.Empresa_Id," & vbCrLf &
                  "        PxIF.Fixacao_Id," & vbCrLf &
                  "        PxIF.Produto_Id," & vbCrLf &
                  "        PxIF.Movimento," & vbCrLf &
                  "        PxIF.Quantidade," & vbCrLf &
                  "        case when '" & Moeda & "'= 'Dolar' then TotalMoeda else TotalOficial end as TotalMoeda," & vbCrLf &
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf &
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf &
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf &
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf &
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf &
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf &
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf &
                  "                   and (PxIFE.Encargo_Id    = 'FUNRURAL' OR PxIFE.Encargo_Id    = 'FUNRURAL JUDICIAL' OR PxIFE.Encargo_Id    = 'SENAR')),0) as FUNRURAL," & vbCrLf &
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf &
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf &
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf &
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf &
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf &
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf &
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf &
                  "                   and PxIFE.Encargo_Id    = 'PIS'),0) as PIS," & vbCrLf &
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf &
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf &
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf &
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf &
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf &
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf &
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf &
                  "                  and PxIFE.Encargo_Id     = 'ICMS'),0) as ICMS," & vbCrLf &
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf &
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf &
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf &
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf &
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf &
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf &
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf &
                  "                   and PxIFE.Encargo_Id    = 'FETHAB'),0) as FETHAB," & vbCrLf

            ''--Calcula valor total"
            sql &= "        case when '" & Moeda & "'= 'Dolar' then TotalMoeda else TotalOficial end" & vbCrLf &
                   "        + isnull((select case when '" & Moeda & "'= 'Dolar' then sum(isnull(ValorMoeda,0.00)) else sum(isnull(ValorOficial,0.00)) end as ValorMoeda" & vbCrLf &
                   "                    from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf &
                   "                   where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf &
                   "                     and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf &
                   "                     and PxIFE.Pedido_Id     = PxIF.Pedido_Id" & vbCrLf &
                   "                     and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf &
                   "                     and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf &
                   "                     and PxIFE.Encargo_Id    in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR','PIS','ICMS','FETHAB')),0) as ValorTotal," & vbCrLf

            ''--calcula medio"
            sql &= "        case" & vbCrLf &
                   "          when Quantidade = 0" & vbCrLf &
                   "            then 0" & vbCrLf &
                   "            else (((case when 'Dolar'= 'Dolar' then TotalMoeda else TotalOficial end" & vbCrLf &
                   "                   +  isnull((select case when 'Dolar'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end  as ValorMoeda" & vbCrLf &
                   "                                from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf &
                   "                               where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf &
                   "                                 and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf &
                   "                                 and PxIFE.Pedido_Id     = PxIF.Pedido_Id" & vbCrLf &
                   "                                 and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf &
                   "                                 and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf &
                   "                                 and PxIFE.Encargo_Id    in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR','PIS','ICMS','FETHAB') ),0))/Quantidade)*1000)" & vbCrLf &
                   "        end as MedioGeral," & vbCrLf &
                   "        Empresa.Nome AS NomeEmpresa, " & vbCrLf &
                   "        Produtos.Nome AS NomeProduto," & vbCrLf &
                   "        SubOperacoes.Operacao_Id," & vbCrLf &
                   "        SubOperacoes.SubOperacoes_Id," & vbCrLf &
                   "        SubOperacoes.Descricao as Operacao," & vbCrLf &
                   "        Empresa.Reduzido," & vbCrLf &
                   "        Empresa.Cidade as CidadeEmpresa," & vbCrLf &
                   "        Empresa.Estado as EstadoEmpresa," & vbCrLf &
                   "        Clientes.Cliente_Id," & vbCrLf &
                   "        Clientes.Nome as NomeCliente, " & vbCrLf &
                   "        Clientes.Cidade AS CidadeCliente," & vbCrLf &
                   "        Clientes.Estado AS EstadoCliente" & vbCrLf &
                   "   FROM VW_PedidosXItensXFixacoes AS PxIF " & vbCrLf &
                   "  INNER Join Clientes AS Empresa" & vbCrLf &
                   "     ON PxIF.Empresa_Id    = Empresa.Cliente_Id" & vbCrLf &
                   "    AND PxIF.EndEmpresa_Id = Empresa.Endereco_Id " & vbCrLf &
                   "  INNER JOIN Produtos" & vbCrLf &
                   "     ON PxIF.Produto_Id = Produtos.Produto_Id" & vbCrLf &
                   "  INNER JOIN Pedidos" & vbCrLf &
                   "     ON PxIF.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf &
                   "    AND PxIF.EndEmpresa_Id = Pedidos.EndEmpresa_Id " & vbCrLf &
                   "    AND PxIF.Pedido_Id     = Pedidos.Pedido_Id" & vbCrLf &
                   "  INNER JOIN SubOperacoes" & vbCrLf &
                   "     ON Pedidos.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                   "    AND Pedidos.SubOperacao = SubOperacoes.SubOperacoes_Id " & vbCrLf &
                   "  INNER JOIN Clientes" & vbCrLf &
                   "     ON Pedidos.Cliente    = Clientes.Cliente_Id" & vbCrLf &
                   "    AND Pedidos.EndCliente = Clientes.Endereco_Id " & vbCrLf &
                   "  Where PxIF.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'" & vbCrLf

            If lstProduto.SelectedValue <> Nothing Then
                sql &= " AND (PxIF.Produto_Id = '" & lstProduto.SelectedValue & "')"
            End If
            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " AND SubString(PxIF.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

            End If

            ds = Banco.ConsultaDataSet(sql, "Clientes")

            linha = "<HTML>" & vbCrLf
            '<HEAD>
            linha &= "<HEAD>" & vbCrLf
            linha &= "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf
            linha &= "<TITLE>Preco Medio</TITLE>" & vbCrLf
            linha &= "</HEAD>" & vbCrLf

            '<BODY>
            linha &= "<BODY>" & vbCrLf

            '-----------------
            'Cabeçalho Padrao
            '-----------------
            linha &= "<table width= '4000' cellpadding='0' cellspacing='0' Border=1>"

            linha &= "<TR>"
            linha &= "<TD>Empresa</TD>"
            linha &= "<TD>Reduzido</TD>"
            linha &= "<TD>NomeEmpresa</TD>"
            linha &= "<TD>CidadeEmpresa</TD>"
            linha &= "<TD>EstadoEmpresa</TD>"
            linha &= "<TD>Fixacao</TD>"
            linha &= "<TD>Produto</TD>"
            linha &= "<TD>NomeProduto</TD>"
            linha &= "<TD>Cliente</TD>"
            linha &= "<TD>NomeCliente</TD>"
            linha &= "<TD>CidadeCliente</TD>"
            linha &= "<TD>EstadoCliente</TD>"
            linha &= "<TD>Operacao</TD>"
            linha &= "<TD>SubOperacao</TD>"
            linha &= "<TD>Descricao</TD>"

            linha &= "<TD >Movimento</TD>"
            linha &= "<TD >Quantidade</TD>"
            linha &= "<TD >Total Moeda</TD>"
            linha &= "<TD >Funrural</TD>"

            linha &= "<TD >Pis</TD>"
            linha &= "<TD >Icms</TD>"
            linha &= "<TD >Fethab</TD>"

            linha &= "<TD >Valor Total</TD>"
            linha &= "<TD >Medio Geral</TD>"

            linha &= "</TR>"

            If ds.Tables(0).Rows.Count > 0 Then
                For Each dr In ds.Tables(0).Rows
                    linha &= "<TR><TD>" & dr("Empresa_Id") & "</TD>"
                    linha &= "<TD>" & dr("Reduzido") & "</TD>"
                    linha &= "<TD>" & dr("NomeEmpresa") & "</TD>"
                    linha &= "<TD>" & dr("CidadeEmpresa") & "</TD>"
                    linha &= "<TD>" & dr("EstadoEmpresa") & "</TD>"
                    linha &= "<TD>" & dr("Fixacao_Id") & "</TD>"
                    linha &= "<TD>" & dr("Produto_Id") & "</TD>"
                    linha &= "<TD>" & dr("NomeProduto") & "</TD>"
                    linha &= "<TD>" & dr("Cliente_Id") & "</TD>"
                    linha &= "<TD>" & dr("NomeCliente") & "</TD>"
                    linha &= "<TD>" & dr("CidadeCliente") & "</TD>"
                    linha &= "<TD>" & dr("EstadoCliente") & "</TD>"
                    linha &= "<TD>" & dr("Operacao_Id") & "</TD>"
                    linha &= "<TD>" & dr("SubOperacoes_Id") & "</TD>"
                    linha &= "<TD>" & dr("Operacao") & "</TD>"
                    linha &= "<TD>" & dr("Movimento") & "</TD>"
                    linha &= "<TD>" & dr("Quantidade") & "</TD>"
                    linha &= "<TD>" & dr("TotalMoeda") & "</TD>"
                    linha &= "<TD>" & dr("Funrural") & "</TD>"
                    linha &= "<TD>" & dr("Pis") & "</TD>"
                    linha &= "<TD>" & dr("Icms") & "</TD>"
                    linha &= "<TD>" & dr("Fethab") & "</TD>"
                    linha &= "<TD>" & dr("ValorTotal") & "</TD>"
                    linha &= "<TD>" & dr("MedioGeral") & "</TD>"
                    linha &= "</TR>"
                Next
            End If
            Try
                strm = New StreamWriter(arquivo, True)
                strm.WriteLine(linha)
                strm.Close()
                ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
            Finally
                strm.Close()
            End Try
            'Else
            'MsgBox(Me.Page, "Usuário sem permissão para emitir HTML.")
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PrecoMedioV20")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class