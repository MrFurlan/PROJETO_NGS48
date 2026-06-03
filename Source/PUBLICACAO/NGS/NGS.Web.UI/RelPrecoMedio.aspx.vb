Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelPrecoMedio
    Inherits BasePage

    Dim sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Comercial)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelPrecoMedio", "ACESSAR") Then
                    txtDataInicio.Text = "01/01/" & Format(Now, "yyyy")
                    txtDataFim.Text = Format(Now, "dd/MM/yyyy")
                    BuscaGrupo()
                    ddl.Carregar(lstEmpresa, CarregarDDL.Tabela.EmpresasConsolidadas, "", False)
                    Funcoes.VerificaEmpresa(lstEmpresa)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Comercial.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            lstEmpresa.Enabled = False
        End If
    End Sub

    Private Sub BuscaProduto()
        Dim sql As String
        Dim Codigo As String
        Dim Descricao As String

        sql = "SELECT Produtos.Produto_Id as Codigo,Produtos.Descricao  as DescricaoProd, Produtos.Agrupar "
        sql &= " FROM Produtos "
        sql &= " INNER Join"
        sql &= " GruposDeEstoques ON Produtos.Grupo = GruposDeEstoques.Grupo_Id"
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

        sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao"
        sql &= " FROM         GruposDeEstoques INNER JOIN"
        sql &= "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo"
        sql &= " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) and Produtos.Agrupar = 'N' order by GruposDeEstoques.Descricao"

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

        sql = "SELECT Safra_Id  "
        sql &= " FROM Safras "

        lstProduto.DataSource = Banco.ConsultaDataSet(sql, "Safras")
        lstProduto.DataTextField = "Safra_Id"
        lstProduto.DataValueField = "Safra_Id"
        lstProduto.DataBind()
        lstProduto.Items.Insert(0, "")
        lstProduto.SelectedIndex = 0
    End Sub

    Protected Sub RelatorioPDF()
        Dim crpt As New ReportDocument()

        Try
            Dim sql As String
            Dim dsCalculo As New DataSet
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

            Dim Moeda As String = String.Empty

            If RadioDolar.Checked Then
                Moeda = "Dolar"
            ElseIf RadioReais.Checked Then
                Moeda = "Reais"
            End If

            Dim int As Integer = 0

            sql = " SELECT  Empresa_Id, EndEmpresa_Id, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf,Peso as PesoFrete, 0.00 AS Peso,"
            sql &= " (SELECT  case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else SUM(ValorOficial) end  AS ValorEncargos"
            sql &= " FROM ConhecimentosDeTransportesXEncargos"
            sql &= " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'OUTROSCREDITOS')) +"
            sql &= " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else  SUM(ValorOficial) end AS ValorEncargos"
            sql &= " FROM ConhecimentosDeTransportesXEncargos"
            sql &= " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'ESTADIAS')) +"
            sql &= " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else SUM(ValorOficial) end AS ValorEncargos"
            sql &= " FROM ConhecimentosDeTransportesXEncargos"
            sql &= " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'SOBRAS')) -"
            sql &= " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else SUM(ValorOficial)  end AS ValorEncargos"
            sql &= " FROM ConhecimentosDeTransportesXEncargos"
            sql &= " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'QUEBRAS')) -"
            sql &= " (SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorOficial) else  SUM(ValorOficial) end AS ValorEncargos"
            sql &= " FROM ConhecimentosDeTransportesXEncargos"
            sql &= " WHERE (ConhecimentosDeTransportes.Empresa_Id = Empresa_Id) AND (ConhecimentosDeTransportes.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (ConhecimentosDeTransportes.Conhecimento_Id = Conhecimento_Id) AND (Encargo_Id = 'OUTROSDESCONTOS ')) AS ValorEncargos, Conhecimento_Id, "
            sql &= " 0.00 AS ValorMoeda, case when '" & Moeda & "'= 'Dolar' then ValorOficial / (select Indice from Cotacoes where Indexador_Id = 3 and Data_Id =Movimento) else  ValorOficial end AS VlrFrete, Serie_Nf, Numero_Nf"
            sql &= " INTO [#TempAuxFretes]"
            sql &= " FROM ConhecimentosDeTransportes"
            sql &= " Where ConhecimentosDeTransportes.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'"

            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " AND SubString(ConhecimentosDeTransportes.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

            End If
            sql &= ";"

            ''Busca Conhecimento de transporte com fixacoes
            sql &= " SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf,PesoFrete, Peso,"
            sql &= " ISNULL(ValorEncargos, 0) AS ValorEncargos,0.00 as ValorMoeda ,VlrFrete + ISNULL(ValorEncargos, 0) AS VlrFrete, Serie_Nf, Numero_Nf, "
            sql &= " NotasFiscais.Operacao, NotasFiscais.SubOperacao, SubOperacoes.Classe, SubOperacoes.Descricao AS DescricaoOperacao,"
            sql &= " (SELECT  TOP (1) Produto_Id"
            sql &= " FROM  NotasFiscaisXItens"
            sql &= " WHERE  (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) "
            sql &= " AND (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) "
            sql &= " AS Produto"
            sql &= " INTO  [#TempAuxFretes1]"
            sql &= " FROM  [#TempAuxFretes] INNER JOIN"
            sql &= " NotasFiscais ON [#TempAuxFretes].Empresa_Id = NotasFiscais.Empresa_Id AND "
            sql &= " [#TempAuxFretes].EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND [#TempAuxFretes].Cliente_Nf = NotasFiscais.Cliente_Id AND "
            sql &= " [#TempAuxFretes].EndCliente_Nf = NotasFiscais.EndCliente_Id AND [#TempAuxFretes].EntradaSaida_Nf = NotasFiscais.EntradaSaida_Id AND "
            sql &= " [#TempAuxFretes].Serie_Nf = NotasFiscais.Serie_Id AND [#TempAuxFretes].Numero_Nf = NotasFiscais.Nota_Id INNER JOIN"
            sql &= " SubOperacoes ON  NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND "
            sql &= " NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id"
            sql &= " INNER Join VW_PedidosXItensXFixacoes pif ON NotasFiscais.Pedido = pif.Pedido_Id"
            sql &= " Where NotasFiscais.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'"

            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " AND SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "
            End If
            sql &= ";"

            ''Busca Conhecimento de transporte sem fixacoes apenas transferencias
            sql &= " insert into [#TempAuxFretes1] SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, Cliente_Nf, EndCliente_Nf, EntradaSaida_Nf,PesoFrete, Peso,"
            sql &= " ISNULL(ValorEncargos, 0) AS ValorEncargos,0.00 as ValorMoeda ,VlrFrete + ISNULL(ValorEncargos, 0) AS VlrFrete, Serie_Nf, Numero_Nf, "
            sql &= " NotasFiscais.Operacao, NotasFiscais.SubOperacao, SubOperacoes.Classe, SubOperacoes.Descricao AS DescricaoOperacao,"
            sql &= " (SELECT  TOP (1) Produto_Id"
            sql &= " FROM  NotasFiscaisXItens"
            sql &= " WHERE  (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) "
            sql &= " AND (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) "
            sql &= " AS Produto"
            '' sql &= " INTO  [#TempAuxFretes1]"
            sql &= " FROM  [#TempAuxFretes] INNER JOIN"
            sql &= " NotasFiscais ON [#TempAuxFretes].Empresa_Id = NotasFiscais.Empresa_Id AND "
            sql &= " [#TempAuxFretes].EndEmpresa_Id = NotasFiscais.EndEmpresa_Id AND [#TempAuxFretes].Cliente_Nf = NotasFiscais.Cliente_Id AND "
            sql &= " [#TempAuxFretes].EndCliente_Nf = NotasFiscais.EndCliente_Id AND [#TempAuxFretes].EntradaSaida_Nf = NotasFiscais.EntradaSaida_Id AND "
            sql &= " [#TempAuxFretes].Serie_Nf = NotasFiscais.Serie_Id AND [#TempAuxFretes].Numero_Nf = NotasFiscais.Nota_Id INNER JOIN"
            sql &= " SubOperacoes ON  NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND "
            sql &= " NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id"
            ' sql &= " INNER Join PedidosXItensXFixacoes ON NotasFiscais.Pedido = PedidosXItensXFixacoes.Pedido_Id"
            sql &= " Where SubOperacoes.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' and NotasFiscais.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'"
            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " AND SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

            End If
            sql &= ";"


            ''Busca notas com fixaçoes 
            sql &= " SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id,Sum(0.00) as PesoFrete, NotasFiscais.Operacao, NotasFiscais.SubOperacao,"
            sql &= " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,"
            sql &= "  (SELECT TOP (1) Produto_Id"
            sql &= "  FROM NotasFiscaisXItens"
            sql &= "  WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= "  (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= "  (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) "
            sql &= "  AS Produto,"
            sql &= " (SELECT     TOP (1) PesoFiscal"
            sql &= " FROM NotasFiscaisXItens"
            sql &= " WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) "
            sql &= " AS Peso, case when '" & Moeda & "'= 'Dolar' then ((Sum(pif.TotalMoeda) / Sum(pif.Quantidade)) *  "
            sql &= " (SELECT     TOP (1) PesoFiscal"
            sql &= " FROM NotasFiscaisXItens"
            sql &= " WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id))) "
            sql &= " else ((Sum(pif.TotalOficial) / Sum(pif.Quantidade)) *  "
            sql &= " (SELECT     TOP (1) PesoFiscal"
            sql &= " FROM NotasFiscaisXItens"
            sql &= " WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id))) "
            sql &= " end  as Valor, 0.00 AS VlrFrete, SubOperacoes.Descricao AS DescricaoOperacao"
            sql &= " INTO [#TempNotasFiscais]"
            sql &= " FROM NotasFiscais INNER JOIN"
            sql &= " SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND "
            sql &= " NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN"
            sql &= " NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND "
            sql &= " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id And NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id And "
            sql &= " NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id And NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id And "
            sql &= " NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id"
            sql &= " INNER Join VW_PedidosXItensXFixacoes pif ON NotasFiscais.Pedido = pif.Pedido_Id"
            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " Where SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

            End If
            sql &= " group by NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Operacao, NotasFiscais.SubOperacao, "
            sql &= " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,NotasFiscais.Serie_Id,NotasFiscais.Nota_Id,SubOperacoes.Descricao;"

            ''Busca notas sem fixaçoes apenas transferencia
            sql &= " insert into [#TempNotasFiscais] SELECT NotasFiscais.Empresa_Id, NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id,Sum(0.00) as PesoFrete, NotasFiscais.Operacao, NotasFiscais.SubOperacao,"
            sql &= " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,"
            sql &= "  (SELECT TOP (1) Produto_Id"
            sql &= "  FROM NotasFiscaisXItens"
            sql &= "  WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= "  (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= "  (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) "
            sql &= "  AS Produto,"
            sql &= " (SELECT     TOP (1) PesoFiscal"
            sql &= " FROM NotasFiscaisXItens"
            sql &= " WHERE (NotasFiscais.Empresa_Id = Empresa_Id) AND (NotasFiscais.EndEmpresa_Id = EndEmpresa_Id) AND "
            sql &= " (NotasFiscais.Cliente_Id = Cliente_Id) AND (NotasFiscais.EndCliente_Id = EndCliente_Id) AND "
            sql &= " (NotasFiscais.EntradaSaida_Id = EntradaSaida_Id) AND (NotasFiscais.Serie_Id = Serie_Id) AND (NotasFiscais.Nota_Id = Nota_Id)) "
            sql &= " AS Peso, case when '" & Moeda & "'= 'Dolar' then Sum(NotasFiscaisXItens.Valor) else Sum(NotasFiscaisXItens.Valor) end  as Valor, 0.00 AS VlrFrete, SubOperacoes.Descricao AS DescricaoOperacao"
            'sql &= " INTO [#TempNotasFiscais]"
            sql &= " FROM NotasFiscais INNER JOIN"
            sql &= " SubOperacoes ON NotasFiscais.Operacao = SubOperacoes.Operacao_Id AND "
            sql &= " NotasFiscais.SubOperacao = SubOperacoes.SubOperacoes_Id INNER JOIN"
            sql &= " NotasFiscaisXItens ON NotasFiscais.Empresa_Id = NotasFiscaisXItens.Empresa_Id AND "
            sql &= " NotasFiscais.EndEmpresa_Id = NotasFiscaisXItens.EndEmpresa_Id And NotasFiscais.Cliente_Id = NotasFiscaisXItens.Cliente_Id And "
            sql &= " NotasFiscais.EndCliente_Id = NotasFiscaisXItens.EndCliente_Id And NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id And "
            sql &= " NotasFiscais.Serie_Id = NotasFiscaisXItens.Serie_Id AND NotasFiscais.Nota_Id = NotasFiscaisXItens.Nota_Id"
            sql &= " where SubOperacoes.Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "'"
            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " and SubString(NotasFiscais.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

            End If
            sql &= " group by NotasFiscais.Empresa_Id,NotasFiscais.EndEmpresa_Id, NotasFiscais.EntradaSaida_Id, NotasFiscais.Operacao, NotasFiscais.SubOperacao, "
            sql &= " SubOperacoes.Classe, NotasFiscais.Cliente_Id, NotasFiscais.EndCliente_Id,NotasFiscais.Serie_Id,NotasFiscais.Nota_Id,SubOperacoes.Descricao;"


            ''Busca fixacoes
            sql &= " SELECT pif.Empresa_Id," & vbCrLf & _
                   "        pif.EndEmpresa_Id," & vbCrLf & _
                   "        pif.Produto_Id," & vbCrLf & _
                   "        SubOperacoes.EntradaSaida," & vbCrLf & _
                   "        0.00 as PesoFrete," & vbCrLf & _
                   "        Pedidos.Operacao," & vbCrLf & _
                   "        Pedidos.SubOperacao," & vbCrLf & _
                   "        SubOperacoes.Classe," & vbCrLf & _
                   "        Pedidos.Cliente," & vbCrLf & _
                   "        Pedidos.EndCliente, " & vbCrLf & _
                   "        Produtos.Descricao AS NomeProduto," & vbCrLf & _
                   "        pif.Quantidade AS Peso," & vbCrLf & _
                   "        case when '" & Moeda & "'= 'Dolar' then  pif.TotalMoeda else pif.TotalOficial end AS Valor, " & vbCrLf & _
                   "        isnull((SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorMoeda) else SUM(ValorOficial) end AS ValorMoedaEncargos" & vbCrLf & _
                   "                  FROM VW_PedidosXItensXFixacoesXEncargos " & vbCrLf & _
                   "                 WHERE Empresa_Id    = pif.Empresa_Id" & vbCrLf & _
                   "                   AND EndEmpresa_Id = pif.EndEmpresa_Id" & vbCrLf & _
                   "                   AND Pedido_Id     = pif.Pedido_Id" & vbCrLf & _
                   "                   AND Produto_Id    = pif.Produto_Id" & vbCrLf & _
                   "                   AND Fixacao_Id    = pif.Fixacao_Id" & vbCrLf & _
                   "                   AND Encargo_Id    = 'FRETE'),0) as VlrFrete, " & vbCrLf & _
                   "        Clientes.Nome AS NomeCliente, " & vbCrLf & _
                   "        Clientes.Cidade," & vbCrLf & _
                   "        Clientes.Estado," & vbCrLf & _
                   "        Produtos.Grupo," & vbCrLf & _
                   "        SubOperacoes.Descricao AS NomeOperacao, " & vbCrLf & _
                   "        ISNULL((SELECT case when '" & Moeda & "'= 'Dolar' then SUM(ValorMoeda) else SUM(ValorOficial) end AS ValorMoedaEncargos" & vbCrLf & _
                   "                  FROM VW_PedidosXItensXFixacoesXEncargos" & vbCrLf & _
                   "                 WHERE Empresa_Id    = pif.Empresa_Id" & vbCrLf & _
                   "                   AND EndEmpresa_Id = pif.EndEmpresa_Id" & vbCrLf & _
                   "                   AND Pedido_Id     = pif.Pedido_Id" & vbCrLf & _
                   "                   AND Produto_Id    = pif.Produto_Id" & vbCrLf & _
                   "                   AND Fixacao_Id    = pif.Fixacao_Id" & vbCrLf & _
                   "                   AND Encargo_Id   in ('PIS','FUNRURAL','FUNRURAL JUDICIAL','SENAR','FETHAB')),0) AS ValorMoedaEncargos " & vbCrLf & _
                   "   INTO [#TempFixacaoDePreco]" & vbCrLf & _
                   "   FROM VW_PedidosXItensXFixacoes pif" & vbCrLf & _
                   "  INNER JOIN Pedidos" & vbCrLf & _
                   "     ON pif.Empresa_Id    = Pedidos.Empresa_Id " & vbCrLf & _
                   "    AND pif.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf & _
                   "    AND pif.Pedido_Id     = Pedidos.Pedido_Id " & vbCrLf & _
                   "  INNER JOIN SubOperacoes" & vbCrLf & _
                   "     ON Pedidos.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
                   "    AND Pedidos.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                   "  INNER JOIN Produtos" & vbCrLf & _
                   "     ON  pif.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                   "  INNER JOIN Clientes" & vbCrLf & _
                   "     ON Pedidos.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf & _
                   "    AND Pedidos.EndEmpresa_Id = Clientes.Endereco_Id " & vbCrLf & _
                   "  Where pif.Movimento BETWEEN '" & txtDataInicio.Text.ToSqlDate() & "' AND '" & txtDataFim.Text.ToSqlDate() & "'" & vbCrLf

            If lstEmpresa.SelectedValue <> Nothing Then
                Dim strEmpresa As String()
                strEmpresa = lstEmpresa.SelectedValue.Split(New Char() {"-"})
                sql &= " AND SubString(pif.Empresa_Id,1,8) = '" & strEmpresa(0) & "' "

            End If
            sql &= ";"

            ''--Cria temp Final
            sql &= " SELECT   [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto AS Produto_Id, EntradaSaida_Nf AS EntradaSaida, Sum(PesoFrete) as PesoFrete,"
            sql &= " Operacao AS Operacao_Id, SubOperacao AS SubOperacoes_Id, Classe, Cliente_Nf AS Cliente_Id, EndCliente_Nf AS EndCliente_Id, Produtos.Nome, "
            sql &= " SUM(0.00) AS Peso, SUM(ValorMoeda) AS Valor, SUM(VlrFrete) AS VlrFrete, Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado, "
            sql &= " Produtos.Grupo, DescricaoOperacao AS DescSubOperacao, SUM(ValorEncargos) AS ValorEncargos,Empresa.Cidade as CidadeEmpresa"
            sql &= " INTO [#TempFinal]"
            sql &= " FROM [#TempAuxFretes1] INNER JOIN"
            sql &= " Produtos ON [#TempAuxFretes1].Produto = Produtos.Produto_Id INNER JOIN"
            sql &= " Clientes ON [#TempAuxFretes1].Empresa_Id = Clientes.Cliente_Id AND [#TempAuxFretes1].EndEmpresa_Id = Clientes.Endereco_Id"
            sql &= " INNER JOIN Clientes as Empresa ON Cliente_Nf = Empresa.Cliente_Id AND EndCliente_Nf = Empresa.Endereco_Id  "
            sql &= " WHERE Classe Is Not NULL and  [#TempAuxFretes1].Empresa_Id = '" & "-1" & "' "
            sql &= " GROUP BY [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto, EntradaSaida_Nf, Operacao, "
            sql &= " SubOperacao, Classe, Cliente_Nf, EndCliente_Nf, Produtos.Nome, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, "
            sql &= " DescricaoOperacao, Empresa.Cidade "


            If dpRelPor.SelectedValue = "Movimento" Then
                ''--Buscas Temporarios Recibo Fretes
                sql &= " insert Into #TempFinal SELECT  [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto AS Produto_Id, EntradaSaida_Nf AS EntradaSaida, Sum(PesoFrete) as PesoFrete ,"
                sql &= " Operacao AS Operacao_Id, SubOperacao AS SubOperacoes_Id, Classe, Cliente_Nf AS Cliente_Id, EndCliente_Nf AS EndCliente_Id, Produtos.Nome, "
                sql &= " SUM(0.00) AS Peso, SUM(ValorMoeda) AS Valor, SUM(VlrFrete) AS VlrFrete, Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado, "
                sql &= " Produtos.Grupo, DescricaoOperacao AS DescSubOperacao, SUM(ValorEncargos) AS ValorEncargos,Empresa.Cidade as CidadeEmpresa"
                sql &= " FROM [#TempAuxFretes1] INNER JOIN"
                sql &= " Produtos ON [#TempAuxFretes1].Produto = Produtos.Produto_Id INNER JOIN"
                sql &= " Clientes ON [#TempAuxFretes1].Empresa_Id = Clientes.Cliente_Id AND [#TempAuxFretes1].EndEmpresa_Id = Clientes.Endereco_Id"
                sql &= " INNER JOIN Clientes as Empresa ON Cliente_Nf = Empresa.Cliente_Id AND EndCliente_Nf = Empresa.Endereco_Id  "
                sql &= " WHERE  (Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' or Classe = '" & eClassesOperacoes.VENDAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "' or Classe = '" & eClassesOperacoes.REAJUSTES.ToString & "' or Classe = '" & eClassesOperacoes.REMESSAS.ToString & "') "
                sql &= " GROUP BY [#TempAuxFretes1].Empresa_Id, [#TempAuxFretes1].EndEmpresa_Id, Produto, EntradaSaida_Nf, Operacao, "
                sql &= " SubOperacao, Classe, Cliente_Nf, EndCliente_Nf, Produtos.Nome, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, "
                sql &= " DescricaoOperacao, Empresa.Cidade"

                ''--Buscas Temporarios Notas Fiscais
                sql &= " insert Into #TempFinal SELECT  Empresa_Id, EndEmpresa_id, Produto AS Produto_Id, EntradaSaida_Id,Sum(PesoFrete) as PesoFrete , Operacao AS Operacao_Id, "
                sql &= " SubOperacao AS SubOperacoes_Id, Classe, [#TempNotasFiscais].Cliente_Id, EndCliente_Id, Produtos.Nome, SUM(ISNULL(Peso, 0)) AS Peso, "
                sql &= " SUM(Valor) AS Valor, SUM(VlrFrete) AS VlrFrete, Clientes.Nome AS NomeCliente, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, "
                sql &= " DescricaoOperacao AS DescSubOperacao, SUM(0.00) AS ValorEncargos,Empresa.Cidade as CidadeEmpresa "
                sql &= " FROM [#TempNotasFiscais] INNER JOIN "
                sql &= " Produtos ON [#TempNotasFiscais].Produto = Produtos.Produto_Id INNER JOIN "
                sql &= " Clientes ON [#TempNotasFiscais].Empresa_Id = Clientes.Cliente_Id AND [#TempNotasFiscais].EndEmpresa_Id = Clientes.Endereco_Id"
                sql &= " INNER JOIN Clientes as Empresa ON [#TempNotasFiscais].Cliente_Id = Empresa.Cliente_Id AND EndCliente_Id = Empresa.Endereco_Id "
                sql &= " WHERE     (Classe = '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPRAS.ToString & "' or Classe = '" & eClassesOperacoes.VENDAS.ToString & "' or Classe = '" & eClassesOperacoes.COMPLEMENTACOES.ToString & "' or Classe = '" & eClassesOperacoes.REAJUSTES.ToString & "' or Classe = '" & eClassesOperacoes.REMESSAS.ToString & "')"
                sql &= " GROUP BY Empresa_Id, EndEmpresa_id, Produto, EntradaSaida_Id, Operacao, SubOperacao, Classe, [#TempNotasFiscais].Cliente_Id, EndCliente_Id, "
                sql &= " Produtos.Nome, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Produtos.Grupo, DescricaoOperacao, Empresa.Cidade "

            End If

            If dpRelPor.SelectedValue = "Fixações" Then
                ''--Buscas Temporarios Fixacao de Precos
                sql &= " insert Into #TempFinal SELECT      Empresa_Id, EndEmpresa_Id, Produto_Id, EntradaSaida,Sum(PesoFrete) as PesoFrete , Operacao as Operacao_Id, SubOperacao AS SubOperacoes_Id, Classe, "
                sql &= " Cliente AS Cliente_Id, EndCliente AS EndCliente_Id, NomeProduto AS Nome, SUM(Peso) AS Peso, SUM(Valor - ValorMoedaEncargos) AS Valor, "
                sql &= " SUM(VlrFrete) AS VlrFrete, NomeCliente, [#TempFixacaoDePreco].Cidade, [#TempFixacaoDePreco].Estado, Grupo, NomeOperacao AS DescSubOperacao, SUM(ValorMoedaEncargos) "
                sql &= " AS ValorEncargos,Empresa.Cidade as CidadeEmpresa "
                sql &= " FROM [#TempFixacaoDePreco] "
                sql &= " INNER JOIN Clientes as Empresa ON [#TempFixacaoDePreco].Cliente = Empresa.Cliente_Id AND [#TempFixacaoDePreco].EndCliente = Empresa.Endereco_Id "
                sql &= " WHERE(Classe Is Not NULL) "
                sql &= " GROUP BY Empresa_Id, EndEmpresa_Id, Produto_Id, EntradaSaida, Operacao, SubOperacao, Classe, Cliente, EndCliente, NomeProduto, NomeCliente, "
                sql &= " [#TempFixacaoDePreco].Cidade, [#TempFixacaoDePreco].Estado, Grupo, NomeOperacao,Empresa.Cidade "
            End If

            sql &= " select Empresa_Id,EndEmpresa_Id,case 'N' when 'S' then #TempFinal.Produto_Id else #TempFinal.Grupo end as Produto_Id,EntradaSaida,Sum(PesoFrete) as PesoFrete,Operacao_Id,SubOperacoes_Id,Classe,CASE #TempFinal.Classe WHEN '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' then Cliente_Id else '' end as Cliente_Id ,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' then  EndCliente_Id else 0 end as EndCliente_Id,case 'N' when 'S' then #TempFinal.Nome else GruposDeEstoques.Descricao end as Nome,NomeCliente,Sum(Peso) as Peso,isnull(Sum(Valor),0) as Valor,Sum(VlrFrete) as VlrFrete ,NomeCliente,Cidade,Estado,#TempFinal.Grupo,DescSubOperacao,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' THEN #TempFinal.CidadeEmpresa else '' end as CidadeEmpresa from #TempFinal"
            sql &= " inner join Produtos On #TempFinal.Produto_Id = Produtos.Produto_Id"
            sql &= " INNER JOIN GruposDeEstoques ON  #TempFinal.Grupo = GruposDeEstoques.Grupo_Id"

            sql &= " Where #TempFinal.Produto_Id is not null  "
            If lstProduto.SelectedValue <> Nothing Then
                sql &= " and #TempFinal.Produto_Id='" & lstProduto.SelectedValue & "' "
            End If
            If lstGrupoProduto.SelectedValue <> Nothing Then
                sql &= " and GruposDeEstoques.Grupo_Id='" & lstGrupoProduto.SelectedValue & "' "
            End If

            sql &= " group by Empresa_Id,EndEmpresa_Id,case 'N' when 'S' then #TempFinal.Produto_Id else #TempFinal.Grupo end,EntradaSaida,Operacao_Id,SubOperacoes_Id,Classe,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' then Cliente_Id else '' end  ,CASE #TempFinal.Classe WHEN '" & eClassesOperacoes.TRANSFERENCIAS.ToString & "' then  EndCliente_Id else 0 end,case 'N' when 'S' then #TempFinal.Nome else GruposDeEstoques.Descricao end,NomeCliente,NomeCliente,Cidade,Estado,#TempFinal.Grupo,DescSubOperacao,CASE #TempFinal.Classe WHEN 'TRANSFERENCIAS' THEN #TempFinal.CidadeEmpresa else '' end"
            sql &= " order by Empresa_Id,EndEmpresa_Id,Produto_Id,EntradaSaida,Operacao_Id,SubOperacoes_Id,Cliente_Id,EndCliente_Id"


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

            If DropOperacao.SelectedValue = "N" Then

                Dim tbSoma As DataTable = New DataTable("ComercialDeInsumos")
                Dim dsSoma As New DataSet
                'Dim drSomaOperacao() As DataRow

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

                crpt.FileName = Server.MapPath("~/Reports/CrPrecoMedioGraneis.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim arquivo As String = Server.MapPath(NomeArquivo)

                crpt.SetDataSource(dsSoma)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As ParameterFieldDefinitions
                Dim crparameterfielddefinition As ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Session("ssNomeEmpresa")
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Session("ssCidadeEmpresa")
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Estado")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Session("ssEstadoEmpresa")
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

                crpt.FileName = "~/Reports/CrPrecoMedioGraneisOperacao.rpt"
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim arquivo As String = Server.MapPath(NomeArquivo)

                crpt.SetDataSource(dsOperacao)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As ParameterFieldDefinitions
                Dim crparameterfielddefinition As ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Session("ssNomeEmpresa")
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Session("ssCidadeEmpresa")
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                crparameterfielddefinition = crparameterfielddefinitions.Item("Estado")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = Session("ssEstadoEmpresa")
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
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        Finally
            crpt.Close()
            crpt.Dispose()
        End Try
    End Sub

    Protected Sub RelatorioHTML()
        Try
            Dim linha As String
            Dim sql As String
            Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
            Dim arquivo As String = Server.MapPath(NomeArquivo)
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

            sql = " SELECT PxIF.Empresa_Id," & vbCrLf & _
                  "        PxIF.Fixacao_Id," & vbCrLf & _
                  "        PxIF.Produto_Id," & vbCrLf & _
                  "        PxIF.Movimento," & vbCrLf & _
                  "        PxIF.Quantidade," & vbCrLf & _
                  "        case when '" & Moeda & "'= 'Dolar' then TotalMoeda else TotalOficial end as TotalMoeda," & vbCrLf & _
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf & _
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf & _
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf & _
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf & _
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf & _
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf & _
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf & _
                  "                   and (PxIFE.Encargo_Id    = 'FUNRURAL' OR PxIFE.Encargo_Id    = 'FUNRURAL JUDICIAL' OR PxIFE.Encargo_Id    = 'SENAR')),0) as FUNRURAL," & vbCrLf & _
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf & _
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf & _
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf & _
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf & _
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id" & vbCrLf & _
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf & _
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf & _
                  "                   and PxIFE.Encargo_Id    = 'PIS'),0) as PIS," & vbCrLf & _
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf & _
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf & _
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf & _
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf & _
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf & _
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf & _
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf & _
                  "                   and PxIFE.Encargo_Id    ='ICMS'),0) as ICMS," & vbCrLf & _
                  "        isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf & _
                  "                  from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf & _
                  "                 where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf & _
                  "                   and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf & _
                  "                   and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf & _
                  "                   and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf & _
                  "                   and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf & _
                  "                   and PxIFE.Encargo_Id    ='FETHAB'),0) as FETHAB," & vbCrLf
            ''--Calcula valor total"
            sql &= "        case when '" & Moeda & "'= 'Dolar' then TotalMoeda else TotalOficial end" & vbCrLf & _
                   "        + isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end as ValorMoeda" & vbCrLf & _
                   "                    from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf & _
                   "                   where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf & _
                   "                     and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf & _
                   "                     and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf & _
                   "                     and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf & _
                   "                     and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf & _
                   "                     and PxIFE.Encargo_Id    in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR','PIS','ICMS','FETHAB' )),0) as ValorTotal," & vbCrLf

            ''--calcula medio"
            sql &= "        ((case when '" & Moeda & "'= 'Dolar' then TotalMoeda else TotalOficial end" & vbCrLf & _
                   "        + isnull((select case when '" & Moeda & "'= 'Dolar' then sum(ValorMoeda) else sum(ValorOficial) end  as ValorMoeda" & vbCrLf & _
                   "                    from VW_PedidosXItensXFixacoesXEncargos PxIFE" & vbCrLf & _
                   "                   where PxIFE.Empresa_Id    = PxIF.Empresa_Id" & vbCrLf & _
                   "                     and PxIFE.EndEmpresa_Id = PxIF.EndEmpresa_Id" & vbCrLf & _
                   "                     and PxIFE.Pedido_Id     = PxIF.Pedido_Id " & vbCrLf & _
                   "                     and PxIFE.Produto_Id    = PxIF.Produto_Id" & vbCrLf & _
                   "                     and PxIFE.Fixacao_Id    = PxIF.Fixacao_Id" & vbCrLf & _
                   "                     and PxIFE.Encargo_Id    in ('FUNRURAL','FUNRURAL JUDICIAL','SENAR','PIS','ICMS','FETHAB')),0)/Quantidade)*1000) as MedioGeral," & vbCrLf & _
                   "        Empresa.Nome AS NomeEmpresa, " & vbCrLf & _
                   "        Produtos.Nome AS NomeProduto," & vbCrLf & _
                   "        SubOperacoes.Operacao_Id," & vbCrLf & _
                   "        SubOperacoes.SubOperacoes_Id," & vbCrLf & _
                   "        SubOperacoes.Descricao as Operacao," & vbCrLf & _
                   "        Empresa.Reduzido," & vbCrLf & _
                   "        Empresa.Cidade as CidadeEmpresa," & vbCrLf & _
                   "        Empresa.Estado as EstadoEmpresa," & vbCrLf & _
                   "        Clientes.Cliente_Id," & vbCrLf & _
                   "        Clientes.Nome as NomeCliente, " & vbCrLf & _
                   "        Clientes.Cidade AS CidadeCliente," & vbCrLf & _
                   "        Clientes.Estado AS EstadoCliente" & vbCrLf & _
                   "   FROM VW_PedidosXItensXFixacoes AS PxIF " & vbCrLf & _
                   "  INNER Join Clientes AS Empresa" & vbCrLf & _
                   "     ON PxIF.Empresa_Id = Empresa.Cliente_Id" & vbCrLf & _
                   "    AND PxIF.EndEmpresa_Id = Empresa.Endereco_Id" & vbCrLf & _
                   "  INNER JOIN Produtos" & vbCrLf & _
                   "     ON PxIF.Produto_Id = Produtos.Produto_Id" & vbCrLf & _
                   "  INNER JOIN Pedidos" & vbCrLf & _
                   "     ON PxIF.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf & _
                   "    AND PxIF.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf & _
                   "    AND PxIF.Pedido_Id     = Pedidos.Pedido_Id" & vbCrLf & _
                   "  INNER JOIN SubOperacoes" & vbCrLf & _
                   "     ON Pedidos.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf & _
                   "    AND Pedidos.SubOperacao = SubOperacoes.SubOperacoes_Id" & vbCrLf & _
                   "  INNER JOIN Clientes" & vbCrLf & _
                   "     ON Pedidos.Cliente    = Clientes.Cliente_Id" & vbCrLf & _
                   "    AND Pedidos.EndCliente = Clientes.Endereco_Id " & vbCrLf & _
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
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lstGrupoProduto_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles lstGrupoProduto.SelectedIndexChanged
        Try
            BuscaProduto()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelPrecoMedio", "RELATORIO") Then
                If RadPDF.Checked Then
                    RelatorioPDF()
                Else
                    RelatorioHTML()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelPrecoMedio")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class