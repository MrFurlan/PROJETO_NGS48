Imports NGS.Lib.Negocio

Public Class VincularNotaxTitulo
    Inherits BasePage

    Dim Sql As String
    Dim SqlArray As New ArrayList

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("VincularNotaxTitulo", "ACESSAR") Then
                CargaUnidadeDeNegocioEmpresaCliente()
                HID.Value = Guid.NewGuid().ToString()
                ucConsultaClientes.SetarHID(HID.Value)
                lnkVincular.Parent.Visible = False
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
    End Sub

    Private Function Validar() As Boolean
        If String.IsNullOrWhiteSpace(DdlUnidadeConsultaTitulos.SelectedValue) Then
            MsgBox(Me.Page, "Informe a unidade de negócio!")
            Return False
        End If
        If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
            MsgBox(Me.Page, "Informe a empresa!")
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            MsgBox(Me.Page, "Informe o cliente!")
            Return False
        End If
        If String.IsNullOrWhiteSpace(txtCodigoPedido.Value) Then
            MsgBox(Me.Page, "Informe o pedido!")
            Return False
        End If
        Return True
    End Function
    Private Function ObterTipoTitulo(tituloId As String) As String
        Dim sql As String
        Dim ds As DataSet

        ' Verifica se está no ContasAPagar
        sql = "SELECT 1 FROM ContasAPagar WHERE Registro_Id = " & tituloId
        ds = Banco.ConsultaDataSet(sql, "ContasAPagar")
        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return "P" ' Contas a Pagar
        End If

        ' Verifica se está no ContasAReceber
        sql = "SELECT 1 FROM ContasAReceber WHERE Registro_Id = " & tituloId
        ds = Banco.ConsultaDataSet(sql, "ContasAReceber")
        If ds IsNot Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
            Return "R" ' Contas a Receber
        End If

        Return "" ' Não encontrado
    End Function

    Private Sub Vincular()
        Sql = "INSERT INTO NotaFiscalXTitulo                                                                                                           " & vbCrLf &
              "           (Empresa_Id                                                                                                                  " & vbCrLf &
              "           ,EndEmpresa_Id                                                                                                               " & vbCrLf &
              "           ,Cliente_Id                                                                                                                  " & vbCrLf &
              "           ,EndCliente_Id                                                                                                               " & vbCrLf &
              "           ,EntradaSaida_Id                                                                                                             " & vbCrLf &
              "           ,Serie_Id                                                                                                                    " & vbCrLf &
              "           ,Nota_Id                                                                                                                     " & vbCrLf &
              "           ,Titulo_Id)                                                                                                                  " & vbCrLf &
              "        VALUES(                                                                                                                         " & vbCrLf &
              "           '" & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(1).Text & "'                                                  " & vbCrLf &
              "           ," & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(2).Text & "                                                   " & vbCrLf &
              "           ,'" & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(3).Text & "'                                                 " & vbCrLf &
              "           ," & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(4).Text & "                                                   " & vbCrLf &
              "           ,'" & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(7).Text & "'                                                 " & vbCrLf &
              "           ," & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(5).Text & "                                                   " & vbCrLf &
              "           ," & GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(6).Text & "                                                   " & vbCrLf &
              "           ," & GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(2).Text & ")" & vbCrLf

        SqlArray.Clear()
        SqlArray.Add(Sql)
        Dim tabelaFinanceira As String = IIf(GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(1).Text = "P", "ContasAPagar", "ContasAReceber")
        Dim registroId As String = GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(2).Text
        Dim contratoBancario As String = Funcoes.EliminarCaracteresEspeciais(txtContratoBancario.Text)

        ' Montar texto da observação
        Dim usuario As String = Funcoes.EliminarCaracteresEspeciais(Session("ssNomeUsuario"))
        Dim dataHora As String = Now.ToString("dd/MM/yyyy HH:mm")
        Dim notaId As String = GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(6).Text
        Dim tituloId As String = GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(2).Text
        Dim textoLog As String = "Vinculado o título " & tituloId & " da nota " & notaId & " por " & usuario & " em " & dataHora
        textoLog = Replace(textoLog, "'", "''") ' Escapar aspas simples

        ' Atualizar o contrato bancário e adicionar log na observação (concatenando com a anterior)
        Sql = "UPDATE " & tabelaFinanceira & vbCrLf &
      "SET ContratoBancario = '" & contratoBancario & "'," & vbCrLf &
      "    ObservacoesControleInterno = CASE" & vbCrLf &
      "        WHEN ObservacoesControleInterno LIKE '%" & textoLog & "%' THEN ObservacoesControleInterno" & vbCrLf &
      "        ELSE ISNULL(ObservacoesControleInterno + CHAR(13) + CHAR(10), '') + '" & textoLog & "'" & vbCrLf &
      "    END" & vbCrLf &
      "WHERE Registro_id = " & registroId

        SqlArray.Add(Sql)

        If Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, "Titulo vinculado com Sucesso.", eTitulo.Sucess)
        Else
            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
        End If
    End Sub

    Private Sub Desvincular(ByVal tituloId As String, ByVal notaId As String, ByVal tipo As String, ByVal registroId As String)
        Dim usuario As String = Funcoes.EliminarCaracteresEspeciais(Session("ssNomeUsuario"))
        Dim dataHora As String = Now.ToString("dd/MM/yyyy HH:mm")
        Dim textoLog As String = "Desvinculado o título " & tituloId & " da nota " & notaId & " por " & usuario & " em " & dataHora
        textoLog = Replace(textoLog, "'", "''")

        Dim tabelaFinanceira As String = IIf(tipo = "P", "ContasAPagar", "ContasAReceber")

        Dim SqlArray As New ArrayList

        ' Atualizar log no campo de observação
        Dim sql As String = "UPDATE " & tabelaFinanceira & vbCrLf &
    "SET ObservacoesControleInterno = CASE" & vbCrLf &
    "    WHEN ObservacoesControleInterno LIKE '%" & textoLog & "%' THEN ObservacoesControleInterno" & vbCrLf &
    "    ELSE ISNULL(ObservacoesControleInterno + CHAR(13) + CHAR(10), '') + '" & textoLog & "'" & vbCrLf &
    "END" & vbCrLf &
    "WHERE Registro_id = " & registroId

        SqlArray.Add(sql)

        Banco.GravaBanco(SqlArray)
    End Sub

    Private Sub Limpar()

        Session.Remove("objClienteExtrato" & HID.Value)
        Session.Remove("objVincularNotaxTitulo" & HID.Value)

        txtCodigoCliente.Value = String.Empty
        txtClientes.Text = String.Empty
        txtPedido.Text = String.Empty
        txtContratoBancario.Text = String.Empty
        NumContratoBancario.Visible = False
        GridTitulo.DataSource = New List(Of Object)
        GridTitulo.DataBind()
        GridNotaxTitulo.DataSource = New List(Of Object)
        GridNotaxTitulo.DataBind()

        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        ucConsultaPedidos.SetarHID(HID.Value)

        LiberaEmpresa()
    End Sub
    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Private Sub CarregarTitulos()
        Dim strEmpresa() As String = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")

        Dim sql As String = "select t.* from (" & vbCrLf &
                               "select 'P' AS Tipo, ContasAPagar.Registro_Id, isnull(nXt.Nota_Id,0) AS Nota, ContasAPagar.Movimento, ContasAPagar.Vencimento, isnull(ContasAPagar.ContratoBancario,'') AS ContratoBancario, ContasAPagar.ValorDoDocumento " & vbCrLf &
                               "from ContasAPagar " & vbCrLf &
                               "        left join NotaFiscalXTitulo nXt" & vbCrLf &
                               "				on  nXt.Empresa_Id      = ContasAPagar.EmpresaPedido " & vbCrLf &
                               "				and nXt.EndEmpresa_Id   = ContasAPagar.EndEmpresaPedido " & vbCrLf &
                               "				and nXt.Titulo_Id       = ContasAPagar.Registro_Id " & vbCrLf &
                               "where ContasAPagar.EmpresaPedido = '" & strEmpresa(0) & "'" & vbCrLf &
                               "  and ContasAPagar.EndEmpresaPedido = " & strEmpresa(1) & vbCrLf &
                               "  and ContasAPagar.Pedido = " & txtPedido.Text & vbCrLf &
                               "  and ContasAPagar.Situacao = 1" & vbCrLf &
                               "  and ContasAPagar.Grupado not in ('M')" & vbCrLf
        If chkBaixado.Checked Then
            sql &= "  and ContasAPagar.provisao = 1" & vbCrLf
        Else
            sql &= "  and ContasAPagar.provisao not in(3)" & vbCrLf
        End If

        sql &= "union all " & vbCrLf &
                               "select 'R' AS Tipo, ContasAReceber.Registro_Id, isnull(nXt.Nota_Id,0) AS Nota, ContasAReceber.Movimento, ContasAReceber.Vencimento, isnull(ContasAReceber.ContratoBancario,'') AS ContratoBancario, ContasAReceber.ValorDoDocumento " & vbCrLf &
                               "from ContasAReceber " & vbCrLf &
                               "        left join NotaFiscalXTitulo nXt" & vbCrLf &
                               "				on  nXt.Empresa_Id      = ContasAReceber.EmpresaPedido " & vbCrLf &
                               "				and nXt.EndEmpresa_Id   = ContasAReceber.EndEmpresaPedido " & vbCrLf &
                               "				and nXt.Titulo_Id       = ContasAReceber.Registro_Id " & vbCrLf &
                               "where ContasAReceber.EmpresaPedido = '" & strEmpresa(0) & "'" & vbCrLf &
                               "  and ContasAReceber.EndEmpresaPedido = " & strEmpresa(1) & vbCrLf &
                               "  and ContasAReceber.Pedido = " & txtPedido.Text & vbCrLf &
                               "  and ContasAReceber.Situacao = 1" & vbCrLf &
                               "  and ContasAReceber.Grupado not in ('M')" & vbCrLf
        If chkBaixado.Checked Then
            sql &= "  and ContasAReceber.provisao = 1" & vbCrLf
        Else
            sql &= "  and ContasAReceber.provisao not in(3)" & vbCrLf
        End If

        sql &= ") as t order by t.Movimento"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Titulos")

        If ds Is Nothing OrElse ds.Tables Is Nothing OrElse Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Nenhum registro encontrado!")
        Else
            GridTitulo.DataSource = ds
            GridTitulo.DataBind()
        End If
    End Sub

    Private Sub CarregarTitulosxNotas()
        Dim strEmpresa() As String = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")

        Sql &= "select  nf.Nota_Id, " & vbCrLf &
                "     	nf.Serie_Id, " & vbCrLf &
                "     	nf.EntradaSaida_Id, " & vbCrLf &
                "     	nf.Cliente_Id, " & vbCrLf &
                "     	nf.EndCliente_Id, " & vbCrLf &
                "     	nf.Empresa_Id, " & vbCrLf &
                "     	nf.EndEmpresa_Id, " & vbCrLf &
                "     	Item.Valor, " & vbCrLf &
                "     	isnull(nXt.Titulo_Id,0) AS Titulo_Id, 
                (SELECT TOP 1 'P' FROM ContasAPagar WHERE Registro_Id = nXt.Titulo_Id
                UNION
                 SELECT TOP 1 'R' FROM ContasAReceber WHERE Registro_Id = nXt.Titulo_Id) AS Tipo" & vbCrLf &
                "    from NotasFiscais nf " & vbCrLf &
                "		inner join (Select nXi.Empresa_Id, nXi.EndEmpresa_Id, nXi.Cliente_Id, nXi.EndCliente_Id, " & vbCrLf &
                "		                   nXi.EntradaSaida_Id, nXi.Serie_Id, nXi.Nota_Id, sum(nXi.Valor) AS Valor " & vbCrLf &
                "		            From NotasFiscaisXItens nXi " & vbCrLf &
                "		            where nXi.Empresa_Id = '" & strEmpresa(0) & "'" & vbCrLf &
                "                      and nXi.EndEmpresa_Id = " & strEmpresa(1) & vbCrLf &
                "                      and nXi.Pedido = " & txtPedido.Text.Trim() & vbCrLf &
                "                    group by nXi.Empresa_Id, nXi.EndEmpresa_Id, nXi.Cliente_Id, nXi.EndCliente_Id, " & vbCrLf &
                "                             nXi.EntradaSaida_Id, nXi.Serie_Id, nXi.Nota_Id) AS Item " & vbCrLf &
                "				on  Item.Empresa_Id      = nf.Empresa_Id " & vbCrLf &
                "				and Item.EndEmpresa_Id   = nf.EndEmpresa_Id " & vbCrLf &
                "				and Item.Cliente_Id      = nf.Cliente_Id " & vbCrLf &
                "				and Item.EndCliente_Id   = nf.EndCliente_Id " & vbCrLf &
                "				and Item.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
                "				and Item.Serie_Id        = nf.Serie_Id " & vbCrLf &
                "				and Item.Nota_Id         = nf.Nota_Id " & vbCrLf &
                "        left join NotaFiscalXTitulo nXt " & vbCrLf &
                "				on  nXt.Empresa_Id      = nf.Empresa_Id " & vbCrLf &
                "				and nXt.EndEmpresa_Id   = nf.EndEmpresa_Id " & vbCrLf &
                "				and nXt.Cliente_Id      = nf.Cliente_Id " & vbCrLf &
                "				and nXt.EndCliente_Id   = nf.EndCliente_Id " & vbCrLf &
                "				and nXt.EntradaSaida_Id = nf.EntradaSaida_Id " & vbCrLf &
                "				and nXt.Serie_Id        = nf.Serie_Id " & vbCrLf &
                "				and nXt.Nota_Id         = nf.Nota_Id " & vbCrLf &
                "    where nf.Empresa_Id = '" & strEmpresa(0) & "'" & vbCrLf &
                "     and nf.EndEmpresa_Id = " & strEmpresa(1) & vbCrLf &
                "     and nf.Pedido = " & txtPedido.Text.Trim() & vbCrLf &
                "     and nf.Situacao = 1 " & vbCrLf &
                "     and nf.TipoDeDocumento = 1 "

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotasFiscais")
        GridNotaxTitulo.DataSource = ds
        GridNotaxTitulo.DataBind()

        If ds Is Nothing OrElse ds.Tables Is Nothing OrElse Not ds.Tables(0).Rows.Count > 0 Then
            MsgBox(Me.Page, "Nenhum registro encontrado!")
            Exit Sub
        End If
    End Sub

    Private Sub Relatorio()

        Dim strEmpresa() As String = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")

        Sql = "SELECT t.Registro_Id, " & vbCrLf &
                "		case " & vbCrLf &
                "			when t.Provisao = 1" & vbCrLf &
                "				then 'BAIXADO'" & vbCrLf &
                "				else 'ABERTO'" & vbCrLf &
                "			end as Situacao, " & vbCrLf &
                "		convert(varchar,t.Prorrogacao,103) as Vencimento, convert(varchar,t.Baixa,103) as Baixa, t.Cliente, c.Nome, t.ValorLiquido, t.Historico, " & vbCrLf &
                "		case" & vbCrLf &
                "			when isnull(nt.Nota_Id,0) > 0" & vbCrLf &
                "				then nt.EntradaSaida_Id + '-' + convert(varchar,nt.Nota_id)" & vbCrLf &
                "				else ''" & vbCrLf &
                "			end as Nota" & vbCrLf &
                "FROM   ContasAReceber t" & vbCrLf &
                "		inner join Clientes c" & vbCrLf &
                "				on c.Cliente_id = t.Cliente" & vbCrLf &
                "				and c.Endereco_id = t.EndCliente" & vbCrLf &
                "		left join NotaFiscalXTitulo nt" & vbCrLf &
                "				on nt.Titulo_id = t.Registro_id" & vbCrLf &
                "where t.Empresa    = '" & strEmpresa(0) & "'" & vbCrLf &
                "  and t.EndEmpresa = " & strEmpresa(1) & vbCrLf &
                "  and t.Pedido     = " & txtPedido.Text.Trim() & vbCrLf &
                "order by nt.Nota_id" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf &
                "" & vbCrLf &
                "select (ni.EntradaSaida_Id + '-' + convert(varchar,ni.Nota_id)) as Nota, ni.Valor as ValorDaNota, isnull(t.Registro_Id,0) as Titulo, isnull(t.ValorLiquido,0) as ValorDoTitulo, isnull(t.Historico,'') as Historico" & vbCrLf &
                "from NotasFiscaisXItens ni" & vbCrLf &
                "		left join NotasFiscais n" & vbCrLf &
                "				ON n.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "		 		and n.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "				and n.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "				and n.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "				and n.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "				and n.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "				and n.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "		left join NotaFiscalXTitulo nt" & vbCrLf &
                "				ON nt.Empresa_Id       = ni.Empresa_Id" & vbCrLf &
                "				and nt.EndEmpresa_iD   = ni.EndEmpresa_Id" & vbCrLf &
                "				and nt.Cliente_Id      = ni.Cliente_Id" & vbCrLf &
                "				and nt.EndCliente_Id   = ni.EndCliente_Id" & vbCrLf &
                "				and nt.EntradaSaida_Id = ni.EntradaSaida_Id" & vbCrLf &
                "				and nt.Serie_Id        = ni.Serie_Id" & vbCrLf &
                "				and nt.Nota_Id         = ni.Nota_Id" & vbCrLf &
                "		left join ContasAReceber t" & vbCrLf &
                "				on t.Registro_id = nt.Titulo_Id" & vbCrLf &
                "where n.Empresa_Id = '" & strEmpresa(0) & "'" & vbCrLf &
                "  and n.EndEmpresa_Id = " & strEmpresa(1) & vbCrLf &
                "  and n.Pedido = " & txtPedido.Text.Trim() & vbCrLf &
                "  and n.Situacao = 1 " & vbCrLf &
                "  and n.TipoDeDocumento = 1 " & vbCrLf &
                "order by n.Movimento" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "NotasFiscais")

        If ds Is Nothing OrElse ds.Tables Is Nothing OrElse Not ds.Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Nenhum registro encontrado!")
            Exit Sub
        End If

        MsgBox(Me.Page, "Tudo Certo!")

    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objClienteExtrato" & HID.Value) IsNot Nothing Then
            Dim cli As Cliente = Session("objClienteExtrato" & HID.Value)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(cli)
            txtClientes.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteExtrato" & HID.Value)

            txtPedido.Text = String.Empty
            GridTitulo.DataSource = New List(Of Object)
            GridTitulo.DataBind()
            GridNotaxTitulo.DataSource = New List(Of Object)
            GridNotaxTitulo.DataBind()

            BuscarPedido()

        ElseIf Session("objVincularNotaxTitulo" & HID.Value) IsNot Nothing Then
            Dim p As [Lib].Negocio.Pedido = CType(Session("objVincularNotaxTitulo" & HID.Value), [Lib].Negocio.Pedido)
            txtCodigoPedido.Value = p.Codigo
            txtPedido.Text = p.Codigo
            Session.Remove("objVincularNotaxTitulo" & HID.Value)

            GridTitulo.DataSource = New List(Of Object)
            GridTitulo.DataBind()
            GridNotaxTitulo.DataSource = New List(Of Object)
            GridNotaxTitulo.DataBind()

            lnkConsultar_Click(lnkConsultar, Nothing)
        End If
    End Sub

    Protected Sub imgDesvincularTitulo_Click(ByVal sender As Object, ByVal e As System.Web.UI.ImageClickEventArgs)
        Try

            Dim imgTitulo As ImageButton = CType(sender, ImageButton)
            Dim row As GridViewRow = CType(imgTitulo.NamingContainer, GridViewRow)
            Dim SqlArray As New ArrayList
            Dim sql As String = ""

            sql = "Delete From NotaFiscalXTitulo  " & vbCrLf &
                  "where Titulo_Id = '" & row.Cells(9).Text & "'" & vbCrLf &
                  "  and Nota_Id = '" & row.Cells(6).Text & "'" & vbCrLf &
                  "  and Serie_Id = '" & row.Cells(5).Text & "'" & vbCrLf &
                  "  and EntradaSaida_Id = '" & row.Cells(7).Text & "'" & vbCrLf &
                  "  and Cliente_Id = '" & row.Cells(3).Text & "'" & vbCrLf &
                  "  and EndCliente_Id = '" & row.Cells(4).Text & "'" & vbCrLf &
                  "  and Empresa_Id = '" & row.Cells(1).Text & "'" & vbCrLf &
                  "  and EndEmpresa_Id = " & row.Cells(2).Text & vbCrLf
            SqlArray.Add(sql)

            If Banco.GravaBanco(SqlArray) Then
                ' Chamar log de observação para desvincular
                Dim tituloId As String = row.Cells(9).Text
                Dim notaId As String = row.Cells(6).Text
                Dim tipo As String = row.Cells(10).Text ' P ou R\
                Dim t = GridTitulo.SelectedIndex
                'Dim tabelaFinanceira As String = IIf(GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(1).Text = "P", "ContasAPagar", "ContasAReceber")
                Dim registroId As String = row.Cells(9).Text

                Desvincular(tituloId, notaId, tipo, registroId)

                MsgBox(Me.Page, "Titulo desvinculado com Sucesso.", eTitulo.Sucess)
            Else
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlUnidadeConsultaTitulos.SelectedIndexChanged
        Try
            ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue, True)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaCliente_Click(sender As Object, e As EventArgs) Handles cmdBuscaCliente.Click
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) Then
                MsgBox(Me.Page, "É necessário informar a empresa!")
                Exit Sub
            End If

            Dim strJScript As String = ""
            HttpContext.Current.Session("ssCampo") = "Livre"
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClienteExtrato" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdBuscaPedido_Click(sender As Object, e As EventArgs) Handles cmdBuscaPedido.Click
        Try
            If String.IsNullOrWhiteSpace(DdlEmpresaConsultaTitulos.SelectedValue) OrElse String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
                MsgBox(Me.Page, "É necessário informar a empresa e o cliente para buscar o número do pedido!")
                Exit Sub
            End If

            BuscarPedido()

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuscarPedido()
        Dim strEmpresa As String() = DdlEmpresaConsultaTitulos.SelectedValue.Split("-")
        Dim strCliente As String() = txtCodigoCliente.Value.Split("-")
        HttpContext.Current.Session("ssCampo" & HID.Value) = "Pedidos"
        Dim parameters As New Dictionary(Of String, Object)
        parameters("unidade") = ""
        parameters("empresa") = strEmpresa(0)
        parameters("enderecoEmpresa") = strEmpresa(1)
        parameters("cliente") = IIf(strCliente(0) <> "", strCliente(0), "")
        parameters("enderecoCliente") = IIf(strCliente.Length > 1, strCliente(1), "")
        parameters("situacao") = 1
        Popup.ConsultaDePedidos(Me.Page, "objVincularNotaxTitulo" & HID.Value)
        ucConsultaPedidos.BindGridView(parameters)
    End Sub

    Protected Sub DdlEmpresaConsultaTitulos_SelectedIndexChanged(sender As Object, e As EventArgs) Handles DdlEmpresaConsultaTitulos.SelectedIndexChanged
        Try
            txtCodigoCliente.Value = String.Empty
            txtClientes.Text = String.Empty
            txtPedido.Text = String.Empty
            GridTitulo.DataSource = New List(Of Object)
            GridTitulo.DataBind()
            GridNotaxTitulo.DataSource = New List(Of Object)
            GridNotaxTitulo.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridNotaxTitulo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridNotaxTitulo.SelectedIndexChanged
        Try
            If GridTitulo.SelectedIndex = -1 Then
                MsgBox(Me.Page, "Número de título não selecionado")
            ElseIf GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(2).Text = GridNotaxTitulo.Rows(GridNotaxTitulo.SelectedIndex).Cells(9).Text Then
                MsgBox(Me.Page, "Nota Fiscal selecionada já está com esse número")
            Else
                'For i = 0 To GridNotaxTitulo.Rows.Count - 1
                '    If GridNotaxTitulo.Rows(i).Cells(9).Text = GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(2).Text Then
                '        MsgBox(Me.Page, "Titulo selecionado já está vinculado na Nota Fiscal " & GridNotaxTitulo.Rows(i).Cells(6).Text)
                '        Exit Sub
                '    End If
                'Next
                txtContratoBancario.Text = Server.HtmlEncode(Funcoes.EliminarCaracteresEspeciais(GridTitulo.Rows(GridTitulo.SelectedIndex).Cells(5).Text))
                lnkVincular.Parent.Visible = True
                NumContratoBancario.Visible = True
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If Validar() Then
                CarregarTitulos()
                CarregarTitulosxNotas()
                GridTitulo.SelectedIndex = -1
                GridNotaxTitulo.SelectedIndex = -1
            End If
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

    Protected Sub lnkVincular_Click(sender As Object, e As EventArgs) Handles lnkVincular.Click
        Try
            If Funcoes.VerificaPermissao("VincularNotaxTitulo", "GRAVAR") Then
                Vincular()
                CarregarTitulos()
                CarregarTitulosxNotas()
                GridTitulo.SelectedIndex = -1
                GridNotaxTitulo.SelectedIndex = -1
                txtContratoBancario.Text = String.Empty
                lnkVincular.Parent.Visible = False
                NumContratoBancario.Visible = False
            Else
                MsgBox(Me.Page, "Usuario sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridTitulo_SelectedIndexChanged(sender As Object, e As EventArgs) Handles GridTitulo.SelectedIndexChanged
        Try
            CarregarTitulosxNotas()
            GridNotaxTitulo.SelectedIndex = -1
            txtContratoBancario.Text = String.Empty
            lnkVincular.Parent.Visible = False
            NumContratoBancario.Visible = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcel_Click(sender As Object, e As EventArgs) Handles lnkExcel.Click
        Try
            If GridTitulo.Rows.Count = 0 AndAlso GridNotaxTitulo.Rows.Count = 0 Then
                MsgBox(Me.Page, "Sem informações para gerar Relatório", eTitulo.Info)
            Else
                Relatorio()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "VincularNotaxTitulo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class