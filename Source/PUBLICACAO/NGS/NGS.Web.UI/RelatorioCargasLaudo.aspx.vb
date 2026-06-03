Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports System.IO
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class RelatorioCargasLaudo
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioCargasLaudo", "ACESSAR") Then
                    TxtDataInicial.Text = String.Format("01/{0}/{1}", DateTime.Now.Month.ToString().PadLeft(2, "0"), DateTime.Now.Year)
                    TxtDataFinal.Text = Format(Today, "dd/MM/yyyy")
                    ddl.Carregar(DdlUnidadeDeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
                    Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocio, DdlEmpresaCliente)
                    ddl.Carregar(DdlClasseOperacao, CarregarDDL.Tabela.ClasseDeOperacao, "Operacao = 1 ", True)
                    ddl.Carregar(ddlClasseSuboperacao, CarregarDDL.Tabela.ClasseDeOperacao, "subOperacao = 1 ", True)
                    ucConsultaClientes.SetarHID(HID.Value)
                    LiberaEmpresa()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlUnidadeDeNegocio_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregaEmpresa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub CarregaEmpresa()
        ddl.Carregar(DdlEmpresaCliente, CarregarDDL.Tabela.Empresas, DdlUnidadeDeNegocio.SelectedValue, True)
    End Sub

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Not Session("objCliente" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Private Sub Limpar()
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeDeNegocio, DdlEmpresaCliente)
        chkConsolidarCliente.Checked = False
        chkConsolidarEmpresa.Checked = False
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        DdlClasseOperacao.SelectedIndex = 0
        TxtDataInicial.Text = String.Format("01/{0}/{1}", DateTime.Now.Month.ToString().PadLeft(2, "0"), DateTime.Now.Year)
        TxtDataFinal.Text = Format(Today, "dd/MM/yyyy")
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocio.Enabled = False
            DdlEmpresaCliente.Enabled = False
        End If
    End Sub

    Function getDataSet() As DataSet
        Dim sql As String

        sql = "SELECT Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Clientes.Reduzido, Pesagem.Produto, Produtos.Descricao, Pesagem.Movimento," & vbCrLf &
              "       ISNULL((SELECT SUM(Liquido) AS Entradas " & vbCrLf &
              "                 FROM Pesagem AS PesagemEntrada " & vbCrLf &
              "                WHERE Sequencia_Id   = 0" & vbCrLf &
              "                  And Situacao       = 1" & vbCrLf &
              "                  AND EntradaSaida   = 'E'" & vbCrLf &
              "                  AND SegundaPesagem > 0 " & vbCrLf &
              "                  AND Empresa_Id     = Pesagem.Empresa_Id" & vbCrLf &
              "                  AND EndEmpresa_Id  = Pesagem.EndEmpresa_Id" & vbCrLf &
              "                  AND Produto        = Pesagem.Produto" & vbCrLf &
              "                  AND Movimento      = Pesagem.Movimento" & vbCrLf &
              "             ), 0) AS Entradas," & vbCrLf &
              "       ISNULL((SELECT SUM(Liquido) AS Saidas " & vbCrLf &
              "                 FROM Pesagem AS PesagemSaida " & vbCrLf &
              "                WHERE Sequencia_Id   = 0" & vbCrLf &
              "                  And Situacao       = 1" & vbCrLf &
              "                  AND EntradaSaida   = 'S'" & vbCrLf &
              "                  AND SegundaPesagem > 0" & vbCrLf &
              "                  AND Empresa_Id     = Pesagem.Empresa_Id" & vbCrLf &
              "                  AND EndEmpresa_Id  = Pesagem.EndEmpresa_Id" & vbCrLf &
              "                  AND Produto        = Pesagem.Produto" & vbCrLf &
              "                  AND Movimento      = Pesagem.Movimento" & vbCrLf &
              "              ), 0) AS Saidas " & vbCrLf &
              "  FROM Pesagem" & vbCrLf &
              " INNER JOIN Clientes" & vbCrLf &
              "    ON Pesagem.Empresa_Id    = Clientes.Cliente_Id" & vbCrLf &
              "   AND Pesagem.EndEmpresa_Id = Clientes.Endereco_Id" & vbCrLf &
              " INNER JOIN Produtos" & vbCrLf &
              "    ON Pesagem.Produto = Produtos.Produto_Id " & vbCrLf

        If DdlClasseOperacao.SelectedIndex > 0 Or ddlClasseSuboperacao.SelectedIndex > 0 Then
            sql &= " INNER JOIN Pedidos" & vbCrLf &
                   "    ON Pesagem.Empresa_Id    = Pedidos.Empresa_Id" & vbCrLf &
                   "   AND Pesagem.EndEmpresa_Id = Pedidos.EndEmpresa_Id" & vbCrLf &
                   "   AND Pesagem.Pedido        = Pedidos.Pedido_Id" & vbCrLf
            If DdlClasseOperacao.SelectedIndex > 0 Then
                sql &= " INNER JOIN Operacoes" & vbCrLf &
                       "    ON Pedidos.Operacao = Operacoes.Operacao_Id" & vbCrLf &
                       "   AND Operacoes.Classe ='" & DdlClasseOperacao.SelectedValue & "'"
            End If
            If ddlClasseSuboperacao.SelectedIndex > 0 Then
                sql &= " INNER JOIN SubOperacoes" & vbCrLf &
                       "    ON Pedidos.Operacao    = SubOperacoes.Operacao_Id" & vbCrLf &
                       "   AND Pedidos.SubOperacao = SubOperacoes.subOperacoes_Id" & vbCrLf &
                       "   AND SubOperacoes.Classe ='" & ddlClasseSuboperacao.SelectedValue & "'"
            End If
        End If

        sql &= " WHERE Pesagem.Sequencia_Id = 0" & vbCrLf &
               "   AND Pesagem.Situacao     = 1" & vbCrLf &
               "   AND Movimento BETWEEN '" & TxtDataInicial.Text.ToSqlDate() & "' AND '" & TxtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

        If Not String.IsNullOrWhiteSpace(DdlEmpresaCliente.SelectedValue) Then
            If chkConsolidarEmpresa.Checked Then
                sql &= " AND Left(Pesagem.Empresa_Id, 8) = '" & Left(DdlEmpresaCliente.SelectedValue.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= " AND Pesagem.Empresa_Id = '" & DdlEmpresaCliente.SelectedValue.Split("-")(0) & "'" & vbCrLf &
                       " AND Pesagem.EndEmpresa_Id = " & DdlEmpresaCliente.SelectedValue.Split("-")(1)
            End If
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            If chkConsolidarCliente.Checked Then
                sql &= " AND Left(Pesagem.Cliente, 8) = '" & Left(txtCodigoCliente.Value.Split("-")(0), 8) & "'" & vbCrLf
            Else
                sql &= " AND Pesagem.Cliente = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf &
                       " AND Pesagem.EndCliente = " & txtCodigoCliente.Value.Split("-")(1)
            End If
        End If

        If ucSelecaoProduto.TemSelecionado Then
            Dim ret As New ArrayList
            ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("Pesagem.Produto", "Pesagem.Produto", "", True)
            sql &= " and " & ret(0) & vbCrLf
        End If

        sql &= " GROUP BY Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id, Clientes.Nome, Clientes.Cidade, Clientes.Estado, Clientes.Reduzido, Pesagem.Produto, Produtos.Descricao, Pesagem.Movimento " & vbCrLf &
               " ORDER BY Pesagem.Produto, Pesagem.Empresa_Id, Pesagem.EndEmpresa_Id, Pesagem.Movimento" & vbCrLf

        Return Banco.ConsultaDataSet(sql, "CargasLaudo")
    End Function

    Private Function getParam() As String
        Dim param As String = ""
        If Not String.IsNullOrWhiteSpace(DdlUnidadeDeNegocio.SelectedValue) Then
            param &= "Unidade: " & DdlUnidadeDeNegocio.SelectedItem.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(DdlEmpresaCliente.SelectedValue) Then
            Dim objEmpresa = New Cliente(DdlEmpresaCliente.SelectedValue.Split("-")(0), DdlEmpresaCliente.SelectedValue.Split("-")(1))
            param &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(TxtDataInicial.Text) Then
            param &= "Período: " & TxtDataInicial.Text & " á: " & TxtDataFinal.Text & vbCrLf
        End If
        If ucSelecaoProduto.TemSelecionado Then
            Dim ret As New ArrayList
            ret = ucSelecaoProduto.GetSqlEParametrosRelatorio("Pesagem.Produto", "Pesagem.Produto", "", True)
            param &= "" & ret(1) & vbCrLf
        End If

        Return param
    End Function

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            Dim parameters = New Dictionary(Of String, Object)()
            Dim ds As DataSet = getDataSet()

            If Funcoes.VerificaPermissao("RelatorioCargasLaudo", "RELATORIO") Then
                parameters.Add("ParametersConsulta", getParam())

                Funcoes.BindReport(Me.Page, ds, "Cr_CargasLaudo", eExportType.PDF, parameters)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkHtml_Click(sender As Object, e As EventArgs) Handles lnkHtml.Click
        Try
            Dim saldo As Integer

            If CDate(TxtDataInicial.Text) > CDate(TxtDataFinal.Text) Then
                MsgBox(Me.Page, "Data inicial não pode ser maior que a data final.")
                TxtDataInicial.Focus()
            Else
                Dim Ds As DataSet = getDataSet()
                If Ds Is Nothing Then
                    MsgBox(Me.Page, "Período sem movimento.")
                ElseIf Ds.Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Período sem movimento.")
                Else
                    Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
                    Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
                    Dim linha As String

                    Dim strm As StreamWriter = Nothing
                    If Dir(arquivo).Length > 0 Then Kill(arquivo)

                    linha = "<!DOCTYPE HTML PUBLIC '-//W3C//DTD HTML 4.0 Transitional//EN'>" & vbCrLf & _
                            "<HTML>" & vbCrLf
                    '<HEAD>
                    linha = "<HEAD>" & vbCrLf & _
                            "<meta http-equiv='Content-Type' content='text/html; charset=utf-8'>" & vbCrLf & _
                            "<link href='" & HttpContext.Current.Server.MapPath(Request.Cookies("empresa").Value) & ".css' type='text/css' rel='stylesheet' runat='server'>" & vbCrLf & _
                            "<TITLE>" & "Totalizador Diário de Cargas e Descargas de " & TxtDataInicial.Text & " à " & TxtDataFinal.Text & "</TITLE>" & vbCrLf & _
                            "</HEAD>" & vbCrLf

                    '<BODY>
                    linha &= "<BODY>" & vbCrLf

                    '-----------------
                    'Cabeçalho Padrao
                    '-----------------
                    linha &= "<table width= '100%' cellpadding='0' cellspacing='0' Border=1>" & vbCrLf

                    linha = "<TR>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>Produto</A></TD>" & vbCrLf & _
                            "<TD style='width: 200px;'><A Class='MsSerif9b'>Descrição</A></TD>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>Empresa</A></TD>" & vbCrLf & _
                            "<TD style='width: 300px;'><A Class='MsSerif9b'>Nome</A></TD>" & vbCrLf & _
                            "<TD style='width: 200px;'><A Class='MsSerif9b'>Cidade</A></TD>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>UF</A></TD>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>Movimento</A></TD>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>Entradas</A></TD>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>Saídas</A></TD>" & vbCrLf & _
                            "<TD><A Class='MsSerif9b'>Saldo</A></TD>" & vbCrLf & _
                            "</TR>" & vbCrLf

                    For Each drCargas As DataRow In Ds.Tables(0).Rows
                        linha = "<TR>" & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & drCargas("Produto") & "</A></TD>" & vbCrLf & _
                                "<TD style='width: 200px;'><A Class='MsSerif8b'>" & drCargas("Descricao") & "</A></TD>" & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & drCargas("Empresa_Id") & "</A></TD>" & vbCrLf & _
                                "<TD style='width: 300px;'><A Class='MsSerif8b'>" & drCargas("Nome") & "</A></TD>" & vbCrLf & _
                                "<TD style='width: 200px;'><A Class='MsSerif8b'>" & drCargas("Cidade") & "</A></TD>" & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & drCargas("Estado") & "</A></TD>" & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & drCargas("Movimento") & "</A></TD>" & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & drCargas("Entradas") & "</A></TD>" & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & drCargas("Saidas") & "</A></TD>" & vbCrLf & _
                                drCargas("Entradas") - drCargas("Saidas") & vbCrLf & _
                                "<TD><A Class='MsSerif8b'>" & saldo & "</A></TD>" & vbCrLf & _
                                "</TR>" & vbCrLf
                    Next

                    linha = "</table>" & vbCrLf & _
                            "</BODY>" & vbCrLf & _
                            "</HTML>" & vbCrLf

                    Try
                        strm = New StreamWriter(arquivo, True)
                        strm.WriteLine(linha)
                        strm.Close()
                        ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString, "window.open('" & NomeArquivo & "');", True)
                    Finally
                        strm.Close()
                    End Try
                End If
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

    Protected Sub btnConsultarCliente_Click(sender As Object, e As EventArgs) Handles btnConsultarCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objCliente" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioCargasLaudo")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class