Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class LancamentosContabeisXCustos
    Inherits BasePage

    Dim Sql As String = ""
    Dim Sqla As String = ""

    Dim SqlArray As New ArrayList
    Dim Empresa() As String = Nothing
    Dim Cliente() As String = Nothing
    Dim SequenciaLote As Integer = 0

    Dim Condicao As String = ""
    Dim Codigo As String = ""
    Dim Endereco As Integer = 0
    Dim Mensagem As String = ""


    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Custos)
        If Not IsPostBack And IsConnect Then

            If Not UsuarioServidor.KeyCodeActive Then
                MsgBox(Me.Page, "Sistema com chave de licença expirada. Entre em contato com o suporte.", "~/ApuracaoDeCustos.aspx", eTitulo.Info)
                Exit Sub
            End If

            If Funcoes.VerificaPermissao("LancamentosContabeisXCustos", "ACESSAR") Then
                CarregarUnidade()
                CarregarHistorico()
                CarregarLote()
                CarregarCusto()
                CarregarGrupo()
                CarregarCampos()
                Limpar()
                VerificaUnidade()
                HID.Value = Guid.NewGuid().ToString
                ucConsultaClientes.SetarHID(HID.Value)
                ucConsultaPlanoDeContas.SetarHID(HID.Value)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/ApuracaoDeCustos.aspx")
                Exit Sub
            End If
        End If

    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            ddlUnidade.SelectedValue = Dr("AcessoUnidade")
            EmpresasPorUnidade()
            ddlEmpresa.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
        Next
    End Sub

#Region "Funções e Sub-rotinas"

    Private Sub CarregarCampos()
        txtContaDebito.Text = HttpContext.Current.Session("Debito_Conta") & " - " & HttpContext.Current.Session("Debito_Titulo")
        txtContaCredito.Text = HttpContext.Current.Session("Credito_Conta") & " - " & HttpContext.Current.Session("Credito_Titulo")
        If Not String.IsNullOrEmpty(HttpContext.Current.Session("Credito_Cliente")) AndAlso Not String.IsNullOrEmpty(HttpContext.Current.Session("Credito_Cliente")) _
             AndAlso Not String.IsNullOrEmpty(HttpContext.Current.Session("Credito_Cliente")) AndAlso Not String.IsNullOrEmpty(HttpContext.Current.Session("Credito_Cliente")) Then
            txtClienteDebito.Text = Funcoes.FormatarCpfCnpj(HttpContext.Current.Session("Credito_Conta")) & " - " & HttpContext.Current.Session("Credito_Cliente")
            txtClienteCredito.Text = Funcoes.FormatarCpfCnpj(HttpContext.Current.Session("Credito_Cliente")) & " - " & HttpContext.Current.Session("Credito_Cliente")
        End If
    End Sub

    Private Sub CarregarUnidade()
        Sql = "SELECT C.Cliente_Id as Codigo, C.Nome as Descricao " & vbCrLf & _
              "FROM Clientes C " & vbCrLf & _
              "INNER JOIN ClientesXTipos CT " & vbCrLf & _
              "ON C.Cliente_Id = CT.Cliente_Id " & vbCrLf & _
              "WHERE CT.Tipo_Id = 050 " & vbCrLf & _
              "ORDER BY Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlUnidade.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        ddlUnidade.Items.Insert(0, "")
        ddlUnidade.SelectedIndex = 0
    End Sub

    Private Sub CarregarHistorico()
        Sql = "SELECT  Descricao " & vbCrLf & _
              " FROM Historicos Order by Descricao" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlHistorico.Items.Add(New ListItem(Dr("Descricao"), Dr("Descricao")))
        Next

        ddlHistorico.Items.Insert(0, "")
        ddlHistorico.SelectedIndex = 0
    End Sub

    Private Sub CarregarLote()
        ddlLote.Items.Clear()
        Sql = "SELECT     Lotes.Lote_Id as Codigo, convert(varchar,Lotes.Lote_Id) + '-' + Lotes.Descricao  as Descricao " & vbCrLf & _
              " FROM     Lotes RIGHT OUTER JOIN " & vbCrLf & _
              " Sistemas ON Lotes.Sistema_Id = Sistemas.Sistema_Id " & vbCrLf & _
              " Where Sistemas.Sistema_Id = 2" & vbCrLf

        ddlLote.DataValueField = "Codigo"
        ddlLote.DataTextField = "Descricao"
        ddlLote.DataSource = Banco.ConsultaDataSet(Sql, "Lotes")
        ddlLote.DataBind()
        ddlLote.Items.Insert(0, "")
        ddlLote.ClearSelection()
        ddlLote.SelectedIndex = 1
    End Sub

    Private Sub CarregarCusto()
        ddlCustoDebito.Items.Clear()
        Sql = "SELECT CentroDeCusto_Id as Codigo, convert(varchar,REPLICATE('0', 5 - LEN(CAST(CentroDeCusto_Id AS varchar))) + CAST(CentroDeCusto_Id AS varchar)) + '-' + Descricao as Descricao " & vbCrLf & _
              " FROM CentrosDeCustos " & vbCrLf & _
              " WHERE Ativo = 1 and len(CentroDeCusto_ID) = 5 " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlCustoDebito.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
            ddlCustoCredito.Items.Add(New ListItem(Dr("Descricao"), Dr("Codigo")))
        Next

        ddlCustoDebito.Items.Insert(0, "")
        ddlCustoDebito.SelectedIndex = 0
        ddlCustoCredito.Items.Insert(0, "")
        ddlCustoCredito.SelectedIndex = 0
    End Sub

    Private Sub CarregarGrupo()
        DdlGrupo.Items.Clear()

        Sql = " SELECT     distinct GruposDeEstoques.Grupo_Id, GruposDeEstoques.Descricao" & vbCrLf & _
              " FROM         GruposDeEstoques INNER JOIN" & vbCrLf & _
              "              Produtos ON GruposDeEstoques.Grupo_Id = Produtos.Grupo" & vbCrLf & _
              " WHERE     (LEN(GruposDeEstoques.Grupo_Id) = 5) order by GruposDeEstoques.Descricao" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            DdlGrupo.Items.Add(New ListItem(Dr("Grupo_Id") & "-" & Dr("Descricao"), Dr("Grupo_Id")))
        Next

        DdlGrupo.Items.Insert(0, "")
        DdlGrupo.SelectedIndex = 0
    End Sub

    Private Sub CarregarProduto()
        ddlProduto.Items.Clear()

        Sql = "SELECT Produto_Id, Nome FROM Produtos WHERE Grupo = '" & DdlGrupo.SelectedValue & "'" & vbCrLf & _
              " Order by Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            ddlProduto.Items.Add(New ListItem(Dr("Produto_Id") & "-" & Dr("Nome"), Dr("Produto_Id")))
        Next

        ddlProduto.Items.Insert(0, "")
        ddlProduto.SelectedIndex = 0
    End Sub

    Private Sub CarregarEmpresa()
        Dim SMgr As ScriptManager = ScriptManager.GetCurrent(Page)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String

        Dim Cnpj As String

        ddlEmpresa.Items.Clear()

        Sql = "SELECT     Clientes.Cliente_Id AS Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM         Clientes INNER JOIN " & vbCrLf & _
              " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id INNER JOIN " & vbCrLf & _
              " GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id " & vbCrLf & _
              " WHERE     (ClientesXTipos.Tipo_Id = 001) and (GruposXEmpresas.Empresa_Id=" & ddlUnidade.SelectedValue & ")" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Empresas").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Dr("Reduzido") & " - " & Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj
            ddlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next
        If ddlEmpresa.Items.Count = 1 Then
            ddlEmpresa.SelectedIndex = 0
            'CarregarPlanoDeContas("D")
        Else
            ddlEmpresa.Items.Insert(0, "")
            ddlEmpresa.ClearSelection()
            ddlEmpresa.Focus()
        End If
    End Sub

    Private Sub Limpar()
        txtContaDebito.Text = ""
        txtClienteDebito.Text = ""
        ddlCustoDebito.SelectedIndex = 0

        txtContaCredito.Text = ""
        txtClienteCredito.Text = ""
        ddlCustoCredito.SelectedIndex = 0

        ddlHistorico.SelectedIndex = 0
        txtValor.Text = ""
        txtComplemento.Text = ""

        DdlGrupo.SelectedIndex = 0
        ddlProduto.Items.Clear()

        txtMovimento.Text = Format(Today, "dd/MM/yyyy")

        HttpContext.Current.Session("Debito_Conta") = ""
        HttpContext.Current.Session("Debito_Titulo") = ""
        HttpContext.Current.Session("Debito_TemCliente") = ""
        HttpContext.Current.Session("Debito_TipoDeCliente") = 0
        HttpContext.Current.Session("Debito_TemProduto") = ""
        HttpContext.Current.Session("Debito_TemCusto") = ""

        HttpContext.Current.Session("Credito_Conta") = ""
        HttpContext.Current.Session("Credito_Titulo") = ""
        HttpContext.Current.Session("Credito_TemCliente") = ""
        HttpContext.Current.Session("Credito_TipoDeCliente") = 0
        HttpContext.Current.Session("Credito_TemProduto") = ""
        HttpContext.Current.Session("Credito_TemCusto") = ""

        HttpContext.Current.Session("Debito_Cliente") = ""
        HttpContext.Current.Session("Debito_ClienteEnd") = 0
        HttpContext.Current.Session("Debito_ClienteNome") = ""
        HttpContext.Current.Session("Debito_ClienteCidade") = ""
        HttpContext.Current.Session("debito_ClienteEstado") = ""

        HttpContext.Current.Session("Credito_Cliente") = ""
        HttpContext.Current.Session("Credito_ClienteEnd") = 0
        HttpContext.Current.Session("Credito_ClienteNome") = ""
        HttpContext.Current.Session("Credito_ClienteCidade") = ""
        HttpContext.Current.Session("Credito_ClienteEstado") = ""

        btnContasCredito.Enabled = True
        btnClientesCredito.Enabled = True
        ddlCustoCredito.Enabled = True

        btnContasDebito.Enabled = True
        btnClientesDebito.Enabled = True
        ddlCustoDebito.Enabled = True
        If ddlUnidade.SelectedIndex = 0 Then
            VerificaUnidade()
        End If

        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidade.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Private Sub TotalizaLote()
        txtTotalDebitos.Text = ""
        txtTotalCreditos.Text = ""
        txtDiferenca.Text = ""

        If ValidarLote() Then
            Empresa = ddlEmpresa.SelectedValue.Split("-")
            Sql &= "'" & Empresa(0) & "',"
            Sql &= Empresa(1) & ","

            Sql = "SELECT Sum(DebitoOficial) As Debitos, Sum(CreditoOficial) as Creditos, Sum(DebitoOficial - CreditoOficial) as Diferenca" & vbCrLf & _
                   " FROM RazaoXCustos" & vbCrLf & _
                   " WHERE Empresa_Id = '" & Empresa(0) & "' AND EndEmpresa_Id = " & vbCrLf & _
                   Empresa(1) & " And Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & vbCrLf & _
                   "' and Lote_Id = " & ddlLote.SelectedValue & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Razao").Tables(0).Rows
                If Not IsDBNull(Dr("Debitos")) Then
                    txtTotalDebitos.Text = CDbl((Dr("Debitos"))).ToString("N2")
                End If
                If Not IsDBNull(Dr("Creditos")) Then
                    txtTotalCreditos.Text = CDbl((Dr("Creditos"))).ToString("N2")
                End If
                If Not IsDBNull(Dr("Diferenca")) Then
                    txtDiferenca.Text = CDbl((Dr("Diferenca"))).ToString("N2")
                    If CDbl((Dr("Diferenca"))) < 0 Then
                        txtDiferenca.ForeColor = Drawing.Color.Red
                    Else
                        txtDiferenca.ForeColor = Drawing.Color.Blue
                    End If
                End If
            Next

            ddlSequencia.Items.Clear()

            Sql = "SELECT Sequencia_Id" & vbCrLf & _
                   " FROM RazaoXCustos" & vbCrLf & _
                   " WHERE Empresa_Id = '" & Empresa(0) & "' AND EndEmpresa_Id = " & vbCrLf & _
                   Empresa(1) & " And Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & vbCrLf & _
                   "' and Lote_Id = " & ddlLote.SelectedValue & vbCrLf & _
                   " Order by Sequencia_Id" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                ddlSequencia.Items.Add(New ListItem(Dr("Sequencia_Id"), Dr("Sequencia_Id")))
            Next

            ddlSequencia.Items.Insert(0, "")
            ddlSequencia.SelectedIndex = 0

            ddlSequencia.Focus()
        End If
    End Sub

    Function ValidarLote()
        Dim ok As Boolean = True

        If ddlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            ok = False
        End If
        If ddlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Empresa é obrigatório.")
            ok = False
        End If
        If txtMovimento.Text = "" Then
            MsgBox(Me.Page, "Data de movimento é obrigatório.")
            ok = False
        End If
        If ddlLote.SelectedValue = "" Then
            MsgBox(Me.Page, "Lote é obrigatório.")
            ok = False
        End If

        Return ok
    End Function

    Function Validar()
        Dim ok As Boolean = True

        If txtMovimento.Text = "" Then
            MsgBox(Me.Page, "Movimento é obrigatório.")
            ok = False
        End If
        If ddlLote.SelectedValue = "" Then
            MsgBox(Me.Page, "Lote é obrigatório.")
            ok = False
        End If
        If ddlUnidade.SelectedValue = "" Then
            MsgBox(Me.Page, "Unidade de negócio é obrigatório.")
            ok = False
        End If

        '*******
        '* Consiste os campos de débito
        '*******
        If ddlEmpresa.SelectedValue = "" Then
            MsgBox(Me.Page, "Informe a empresa.")
            ok = False
        End If
        If HttpContext.Current.Session("Debito_Conta") = "" And HttpContext.Current.Session("Credito_Conta") = "" Then
            MsgBox(Me.Page, "Conta de débito ou crédito é obrigatório.")
        End If
        If HttpContext.Current.Session("Debito_TemCliente") = "S" And HttpContext.Current.Session("Debito_Cliente") = "" Then
            MsgBox(Me.Page, "Selecione o cliente do débito.")
            ok = False
        End If
        If HttpContext.Current.Session("Debito_TemCusto") = "S" And ddlCustoDebito.SelectedValue = "" Then
            MsgBox(Me.Page, "Selecione o centro de custo do débito.")
            ok = False
        End If
        If HttpContext.Current.Session("Debito_TemCusto") <> "S" And ddlCustoDebito.SelectedValue <> "" Then
            MsgBox(Me.Page, "Esta conta de débito não tem centro de custo.")
            ok = False
        End If
        If HttpContext.Current.Session("Debito_TemProduto") = "S" And ddlProduto.SelectedValue = "" Then
            MsgBox(Me.Page, "Selecione o produto para a conta de débito.")
            ok = False
        End If

        '----------------------------------
        '* Consiste os campos de crédito
        '----------------------------------
        If HttpContext.Current.Session("Credito_TemCliente") = "S" And HttpContext.Current.Session("Credito_Cliente") = "" Then
            MsgBox(Me.Page, "Selecione o cliente do crédito.")
            ok = False
        End If
        If HttpContext.Current.Session("Credito_TemCusto") = "S" And ddlCustoCredito.SelectedValue = "" Then
            MsgBox(Me.Page, "Selecione o centro de custo do crédito.")
            ok = False
        End If
        If HttpContext.Current.Session("Credito_TemCusto") <> "S" And ddlCustoCredito.SelectedValue <> "" Then
            MsgBox(Me.Page, "Esta conta de crédito não tem centro de custo.")
            ok = False
        End If
        If HttpContext.Current.Session("Credito_TemProduto") = "S" And ddlProduto.SelectedValue = "" Then
            MsgBox(Me.Page, "Selecione o produto para a conta de crédito.")
            ok = False
        End If

        '----------------------------------

        If txtValor.Text = "" Then
            MsgBox(Me.Page, "Valor é obrigatório.")
            ok = False

        ElseIf CDbl(txtValor.Text) = 0 Then
            MsgBox(Me.Page, "Valor é obrigatório.")
            ok = False
        End If
        If ddlHistorico.SelectedValue = "" And txtComplemento.Text = "" Then
            MsgBox(Me.Page, "Informe o histórico ou o complemento de histórico.")
            ok = False
        End If
        If HttpContext.Current.Session("Debito_TemProduto") = "S" Or HttpContext.Current.Session("Credito_TemProduto") = "S" Then
            If ddlProduto.SelectedValue = "" Then
                MsgBox(Me.Page, "Selecione o produto.")
                ok = False
            End If
        End If
        If HttpContext.Current.Session("Debito_TemProduto") <> "S" And HttpContext.Current.Session("Credito_TemProduto") <> "S" Then
            If ddlProduto.SelectedValue <> "" Then
                DdlGrupo.SelectedIndex = 0
                ddlProduto.SelectedIndex = 0
                ok = False
                MsgBox(Me.Page, "Nenhuma conta selecionada permite produto.")
                ok = False
            End If
        End If

        Return ok
    End Function

    Private Sub Excluir()
        If ValidarLote() Then
            If ddlSequencia.Text <> "" Then
                Empresa = ddlEmpresa.SelectedValue.Split("-")

                Sql = " Delete from RazaoXCustos WHERE Empresa_Id = '" & Empresa(0) & "' AND " & vbCrLf & _
                      " EndEmpresa_Id = " & Empresa(1) & " AND " & vbCrLf & _
                      " Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & "' And " & vbCrLf & _
                      " Lote_Id = " & ddlLote.SelectedValue & " AND " & vbCrLf & _
                      " Sequencia_Id = " & ddlSequencia.SelectedValue & vbCrLf
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    Dim Mensagem As String = Replace(HttpContext.Current.Session("ssMessage"), "'", "")
                    MsgBox(Me.Page, Mensagem, eTitulo.Info)
                End If
            Else
                MsgBox(Me.Page, "Informe a sequencia do lote.")
            End If
        End If
    End Sub

    Private Function Sequencia(ByVal Empresa As String, ByVal EndEmpresa As Integer)
        Dim Sqll As String = String.Empty

        Sqll = "SELECT isnull(Max(Sequencia_Id),0) + 1 as Sequencia" & vbCrLf & _
               " FROM RazaoXCustos" & vbCrLf & _
               " WHERE Empresa_Id = '" & Empresa & "' AND EndEmpresa_Id = " & vbCrLf & _
               EndEmpresa & " And Movimento_Id = '" & txtMovimento.Text.ToSqlDate() & vbCrLf & _
               "' and Lote_Id = " & ddlLote.SelectedValue & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sqll, "Razao").Tables(0).Rows
            Return Dr("Sequencia")
        Next
        Return Nothing
    End Function

    Private Sub Gravar()
        Dim data As String = Format(CDate(txtMovimento.Text), "yyyy/MM/dd")

        If HttpContext.Current.Session("Debito_Conta") <> "" Then
            Sql = "INSERT INTO RazaoXCustos " & vbCrLf & _
                  "(Empresa_Id, " & vbCrLf & _
                  "EndEmpresa_Id, " & vbCrLf & _
                  "Conta_Id, " & vbCrLf & _
                  "Cliente_Id, " & vbCrLf & _
                  "EndCliente_Id, " & vbCrLf & _
                  "Movimento_Id, " & vbCrLf & _
                  "Lote_Id, " & vbCrLf & _
                  "Sequencia_Id, " & vbCrLf & _
                  "UnidadeDeNegocio, " & vbCrLf & _
                  "Titulo, " & vbCrLf & _
                  "Pedido, " & vbCrLf & _
                  "PedidoFixacao, " & vbCrLf & _
                  "Produto, " & vbCrLf & _
                  "Custo, " & vbCrLf & _
                  "Indexador, " & vbCrLf & _
                  "DataMoeda, " & vbCrLf & _
                  "DebitoOficial, " & vbCrLf & _
                  "CreditoOficial, " & vbCrLf & _
                  "DebitoMoeda, " & vbCrLf & _
                  "CreditoMoeda, " & vbCrLf & _
                  "DebitoQuantidade," & vbCrLf & _
                  "CreditoQuantidade," & vbCrLf & _
                  "Historico, " & vbCrLf & _
                  "PrevistoRealizado, " & vbCrLf & _
                  "Cliente_Nf, " & vbCrLf & _
                  "EndCliente_Nf, " & vbCrLf & _
                  "EntradaSaida_Nf, " & vbCrLf & _
                  "Serie_Nf, " & vbCrLf & _
                  "Numero_Nf, " & vbCrLf & _
                  "ChequeEntregue, " & vbCrLf & _
                  "PagamentoAutorizado, " & vbCrLf & _
                  "Processo) " & vbCrLf & _
                  "VALUES (" & vbCrLf

            Empresa = ddlEmpresa.SelectedValue.Split("-")

            Sql &= "'" & Empresa(0) & "',"                                                      'CGC EmpresaDebito
            Sql &= Empresa(1) & ","                                                             'Endereco EmpresaDebito
            Sql &= "'" & HttpContext.Current.Session("Debito_Conta") & "',"                     'ContaDebito
            Sql &= "'" & HttpContext.Current.Session("Debito_Cliente") & "',"                   'ClienteDebito
            Sql &= HttpContext.Current.Session("Debito_ClienteEnd") & ","                       'Endereco ClienteDebito

            Sql &= "'" & data & "',"                                                            'Data de movimento
            Sql &= ddlLote.SelectedValue & ","                                                  'Lote

            If ddlSequencia.SelectedValue <> "" Then
                SequenciaLote = ddlSequencia.SelectedValue
            Else
                SequenciaLote = Sequencia(Empresa(0), Empresa(1))
            End If

            Sql &= SequenciaLote & ","                                                          'Sequencia

            Sql &= "'" & ddlUnidade.SelectedValue & "',"                                        'Unidade de negocio Debito
            Sql &= "0,'','',"                                                                   'Titulo, pedido, fixacao

            If HttpContext.Current.Session("Debito_TemProduto") = "S" Then
                Sql &= "'" & ddlProduto.SelectedValue & "',"                                    'Produto Debito
            Else
                Sql &= "'',"
            End If

            If HttpContext.Current.Session("Debito_TemCusto") = "S" Then
                Sql &= ddlCustoDebito.SelectedValue & ","                                       'Codigo de custo debito
            Else
                Sql &= "0,"
            End If

            Sql &= "0,"                                                                         'Indexador
            Sql &= "'" & data & "',"                                                            'Data moeda
            Sql &= Replace(Replace(txtValor.Text, ".", ""), ",", ".") & ","                     'Debito Oficial
            Sql &= "0,0,0,"                                                                     'Credito Oficial, DebitoMoeda, CreditoMoeda

            If Not IsNumeric(txtQuantidade.Text) Then txtQuantidade.Text = "0"
            Sql &= Str(txtQuantidade.Text) & ",0,"   'Quantidade Debito, Quantidade Credito

            If ddlHistorico.SelectedValue <> "" Then
                Sql &= "'" & RTrim(ddlHistorico.SelectedValue) & " " & RTrim(txtComplemento.Text) & "',"  'Historico
            Else
                Sql &= "'" & RTrim(txtComplemento.Text) & "',"  'Historico
            End If

            Sql &= "'P','',0,'','',0,'','','CONTABIL')"

            SqlArray.Add(Sql)

        End If

        If HttpContext.Current.Session("Credito_Conta") <> "" Then

            Sql = "INSERT INTO RazaoXCustos " & vbCrLf & _
                  "(Empresa_Id, " & vbCrLf & _
                  "EndEmpresa_Id, " & vbCrLf & _
                  "Conta_Id, " & vbCrLf & _
                  "Cliente_Id, " & vbCrLf & _
                  "EndCliente_Id, " & vbCrLf & _
                  "Movimento_Id, " & vbCrLf & _
                  "Lote_Id, " & vbCrLf & _
                  "Sequencia_Id, " & vbCrLf & _
                  "UnidadeDeNegocio, " & vbCrLf & _
                  "Titulo, " & vbCrLf & _
                  "Pedido, " & vbCrLf & _
                  "PedidoFixacao, " & vbCrLf & _
                  "Produto, " & vbCrLf & _
                  "Custo, " & vbCrLf & _
                  "Indexador, " & vbCrLf & _
                  "DataMoeda, " & vbCrLf & _
                  "DebitoOficial, " & vbCrLf & _
                  "CreditoOficial, " & vbCrLf & _
                  "DebitoMoeda, " & vbCrLf & _
                  "CreditoMoeda, " & vbCrLf & _
                  "DebitoQuantidade," & vbCrLf & _
                  "CreditoQuantidade," & vbCrLf & _
                  "Historico, " & vbCrLf & _
                  "PrevistoRealizado, " & vbCrLf & _
                  "Cliente_Nf, " & vbCrLf & _
                  "EndCliente_Nf, " & vbCrLf & _
                  "EntradaSaida_Nf, " & vbCrLf & _
                  "Serie_Nf, " & vbCrLf & _
                  "Numero_Nf, " & vbCrLf & _
                  "ChequeEntregue, " & vbCrLf & _
                  "PagamentoAutorizado, " & vbCrLf & _
                  "Processo) " & vbCrLf & _
                  "VALUES (" & vbCrLf

            Empresa = ddlEmpresa.SelectedValue.Split("-")

            Sql &= "'" & Empresa(0) & "',"                                                      'CGC EmpresaDebito
            Sql &= Empresa(1) & ","                                                             'Endereco EmpresaDebito
            Sql &= "'" & HttpContext.Current.Session("Credito_Conta") & "',"                    'ContaDebito
            Sql &= "'" & HttpContext.Current.Session("Credito_Cliente") & "',"                  'ClienteDebito
            Sql &= HttpContext.Current.Session("Credito_ClienteEnd") & ","                      'Endereco ClienteDebito

            Sql &= "'" & data & "',"                                                            'Data de movimento
            Sql &= ddlLote.SelectedValue & ","                                                  'Lote

            If HttpContext.Current.Session("Debito_Conta") <> "" Then
                SequenciaLote += 1
            Else
                If ddlSequencia.SelectedValue <> "" Then
                    SequenciaLote = ddlSequencia.SelectedValue
                Else
                    SequenciaLote = Sequencia(Empresa(0), Empresa(1))
                End If
            End If

            Sql &= SequenciaLote & ","                                                          'Sequencia

            Sql &= "'" & ddlUnidade.SelectedValue & "',"                                        'Unidade de negocio Debito
            Sql &= "0,'','',"                                                                   'Titulo, pedido, fixacao

            If HttpContext.Current.Session("Credito_TemProduto") = "S" Then
                Sql &= "'" & ddlProduto.SelectedValue & "',"                                    'Produto Debito
            Else
                Sql &= "'',"
            End If

            If HttpContext.Current.Session("Credito_TemCusto") = "S" Then
                Sql &= ddlCustoCredito.SelectedValue & ","                                       'Codigo de custo debito
            Else
                Sql &= "0,"
            End If

            Sql &= "0,"                                                                         'Indexador
            Sql &= "'" & data & "',"                                                            'Data moeda
            Sql &= "0,"                                                                         'Indexador
            Sql &= Replace(Replace(txtValor.Text, ".", ""), ",", ".") & ","                     'Credito Oficial
            Sql &= "0,0,"                                                                       'Credito Oficial, DebitoMoeda, CreditoMoeda

            If Not IsNumeric(txtQuantidade.Text) Then txtQuantidade.Text = "0"
            Sql &= "0," & Str(txtQuantidade.Text) & ","   'Quantidade Debito, Quantidade Credito

            If ddlHistorico.SelectedValue <> "" Then
                Sql &= "'" & RTrim(ddlHistorico.SelectedValue) & " " & RTrim(txtComplemento.Text) & "',"  'Historico
            Else
                Sql &= "'" & RTrim(txtComplemento.Text) & "',"  'Historico
            End If

            Sql &= "'P','',0,'','',0,'','','CONTABIL')"

            SqlArray.Add(Sql)

        End If

        If Banco.GravaBanco(SqlArray) Then
            MsgBox(Me.Page, "Registro gravado com Sucesso.", eTitulo.Sucess)
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If
            Limpar()
            TotalizaLote()
        End If
    End Sub

#End Region

    Protected Sub ddlUnidade_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlUnidade.SelectedIndexChanged
        Try
            Limpar()
            ddlLote.SelectedIndex = 1
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If
            EmpresasPorUnidade()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub EmpresasPorUnidade()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        ddlEmpresa.Items.Clear()

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & ddlUnidade.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))
            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")
            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")
            ddlEmpresa.Items.Add(New ListItem(Descricao, Codigo))
        Next

        ddlEmpresa.Items.Insert(0, "")
        ddlEmpresa.SelectedIndex = 0
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ddlEmpresa.SelectedIndexChanged
        Try
            Limpar()

            ddlLote.SelectedIndex = 0
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnContasDebito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContasDebito.Click
        Try
            HttpContext.Current.Session("ssCampo") = "Lancamentos_Debito"
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me.Page, "objContaDebito" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnContasCredito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnContasCredito.Click
        Try
            HttpContext.Current.Session("ssCampo") = "Lancamentos_Credito"
            ucConsultaPlanoDeContas.Limpar()
            ucConsultaPlanoDeContas.BindGridView(True)
            Popup.ConsultaDePlanoDeContas(Me.Page, "objContaCredito" & HID.Value)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClientesDebito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClientesDebito.Click
        Try
            If HttpContext.Current.Session("Debito_TemCliente") = "S" Then
                HttpContext.Current.Session("ssCampo") = "Lancamentos_Debito"
                ucConsultaClientes.Limpar()
                Popup.ConsultaDeClientes(Me.Page, "objClienteLcXcDeb" & HID.Value, "txtNome")
            Else
                MsgBox(Me.Page, "Conta nao permite cliente.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnClientesCredito_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnClientesCredito.Click
        Try
            If HttpContext.Current.Session("Credito_TemCliente") = "S" Then
                HttpContext.Current.Session("ssCampo") = "Lancamentos_Credito"
                ucConsultaClientes.Limpar()
                Popup.ConsultaDeClientes(Me.Page, "objClienteLcXcCred" & HID.Value, "txtNome")
            Else
                MsgBox(Me.Page, "Conta não permite cliente.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub DdlGrupo_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles DdlGrupo.SelectedIndexChanged
        Try
            If HttpContext.Current.Session("Debito_TemProduto") <> "S" And HttpContext.Current.Session("Credito_TemProduto") <> "S" Then
                DdlGrupo.SelectedIndex = 0
                MsgBox(Me.Page, "Nenhuma conta selecionada permite produto.")
            Else
                CarregarProduto()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlLote_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If
            TotalizaLote()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Consultar()
        Limpar()

        If ValidarLote() Then
            If ddlSequencia.Text = "" Then
                MsgBox(Me.Page, "Sequência do lote é obrigatório.")
            Else
                Empresa = ddlEmpresa.SelectedValue.Split("-")
                Sql &= "'" & Empresa(0) & "'," & vbCrLf & _
                       Empresa(1) & "," & vbCrLf

                Sql = "SELECT Conta_Id, Cliente_Id, EndCliente_Id, Custo, Produto, DebitoOficial, CreditoOficial, isnull(DebitoQuantidade,0) as DebitoQuantidade, isnull(CreditoQuantidade,0) as CreditoQuantidade, Historico  " & vbCrLf & _
                      " FROM RazaoXCustos" & vbCrLf & _
                      " WHERE Empresa_Id    ='" & Empresa(0) & "'" & vbCrLf & _
                      "   AND EndEmpresa_Id = " & Empresa(1) & vbCrLf & _
                      "   And Movimento_Id  ='" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
                      "   AND Lote_Id       = " & ddlLote.SelectedValue & vbCrLf & _
                      "   AND Sequencia_Id  = " & ddlSequencia.SelectedValue & vbCrLf

                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Razao").Tables(0).Rows
                    If Dr("DebitoOficial") <> 0 Then
                        txtValor.Text = Dr("DebitoOficial")
                        txtQuantidade.Text = Dr("DebitoQuantidade")
                        Codigo = Dr("Conta_Id")
                        Condicao = "Debito"
                        ConsultaConta()

                        If Dr("Cliente_Id") <> "" Then
                            Codigo = Dr("Cliente_Id")
                            Endereco = Dr("EndCliente_Id")
                            ConsultaCliente()
                        End If

                        If Dr("Custo") <> 0 Then
                            ddlCustoDebito.SelectedValue = Dr("Custo")
                        End If

                        If Dr("Produto") <> "" Then
                            Codigo = Dr("Produto")
                            ConsultaProduto()
                        End If

                        btnContasCredito.Enabled = False
                        btnClientesCredito.Enabled = False
                        ddlCustoCredito.Enabled = False
                    Else
                        txtValor.Text = Dr("CreditoOficial")
                        txtQuantidade.Text = Dr("CreditoQuantidade")
                        Codigo = Dr("Conta_Id")
                        Condicao = "Credito"
                        ConsultaConta()

                        If Dr("Cliente_Id") <> "" Then
                            Codigo = Dr("Cliente_Id")
                            Endereco = Dr("EndCliente_Id")
                            ConsultaCliente()
                        End If

                        If Dr("Custo") <> 0 Then
                            ddlCustoCredito.SelectedValue = Dr("Custo")
                        End If

                        If Dr("Produto") <> "" Then
                            Codigo = Dr("Produto")
                            ConsultaProduto()
                        End If

                        btnContasDebito.Enabled = False
                        btnClientesDebito.Enabled = False
                        ddlCustoDebito.Enabled = False
                    End If
                    txtComplemento.Text = Dr("Historico")
                Next
            End If
        End If
    End Sub

    Private Sub ConsultaConta()
        Sqla = "SELECT Conta_Id, Titulo, Cliente, TipoDeCliente, Produto, CentroDeCusto  " & vbCrLf & _
               "  FROM PlanoDeContas " & vbCrLf & _
               " Where conta_id = '" & Codigo & "'" & vbCrLf

        For Each Dra As DataRow In Banco.ConsultaDataSet(Sqla, "Romaneios").Tables(0).Rows
            If Condicao = "Debito" Then
                HttpContext.Current.Session("Debito_Conta") = Dra("Conta_Id")
                HttpContext.Current.Session("Debito_titulo") = Dra("Titulo")
                HttpContext.Current.Session("Debito_TemCliente") = Dra("Cliente")
                HttpContext.Current.Session("Debito_TipoDeCliente") = Dra("TipoDeCliente")
                HttpContext.Current.Session("Debito_TemProduto") = Dra("Produto")
                HttpContext.Current.Session("Debito_TemCusto") = Dra("CentroDeCusto")
                txtContaDebito.Text = Dra("Conta_Id") & " - " & Dra("Titulo")
            Else
                HttpContext.Current.Session("Credito_Conta") = Dra("Conta_Id")
                HttpContext.Current.Session("Credito_titulo") = Dra("Titulo")
                HttpContext.Current.Session("Credito_TemCliente") = Dra("Cliente")
                HttpContext.Current.Session("Credito_TipoDeCliente") = Dra("TipoDeCliente")
                HttpContext.Current.Session("Credto_TemProduto") = Dra("Produto")
                HttpContext.Current.Session("Credito_TemCusto") = Dra("CentroDeCusto")
                txtContaCredito.Text = Dra("Conta_Id") & " - " & Dra("Titulo")
            End If

        Next
    End Sub

    Private Sub ConsultaCliente()
        Sqla = "SELECT * " & vbCrLf & _
               " FROM Clientes " & vbCrLf & _
               " WHERE Cliente_Id = '" & Codigo & "' and Endereco_Id = " & Endereco & vbCrLf

        For Each Dra As DataRow In Banco.ConsultaDataSet(Sqla, "Clientes").Tables(0).Rows
            If Condicao = "Debito" Then
                HttpContext.Current.Session("Debito_Cliente") = Dra("Cliente_ID")
                HttpContext.Current.Session("Debito_ClienteEnd") = Dra("Endereco_ID")
                HttpContext.Current.Session("Debito_ClienteNome") = UCase(Dra("Nome"))
                HttpContext.Current.Session("Debito_ClienteCidade") = UCase(Dra("Cidade"))
                HttpContext.Current.Session("Debito_ClienteEstado") = UCase(Dra("Estado"))
                txtClienteDebito.Text = Funcoes.FormatarCpfCnpj(Dra("Cliente_ID")) & "- " & UCase(Dra("Nome"))
            Else
                HttpContext.Current.Session("Credito_Cliente") = Dra("Cliente_ID")
                HttpContext.Current.Session("Credito_ClienteEnd") = Dra("Endereco_ID")
                HttpContext.Current.Session("Credito_ClienteNome") = UCase(Dra("Nome"))
                HttpContext.Current.Session("Credito_ClienteCidade") = UCase(Dra("Cidade"))
                HttpContext.Current.Session("Credito_ClienteEstado") = UCase(Dra("Estado"))
                txtClienteCredito.Text = Funcoes.FormatarCpfCnpj(Dra("Cliente_ID")) & "- " & UCase(Dra("Nome"))
            End If

        Next
    End Sub

    Private Sub ConsultaProduto()
        If Codigo <> "0" And Codigo <> "" Then
            Sqla = "SELECT Grupo FROM Produtos WHERE Produto_Id = '" & Codigo & "'"
            For Each Dra As DataRow In Banco.ConsultaDataSet(Sqla, "Clientes").Tables(0).Rows
                DdlGrupo.SelectedValue = Dra("Grupo")
            Next
            CarregarProduto()
            ddlProduto.SelectedValue = Codigo
        End If
    End Sub

    Protected Sub ddlSequencia_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlSequencia.Text <> "" Then
                Consultar()
            Else
                MsgBox(Me.Page, "Selecione uma sequência para a consulta.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCustoDebito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If HttpContext.Current.Session("Debito_TemCusto") <> "S" And ddlCustoDebito.SelectedValue <> "" Then
                ddlCustoDebito.SelectedIndex = 0
                MsgBox(Me.Page, "Esta conta de débito não tem centro de custo.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub ddlCustoCredito_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If HttpContext.Current.Session("Credito_TemCusto") <> "S" And ddlCustoCredito.SelectedValue <> "" Then
                ddlCustoCredito.SelectedIndex = 0
                MsgBox(Me.Page, "Esta conta de crédito não tem centro de custo.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub txtMovimento_TextChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Limpar()
            ddlLote.SelectedIndex = 0
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If

            ddlLote.Focus()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("LancamentosContabeisXCustos", "GRAVAR") Then
                If Validar() Then
                    If ddlSequencia.Text <> "" Then
                        Excluir()
                    End If
                    Gravar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAtualizar_Click(sender As Object, e As EventArgs) Handles lnkAtualizar.Click
        Try
            If Funcoes.VerificaPermissao("LancamentosContabeisXCustos", "ALTERAR") Then
                If Validar() Then
                    Excluir()
                    Gravar()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para alterar registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("LancamentosContabeisXCustos", "EXCLUIR") Then
                Excluir()
                Limpar()
                TotalizaLote()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            Consultar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
            If ddlSequencia.Text <> "" Then
                ddlSequencia.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "LancamentosContabeisXCustos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class