Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class MovimentacoesFinanceiras
    Inherits BasePage

    Dim Sql As String
    Dim Sqla As String

    Dim SqlArray As New ArrayList
    Dim Unidades As New ArrayList

    Dim DS As DataSet

    Dim PeriodoInicial As Date
    Dim PeriodoFinal As Date
    Dim Codigo As String
    Dim Descricao As String
    Dim Cliente As String
    Dim Endereco As String
    Dim campo() As String

    Dim Valor As Decimal
    Dim Data As String

    Dim Registro As Integer
    Dim TemRegistro As String

    Dim Raz_Empresa As String
    Dim Raz_EndEmpresa As Integer
    Dim Raz_Conta As String
    Dim Raz_Cliente As String
    Dim Raz_EndCliente As String
    Dim Raz_UnidadeDeNegocio As String
    Dim Raz_ValorOficial As String
    Dim Raz_ValorMoeda As String

    Dim Raz_Historico As String
    Dim Raz_DebitoCredito As String
    Dim Mensagem As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Financeiro)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("MovimentacoesFinanceiras", "ACESSAR") Then
                CargaUnidadeDeNegocioEmpresaCliente()
                CargaUnidadeConsultaTitulos()
                Empresas()
                TiposDePagamentos()
                Provisoes()
                Carteiras()
                TabContainer1.ActiveTabIndex = 0
                VerificaUnidade()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Financeiro.aspx")
                Exit Sub
            End If
        End If
    End Sub

    Private Sub CargaUnidadeConsultaTitulos()
        ddl.Carregar(DdlUnidadeConsultaTitulos, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(DdlUnidadeConsultaTitulos, DdlEmpresaConsultaTitulos)
    End Sub

    Private Sub CargaEmpresa()
        ddl.Carregar(DdlEmpresaConsultaTitulos, CarregarDDL.Tabela.Empresas, DdlUnidadeConsultaTitulos.SelectedValue, True)
    End Sub

    Private Sub VerificaUnidade()
        Sql = "  SELECT isnull(AcessoUnidade, '') as AcessoUnidade," & vbCrLf & _
              "        isnull(AcessoEmpresa, '') as AcessoEmpresa," & vbCrLf & _
              "        isnull(AcessoEndEmpresa, 0) as AcessoEndEmpresa" & vbCrLf & _
              " from Usuarios" & vbCrLf & _
              " where  Usuario_Id = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Consulta").Tables(0).Rows
            DdlUnidadeDeNegocioEmpresaCliente.SelectedValue = Dr("AcessoUnidade")
            CargaEmpresaCliente()
            DdlEmpresaCliente.SelectedValue = Dr("AcessoEmpresa") & "-" & Dr("AcessoEndEmpresa")
            BancoSolicitante()
        Next
    End Sub

    Private Sub CargaUnidadeDeNegocioEmpresaCliente()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String

        Sql = "SELECT Clientes.Cliente_Id as Codigo, Clientes.Nome, Clientes.Cidade, Clientes.Estado  " & vbCrLf & _
              " FROM Clientes INNER JOIN" & vbCrLf & _
              " ClientesXTipos ON Clientes.Cliente_Id = ClientesXTipos.Cliente_Id" & vbCrLf & _
              " WHERE ClientesXTipos.Tipo_Id = 050 Order by Nome" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "UnidadeDeNegocio").Tables(0).Rows
            Codigo = Dr("Codigo")
            Nome = Dr("Nome")
            Descricao = Nome
            DdlUnidadeDeNegocioEmpresaCliente.Items.Add(New ListItem(Descricao, Codigo))
            DdlUnidadeConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))
        Next

        DdlUnidadeDeNegocioEmpresaCliente.Items.Insert(0, "")
        DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = 0

        DdlUnidadeConsultaTitulos.Items.Insert(0, "")
        DdlUnidadeConsultaTitulos.SelectedIndex = 0

    End Sub

    Protected Sub DdlUnidadeDeNegocioEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CargaEmpresaCliente()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub CargaEmpresaCliente()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        DdlEmpresaCliente.Items.Clear()

        Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
              " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
              " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
              " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "' " & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaCliente.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresaCliente.Items.Insert(0, "")
        DdlEmpresaCliente.SelectedIndex = 0


    End Sub

    Private Sub Empresas()
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Unidade As String = ""

        Dim Cnpj As String

        Sql = " SELECT GruposXEmpresas.Empresa_Id as Unidade" & vbCrLf & _
              " FROM Clientes INNER JOIN" & vbCrLf & _
              " GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id" & vbCrLf & _
              " WHERE GruposXEmpresas.Cliente_Id = '" & HttpContext.Current.Session("ssEmpresa") & "'" & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Unidades.Add(Dr("Unidade"))
        Next

        DdlEmpresaPagadora.Items.Clear()

        Sql = "  SELECT Distinct Clientes.Cliente_Id AS Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado" & vbCrLf & _
              " FROM Clientes INNER JOIN" & vbCrLf & _
              " GruposXEmpresas ON Clientes.Cliente_Id = GruposXEmpresas.Cliente_Id AND Clientes.Endereco_Id = GruposXEmpresas.EndCliente_Id" & vbCrLf & _
              " WHERE " & vbCrLf

        For index = 0 To Unidades.Count - 1
            If index = 0 Then
                Sql &= " GruposXEmpresas.Empresa_Id = '" & Unidades(index) & "'"
            Else
                Sql &= " or GruposXEmpresas.Empresa_Id = '" & Unidades(index) & "'"
            End If
        Next
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows

            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

            DdlEmpresaPagadora.Items.Add(New ListItem(Descricao, Codigo))

        Next

        DdlEmpresaPagadora.Items.Insert(0, "")
        DdlEmpresaPagadora.SelectedIndex = 0
    End Sub

    Private Sub TiposDePagamentos()
        Sql = "SELECT TipoDePagamento_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(TipoDePagamento_Id AS varchar))) + CAST(TipoDePagamento_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM TiposDePagamentos Order By TipoDePagamento_Id"

        ddlTiposDePagamentos.DataValueField = "Codigo"
        ddlTiposDePagamentos.DataTextField = "Descricao"
        ddlTiposDePagamentos.DataSource = Banco.ConsultaDataSet(Sql, "TiposDePagamentos")
        ddlTiposDePagamentos.DataBind()

        ddlTiposDePagamentos.Items.Insert(0, "")
        ddlTiposDePagamentos.SelectedIndex = 0
    End Sub

    Private Sub Provisoes()
        Sql = "SELECT Provisao_Id as Codigo, convert(varchar,REPLICATE('0', 2 - LEN(CAST(Provisao_Id AS varchar))) + CAST(Provisao_Id AS varchar)) + '  -  ' + Descricao as Descricao FROM Provisoes Order By Provisao_Id"

        DdlProvisoes.DataValueField = "Codigo"
        DdlProvisoes.DataTextField = "Descricao"
        DdlProvisoes.DataSource = Banco.ConsultaDataSet(Sql, "Provisoes")
        DdlProvisoes.DataBind()

        DdlProvisoes.Items.Insert(0, "")
        DdlProvisoes.SelectedIndex = 0
    End Sub

    Private Sub Carteiras()
        Sql = "SELECT  Produto_Id AS Codigo, CONVERT(varchar, REPLICATE('0', 9 - LEN(CAST(Produto_Id AS varchar))) + CAST(Produto_Id AS varchar))  + '  -  ' + Descricao AS Descricao  FROM ComprasXProdutos Where Classificacao = 'M' Order By Produto_Id"

        ddlCarteiras.DataValueField = "Codigo"
        ddlCarteiras.DataTextField = "Descricao"
        ddlCarteiras.DataSource = Banco.ConsultaDataSet(Sql, "Carteiras")
        ddlCarteiras.DataBind()

        ddlCarteiras.Items.Insert(0, "")
        ddlCarteiras.SelectedIndex = 0
    End Sub

    Private Sub BancoSolicitante()
        DdlBancoFilial.Items.Clear()
        ddlContaFilial.Items.Clear()

        Cliente = DdlEmpresaCliente.SelectedValue
        campo = Cliente.Split("-")

        Sql = "     SELECT  BancosXContas.Banco_Id,  Bancos.Descricao" & vbCrLf & _
                "   FROM    BancosXContas INNER JOIN" & vbCrLf & _
                "   Bancos  ON BancosXContas.Banco_Id = Bancos.Banco_Id" & vbCrLf & _
                "   WHERE     BancosXContas.Empresa_Id  = '" & campo(0) & "'" & vbCrLf & _
                "   and BancosXContas.EndEmpresa_Id  = " & campo(1) & vbCrLf & _
                "   group by BancosXContas.Banco_Id, Bancos.Descricao"

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Bancos").Tables(0).Rows
            Descricao = Format(Dr("Banco_Id"), "0000") & "- " & Dr("Descricao")
            DdlBancoFilial.Items.Add(New ListItem(Descricao, Dr("Banco_Id")))
        Next

        DdlBancoFilial.Items.Insert(0, "")
        DdlBancoFilial.SelectedIndex = 0
    End Sub

    Private Sub BancoPagadora()
        DdlBancoPagador.Items.Clear()
        DdlContaPagadora.Items.Clear()

        Cliente = DdlEmpresaPagadora.SelectedValue
        campo = Cliente.Split("-")

        Sql = "     SELECT  BancosXContas.Banco_Id,  Bancos.Descricao" & vbCrLf & _
                "   FROM    BancosXContas INNER JOIN" & vbCrLf & _
                "   Bancos  ON BancosXContas.Banco_Id = Bancos.Banco_Id" & vbCrLf & _
                "   WHERE     BancosXContas.Empresa_Id  = '" & campo(0) & "'" & vbCrLf & _
                "   and BancosXContas.EndEmpresa_Id  = " & campo(1) & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Bancos").Tables(0).Rows
            Descricao = Format(Dr("Banco_Id"), "0000") & "- " & Dr("Descricao")
            DdlBancoPagador.Items.Add(New ListItem(Descricao, Dr("Banco_Id")))
        Next

        DdlBancoPagador.Items.Insert(0, "")
        DdlBancoPagador.SelectedIndex = 0
    End Sub

    Function ValidaCampos()
        Mensagem = ""

        If ValidaData(txtMovimento.Text, "Movimento", HttpContext.Current.Session("ssEmpresa"), HttpContext.Current.Session("ssEndEmpresa")) Then
        Else
            Return Mensagem
        End If
        If DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = 0 Then
            Mensagem = "Unidade de negócio é obrigatório."
            Return Mensagem
        End If
        If DdlEmpresaCliente.SelectedIndex = 0 Then
            Mensagem = "Empresa Solicitante é obrigatório."
            Return Mensagem
        End If

        If DdlEmpresaPagadora.Text <> "" Then
            If ddlTiposDePagamentos.Text = "" Then
                Mensagem = "Tipo de Pagamento é obrigatório."
                Return Mensagem
            End If
        End If
        If ddlTiposDePagamentos.Text <> "" Then
            If DdlEmpresaPagadora.Text = "" Then
                Mensagem = "Empresa Pagadora é obrigatório."
                Return Mensagem
            End If
        End If
        If DdlProvisoes.Text = "" Then
            Mensagem = "Previsão é obrigatório."
            Return Mensagem
        End If
        If DdlProvisoes.SelectedValue = 1 Then
            If DdlEmpresaPagadora.SelectedIndex = 0 Then
                Mensagem = "Empresa Pagadora é obrigatório."
                Return Mensagem
            End If
        End If
        If DdlProvisoes.SelectedIndex = 0 Then
            Mensagem = "Previsão é obrigatório."
            Return Mensagem
        End If
        If txtHistorico.Text = "" Then
            Mensagem = "Histórico é obrigatório."
            Return Mensagem
        End If

        If ddlTiposDePagamentos.Text <> "" Then
            If ddlTiposDePagamentos.SelectedValue > 1 Then
                If DdlBancoPagador.Text = "" Then
                    Mensagem = "Banco é obrigatório."
                    Return Mensagem
                End If
            End If
        End If
        If txtVencimento.Text = "" Then
            Mensagem = "Data de Programação é obrigatório."
            Return Mensagem
        End If

        campo = DdlEmpresaCliente.SelectedValue.Split("-")
        If ValidaData(txtVencimento.Text, "Programação", campo(0), campo(1)) Then
        Else
            Return Mensagem
        End If
        If DdlProvisoes.SelectedValue = 1 Then
            If txtValor.Text = "" Then
                Mensagem = "Valor é obrigatório."
                Return Mensagem
            End If
        End If
        If ddlCarteiras.SelectedIndex = 0 Then
            Mensagem = "Carteira é obrigatório."
            Return Mensagem
        End If
        If txtValor.Text = "" Then
            txtValor.Text = 0
        End If
        Valor = CDbl(txtValor.Text)
        If Valor = 0 Then
            Mensagem = "Valor é obrigatório."
            Return Mensagem
        End If

        Return Mensagem
    End Function

    Function LanctosContabeis()
        Sql = "INSERT INTO Razao " & vbCrLf & _
              " (Empresa_Id, " & vbCrLf & _
              " EndEmpresa_Id, " & vbCrLf & _
              " Conta_Id, " & vbCrLf & _
              " Cliente_Id, " & vbCrLf & _
              " EndCliente_Id, " & vbCrLf & _
              " Movimento_Id, " & vbCrLf & _
              " Lote_Id, " & vbCrLf & _
              " Sequencia_Id, " & vbCrLf & _
              " Titulo, " & vbCrLf & _
              " UnidadeDeNegocio, " & vbCrLf & _
              " Indexador, " & vbCrLf & _
              " DataMoeda, " & vbCrLf & _
              " DebitoOficial, " & vbCrLf & _
              " CreditoOficial, " & vbCrLf & _
              " DebitoMoeda, " & vbCrLf & _
              " CreditoMoeda, " & vbCrLf & _
              " Historico, " & vbCrLf & _
              " PrevistoRealizado)" & vbCrLf & _
              " VALUES (" & vbCrLf

        Sql &= "'" & Raz_Empresa & "'"              'Empresa
        Sql &= ", " & Raz_EndEmpresa                'Endereco Empresa 
        Sql &= ", '" & Raz_Conta & "'"              'Conta

        If Len(Raz_Conta) = 7 Then
            Sql &= ", '" & Raz_Cliente & "'"        'Cliente
            Sql &= ", " & Raz_EndCliente            'Endereco do Cliente
        Else
            Sql &= ", ''"                           'Cliente
            Sql &= ", 0"                            'Endereco do Cliente
        End If

        Sql &= ", '" & txtMovimento.Text.ToSqlDate() & "'"     'Data de Movimento
        Sql &= ", 0070"
        Sql &= ", " & Registro                      'Sequencia no Razao = Registro do Titulo
        Sql &= ", " & Registro                      'Numero do Titulo
        Sql &= ", '" & Raz_UnidadeDeNegocio & "'"   'Unidade de Negócio
        Sql &= ", 3"                                'Indexador
        Sql &= ", '" & txtMovimento.Text.ToSqlDate() & "'"     'Data da Moeda

        If Raz_DebitoCredito = "D" Then
            Sql &= ", " & Replace(Raz_ValorOficial, ",", ".")  'Valor Débito
            Sql &= ", 0.0"                              'Valor Crédito
            Sql &= ", " & Replace(Raz_ValorMoeda, ",", ".")  'Valor Débito
            Sql &= ", 0.0"                              'Valor Crédito
        Else
            Sql &= ", 0.0"                              'Valor Debito
            Sql &= ", " & Replace(Raz_ValorOficial, ",", ".")  'Valor Crédito
            Sql &= ", 0.0"                              'Valor Debito
            Sql &= ", " & Replace(Raz_ValorMoeda, ",", ".")  'Valor Crédito
        End If

        Sql &= ", '" & Raz_Historico & "'"          'Histórico
        Sql &= ", 'P')"                             'Previsto/Realizado
        SqlArray.Add(Sql)

        Return True
    End Function

    Private Sub GravaTitulo()
        Dim Cliente As String
        Dim Campo() As String
        Dim Valor As String
        Dim Opcao As String = txtRegistro.Text

        ValidaCampos()

        If Mensagem = "" Then

            'Gera sequencia de titulos'''
            If txtRegistro.Text = "" Then
                Sql = "exec sp_Numerador '" & HttpContext.Current.Session("ssNomeServidor").ToUpper() & "',0,1"
                For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Numerador").Tables(0).Rows
                    Registro = Dr("Sequencia")
                    txtRegistro.Text = Registro
                Next
            Else
                Registro = CInt(txtRegistro.Text)

                Sql = "DELETE FROM MovimentacoesFinanceiras" & vbCrLf & _
                      " WHERE Registro_Id = " & Registro & vbCrLf
                SqlArray.Add(Sql)

                Sql = "DELETE FROM razao" & vbCrLf & _
                      " WHERE Titulo = " & Registro & vbCrLf
                SqlArray.Add(Sql)
            End If

            Data = txtMovimento.Text

            Sql = "INSERT INTO MovimentacoesFinanceiras " & vbCrLf & _
                  " (Registro_Id" & vbCrLf & _
                  " ,Sequencia_Id" & vbCrLf & _
                  " ,Provisao" & vbCrLf & _
                  " ,Carteira" & vbCrLf & _
                  " ,Tributo" & vbCrLf & _
                  " ,Indexador" & vbCrLf & _
                  " ,Moeda" & vbCrLf & _
                  " ,TipoPagto" & vbCrLf & _
                  " ,Situacao" & vbCrLf & _
                  " ,Lote" & vbCrLf & _
                  " ,Movimento" & vbCrLf & _
                  " ,Vencimento" & vbCrLf & _
                  " ,Prorrogacao" & vbCrLf & _
                  " ,DataMoeda" & vbCrLf & _
                  " ,Baixa" & vbCrLf

            Sql &= " ,UnidadeDeNegocio" & vbCrLf & _
                   " ,Empresa" & vbCrLf & _
                   " ,EndEmpresa" & vbCrLf & _
                   " ,Cliente" & vbCrLf & _
                   " ,EndCliente" & vbCrLf & _
                   " ,BancoCliente" & vbCrLf & _
                   " ,AgenciaCliente" & vbCrLf & _
                   " ,DigitoAgenciaCliente" & vbCrLf & _
                   " ,ContaCliente" & vbCrLf & _
                   " ,DigitoContaCliente" & vbCrLf & _
                   " ,ContaContabilCliente" & vbCrLf

            Sql &= " ,EmpresaPagadora" & vbCrLf & _
                   " ,EndEmpresaPagadora" & vbCrLf & _
                   " ,BancoPagador" & vbCrLf & _
                   " ,AgenciaPagadora" & vbCrLf & _
                   " ,DigitoAgenciaPagadora" & vbCrLf & _
                   " ,ContaPagadora" & vbCrLf & _
                   " ,DigitoContaPagadora" & vbCrLf & _
                   " ,ContaContabilPagadora" & vbCrLf

            Sql &= " ,ValorDoDocumento" & vbCrLf & _
                   " ,Descontos" & vbCrLf & _
                   " ,Deducoes" & vbCrLf & _
                   " ,Juros" & vbCrLf & _
                   " ,Acrescimos" & vbCrLf & _
                   " ,ValorLiquido" & vbCrLf

            Sql &= " ,MoedaValorDoDocumento" & vbCrLf & _
                   " ,MoedaDescontos" & vbCrLf & _
                   " ,MoedaDeducoes" & vbCrLf & _
                   " ,MoedaJuros" & vbCrLf & _
                   " ,MoedaAcrescimos" & vbCrLf & _
                   " ,MoedaValorLiquido" & vbCrLf

            Sql &= " ,Historico"

            Sql &= " ,Destinatario" & vbCrLf & _
                   " ,EndDestinatario" & vbCrLf & _
                   " ,solicitacao" & vbCrLf

            Sql &= " ,Cheque" & vbCrLf & _
                   " ,Slips" & vbCrLf

            Sql &= " ,UsuarioInclusao" & vbCrLf & _
                   " ,UsuarioInclusaoData" & vbCrLf & _
                   " ,UsuarioBaixa" & vbCrLf & _
                   " ,UsuarioBaixaData" & vbCrLf & _
                   " ,SituacaoBancaria)" & vbCrLf


            Sql &= " VALUES( "

            Sql &= Registro                   'Registro
            Sql &= ", 0"                      'Sequencia
            Sql &= ", " & DdlProvisoes.SelectedValue
            Sql &= ", '" & ddlCarteiras.SelectedValue & "'"
            Sql &= ", ''"                       'Tributo
            Sql &= ", 3"                        'Indexador
            Sql &= ", 0"                        'Moeda

            If ddlTiposDePagamentos.Text <> "" Then
                Sql &= ", " & ddlTiposDePagamentos.SelectedValue
            Else
                Sql &= ", 0"
            End If

            Sql &= ", 1"                        'Situacao
            Sql &= ", 70"                       'Lote
            Sql &= ", '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf & _
                   ", '" & txtVencimento.Text.ToSqlDate() & "'" & vbCrLf & _
                   ", '" & txtVencimento.Text.ToSqlDate() & "'" & vbCrLf & _
                   ", '" & txtVencimento.Text.ToSqlDate() & "'" & vbCrLf & _
                   ", '" & txtMovimento.Text.ToSqlDate() & "'" & vbCrLf

            Sql &= ", '" & DdlUnidadeDeNegocioEmpresaCliente.SelectedValue & "'"

            Cliente = DdlEmpresaCliente.SelectedValue
            Campo = Cliente.Split("-")
            Sql &= ", '" & Campo(0) & "'"                               'Empresa Cliente
            Sql &= ", " & Campo(1)                                      'Endereco Empresa Cliente
            Sql &= ", '" & Campo(0) & "'"                               'Cliente
            Sql &= ", " & Campo(1)                                      'Endereco Cliente

            If DdlBancoFilial.Text <> "" Then
                Sql &= ", " & DdlBancoFilial.SelectedValue              'Banco Cliente
            Else
                Sql &= ", 0"                                            'Banco Cliente
            End If

            Cliente = ddlContaFilial.SelectedValue
            Campo = Cliente.Split("-")
            Sql &= ", '" & Campo(0) & "'"                               'Agencia Cliente
            Sql &= ", '" & Campo(1) & "'"                               'Digito Agencia Cliente
            Sql &= ", '" & Campo(2) & "'"                               'Conta Cliente
            Sql &= ", '" & Campo(3) & "'"                               'Digito Conta Cliente
            Sql &= ", '" & Campo(4) & "'"                               'Conta Contabil

            If ddlTiposDePagamentos.Text <> "" Then
                Cliente = DdlEmpresaPagadora.SelectedValue
                Campo = Cliente.Split("-")
                Sql &= ", '" & Campo(0) & "'"                           'Empresa Pagadora
                Sql &= ", " & Campo(1)                                  'Endereco Empresa Pagadora

                Sql &= ", '" & DdlBancoPagador.SelectedValue & "'"      'Banco Cliente
                Cliente = DdlContaPagadora.SelectedValue
                Campo = Cliente.Split("-")
                Sql &= ", '" & Campo(0) & "'"                           'Agencia Cliente
                Sql &= ", '" & Campo(1) & "'"                           'Digito Agencia Cliente
                Sql &= ", '" & Campo(2) & "'"                           'Conta Cliente
                Sql &= ", '" & Campo(3) & "'"                           'Digito Conta Cliente
                Sql &= ", '" & Campo(4) & "'"                           'Conta Contabil
            Else
                Sql &= ", ''"                                           'Empresa Pagadora
                Sql &= ", 0"                                            'Endereco Empresa Pagadora
                Sql &= ", 0"                                            'Banco Pagador
                Sql &= ", ''"                                           'Agencia Pagadora
                Sql &= ", ''"                                           'Digito Agencia Pagadora
                Sql &= ", ''"                                           'Conta Corrente pagadora
                Sql &= ", ''"                                           'Digito Conta Pagadora
                Sql &= ", ''"                                           'Conta Contabil Pagadora
            End If

            Valor = Replace(txtValor.Text, ".", "")                     'Valor Documento
            Sql &= ", " & Replace(Valor, ",", ".")
            Sql &= ", 0.0"                                              'Descontos
            Sql &= ", 0.0"                                              'Deducoes
            Sql &= ", 0.0"                                              'Juros
            Sql &= ", 0.0"                                              'Acrescimos
            Sql &= ", " & Replace(Valor, ",", ".")                      'Valor Liquido

            Valor = Replace(DolarizaBaixa(txtValor.Text), ".", "")
            Sql &= ", " & Replace(Valor, ",", ".")                      'Moeda
            Sql &= ", 0.0"                                              'Descontos
            Sql &= ", 0.0"                                              'Deducoes
            Sql &= ", 0.0"                                              'Juros
            Sql &= ", 0.0"                                              'Acrescimos
            Sql &= ", " & Replace(Valor, ",", ".")                      'Moeda Liquido

            Sql &= ", '" & txtHistorico.Text & "'"                      'Historico

            Cliente = DdlEmpresaCliente.SelectedValue
            Campo = Cliente.Split("-")
            Sql &= ", '" & Campo(0) & "'"                               'Destinatario
            Sql &= ", " & Campo(1)                                      'Endereco destinatario
            Sql &= ", 0"                                                'Solicitacao

            If rbChequeSim.Checked Then                                 'Emite Cheque
                Sql &= ", 'N'"
            ElseIf rbChequeNao.Checked Then
                Sql &= ", 'S'"
            Else
                Sql &= ", 'N'"

            End If

            Sql &= ", 'N'"                                              'Slips

            Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"  'Usuario que Incluiu
            Sql &= ", '" & Today.ToSqlDate() & "'"     'Data da Inclusao

            Sql &= ", '" & HttpContext.Current.Session("ssNomeUsuario") & "'"  'Usuario que Baixou
            Sql &= ", '" & Today.ToSqlDate() & "'"    'Data da Baixa
            Sql &= ", 0)"                                              'Situacao bancaria
            SqlArray.Add(Sql)

            '-Gravação no Razão Contabil------------------

            If DdlProvisoes.SelectedValue = 1 Then
                ' Grava Razao Debito
                Registro = CInt(txtRegistro.Text)
                Cliente = DdlEmpresaCliente.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Empresa = Campo(0)                  'EmpresaCliente
                Raz_EndEmpresa = Campo(1)               'Endereco Empresa Cliente
                Cliente = ddlContaFilial.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Conta = Campo(4)                    'Conta Contabil
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente
                Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue

                Raz_ValorOficial = Replace(txtValor.Text, ".", "")      'ValorDoDocumento
                Raz_ValorMoeda = Replace(Valor, ".", "")                'ValorDoDocumento

                Raz_Historico = txtHistorico.Text
                Raz_DebitoCredito = "D"
                LanctosContabeis()

                'Gravar Credito
                Registro = CInt(txtRegistro.Text)
                Cliente = DdlEmpresaPagadora.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Empresa = Campo(0)                  'Empresa Pagadora
                Raz_EndEmpresa = Campo(1)               'Endereco Empresa Pagadora
                Cliente = DdlContaPagadora.SelectedValue
                Campo = Cliente.Split("-")
                Raz_Conta = Campo(4)                    'Conta Contabil
                Raz_Cliente = ""                        'Cliente
                Raz_EndCliente = 0                      'Endereco do Cliente
                Raz_UnidadeDeNegocio = "01"             'Pegar Unidadde de Negocio da EmpresaPagadora

                Raz_ValorOficial = Replace(txtValor.Text, ".", "")      'ValorDoDocumento
                Raz_ValorMoeda = Replace(Valor, ".", "")                'ValorDoDocumento

                Raz_Historico = txtHistorico.Text
                Raz_DebitoCredito = "C"
                LanctosContabeis()

                'Transferencias Financeiras
                Cliente = DdlEmpresaCliente.SelectedValue       'Empresa do Cliente
                Campo = Cliente.Split("-")

                Sql = "Select EmpresaDebito,EnderecoDebito,EmpresaCredito,EnderecoCredito,EmpresaContabil,EnderecoContabil,ContaContabil," & vbCrLf & _
                      "ClienteContabil,EndClienteContabil,DebitoCredito from TransferenciasFinanceiras where " & vbCrLf & _
                      " EmpresaDebito='" & Campo(0) & "' and EnderecoDebito = " & Campo(1) & " and " & vbCrLf

                Cliente = DdlEmpresaPagadora.SelectedValue   'Empresa Pagadora
                Campo = Cliente.Split("-")
                Sql &= "EmpresaCredito = '" & Campo(0) & "' and EnderecoCredito = " & Campo(1) & ""

                For Each DrT As DataRow In Banco.ConsultaDataSet(Sql, "Transferencias").Tables(0).Rows
                    Raz_Empresa = DrT("EmpresaContabil")                'EmpresaCliente
                    Raz_EndEmpresa = DrT("EnderecoContabil")            'Endereco Empresa Cliente
                    Raz_Conta = DrT("ContaContabil")                    'Grupo de Contas
                    Raz_Cliente = DrT("ClienteContabil")                'Cliente
                    Raz_EndCliente = DrT("EndClienteContabil")          'Endereco do Cliente
                    Raz_UnidadeDeNegocio = DdlUnidadeDeNegocioEmpresaCliente.SelectedValue

                    Raz_ValorOficial = Replace(txtValor.Text, ".", "")      'ValorDoDocumento
                    Raz_ValorMoeda = Replace(Valor, ".", "")                'ValorDoDocumento

                    Raz_Historico = txtHistorico.Text                   'Historico
                    Raz_DebitoCredito = DrT("DebitoCredito")            'Debito/Credito

                    LanctosContabeis()

                Next
            End If

            If Banco.GravaBanco(SqlArray) = False Then
                MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
            Else
                Limpar()
                Mensagem = "Registro <" & Registro & "> Gravado/Atualizado com sucesso."
                MsgBox(Me.Page, Mensagem)
            End If
        Else
            MsgBox(Me.Page, Mensagem)
        End If

    End Sub

    Private Sub Limpar()

        lnkExcluir.Parent.Visible = False

        DdlBancoFilial.Items.Clear()
        ddlContaFilial.Items.Clear()

        DdlEmpresaPagadora.SelectedIndex = 0
        DdlBancoPagador.Items.Clear()
        DdlContaPagadora.Items.Clear()

        ddlCarteiras.SelectedIndex = 0
        ddlTiposDePagamentos.SelectedIndex = 0
        DdlProvisoes.SelectedIndex = 2

        txtRegistro.Text = ""
        txtHistorico.Text = ""
        txtMovimento.Text = ""
        txtVencimento.Text = ""
        txtValor.Text = ""

        txtRegistro.Enabled = True

        txtMovimento.Text = Format(Today, "dd/MM/yyyy")
        txtVencimento.Text = Format(Today, "dd/MM/yyyy")
        txtPeriodoInicialConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")
        txtPeriodoFinalConsultaTitulos.Text = Format(Today, "dd/MM/yyyy")

        rbChequeNao.Checked = True

        VerificaUnidade()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            DdlUnidadeDeNegocioEmpresaCliente.Enabled = False
            DdlEmpresaCliente.Enabled = False

            DdlUnidadeConsultaTitulos.Enabled = False
            DdlEmpresaConsultaTitulos.Enabled = False
        End If
    End Sub

    Protected Sub GridConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Registro = GridConsultaTitulos.SelectedRow.Cells(1).Text()
            Limpar()
            txtRegistro.Text = Registro
            ConsultaMovimentacoesFinanceiras()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub ConsultaCLientes(ByVal Cli As String, ByVal EndCli As Integer)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        Sql = "SELECT Cliente_Id as Codigo, Endereco_Id, Nome, Cidade, Estado" & vbCrLf & _
              " FROM Clientes Where  Cliente_Id  = '" & Cli & "' And Endereco_Id = " & EndCli & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

            Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
            Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

            Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
            Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
            Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id"))
        Next
    End Sub

    Protected Sub cmdAlterar_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            GravaTitulo()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub ConsultaMovimentacoesFinanceiras()
        Dim Pedido As Integer = 0
        Dim conta As String = ""
        TemRegistro = ""

        If txtRegistro.Text <> "" Then

            Registro = txtRegistro.Text
            'Limpar()
            txtRegistro.Text = Registro

            txtRegistro.Enabled = False

            Sql = "SELECT  " & vbCrLf & _
                  " Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, " & vbCrLf & _
                  " Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, " & vbCrLf & _
                  " EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf & _
                  " ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, " & vbCrLf & _
                  " DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, ContaContabilPagadora, isnull(Cheque,'N') as Cheque, " & vbCrLf & _
                  " Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, " & vbCrLf & _
                  " Procuracao, ValorDoDocumento, Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, " & vbCrLf & _
                  " MoedaValorDoDocumento, MoedaDescontos, MoedaDeducoes, MoedaJuros, MoedaAcrescimos, " & vbCrLf & _
                  " MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, " & vbCrLf & _
                  " NomeDoDestinatario, Destinacao, solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, " & vbCrLf & _
                  " UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, UsuarioLiberacaoData, " & vbCrLf & _
                  " UsuarioBaixa, UsuarioBaixaData, Grupado, RegistroMestre, Observacoes, SituacaoBancaria, NumeroDoCheque, " & vbCrLf & _
                  " Adiantamento, VencimentoAdto, TaxaAdto, UsuarioLiberacaoBloqueio, UsuarioLiberacaoBloqueioDate, " & vbCrLf & _
                  " UsuarioLiberacaoPedido, UsuarioLiberacaoPedidoDate, UsuarioLiberacaoCheque, UsuarioLiberacaoChequeDate, " & vbCrLf & _
                  " CarteiraAdto, CarteiraDoTitulo, ContratoDeFinanciamento, DataEnvio, Ocorrencia, motivo, DigitoNossoNumero, " & vbCrLf & _
                  " NossoNumero, Timbrado, CodigoDeBarra, CodigoDeBarraDigitado, ModalidadeDePagamento, " & vbCrLf & _
                  " CreditoNaConta, NumeroDaRemessa, SituacaoRemessaBancaria, TipoContaCliente " & vbCrLf

            Sql &= " FROM MovimentacoesFinanceiras " & vbCrLf
            Sql &= " WHERE Registro_Id = " & Registro & "and Situacao = 1"

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "MovimentacoesFinanceiras").Tables(0).Rows
                DdlUnidadeDeNegocioEmpresaCliente.SelectedIndex = DdlUnidadeDeNegocioEmpresaCliente.Items.IndexOf(DdlUnidadeDeNegocioEmpresaCliente.Items.FindByValue(Dr("UnidadeDeNegocio")))
                CargaEmpresaCliente()
                DdlEmpresaCliente.SelectedIndex = DdlEmpresaCliente.Items.IndexOf(DdlEmpresaCliente.Items.FindByValue(Dr("Empresa") & "-" & CStr(Dr("EndEmpresa"))))
                'ConsultaCLientes(Dr("Cliente"), Dr("EndCliente"))
                BancoSolicitante()
                DdlBancoFilial.SelectedIndex = DdlBancoFilial.Items.IndexOf(DdlBancoFilial.Items.FindByValue(Dr("BancoCliente")))
                BancoCliente()
                conta = Dr("AgenciaCliente") & "-" & Dr("DigitoAgenciaCliente") & "-" & Dr("ContaCliente") & "-" & Dr("DigitoContaCliente") & "-" & Dr("ContaContabilCliente")
                ddlContaFilial.SelectedIndex = ddlContaFilial.Items.IndexOf(ddlContaFilial.Items.FindByValue(conta))

                ddlTiposDePagamentos.SelectedIndex = ddlTiposDePagamentos.Items.IndexOf(ddlTiposDePagamentos.Items.FindByValue(Dr("TipoPagto")))

                If Dr("EmpresaPagadora") <> "" Then
                    DdlEmpresaPagadora.SelectedIndex = DdlEmpresaPagadora.Items.IndexOf(DdlEmpresaPagadora.Items.FindByValue(Dr("EmpresaPagadora") & "-" & CStr(Dr("EndEmpresaPagadora"))))
                    BancoPagadora()
                    DdlBancoPagador.SelectedIndex = DdlBancoPagador.Items.IndexOf(DdlBancoPagador.Items.FindByValue(Dr("BancoPagador")))
                    BancoPagador()
                    conta = Dr("AgenciaPagadora") & "-" & Dr("DigitoAgenciaPagadora") & "-" & Dr("ContaPagadora") & "-" & Dr("DigitoContaPagadora") & "-" & Dr("ContaContabilPagadora")
                    DdlContaPagadora.SelectedIndex = DdlContaPagadora.Items.IndexOf(DdlContaPagadora.Items.FindByValue(conta))

                End If

                ddlCarteiras.SelectedIndex = ddlCarteiras.Items.IndexOf(ddlCarteiras.Items.FindByValue(Dr("Carteira")))
                DdlProvisoes.SelectedIndex = DdlProvisoes.Items.IndexOf(DdlProvisoes.Items.FindByValue(Dr("Provisao")))

                txtRegistro.Text = Dr("Registro_Id")
                txtHistorico.Text = Dr("Historico")
                txtMovimento.Text = CDate(Dr("Movimento")).ToStrDate()
                txtVencimento.Text = CDate(Dr("Prorrogacao")).ToStrDate()
                txtValor.Text = Dr("ValorLiquido")

                rbChequeNao.Checked = IIf(Dr("Cheque") = "S", True, False)
                rbChequeSim.Checked = IIf(Dr("Cheque") = "N", True, False)

                If Not IsDBNull(Dr("Observacoes")) Then
                    txtObservacoes.Text = Dr("Observacoes")
                Else
                    txtObservacoes.Text = ""
                End If

                If Dr("Situacao") = 1 Then lnkExcluir.Parent.Visible = True

                TabContainer1.ActiveTabIndex = 0

            Next
        Else
            Mensagem = "Informe o número do registro para consulta."
        End If
    End Sub

    Protected Sub DdlUnidadeConsultaTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim Codigo As String
        Dim Descricao As String
        Dim Nome As String
        Dim Cidade As String
        Dim Cnpj As String

        Try
            DdlEmpresaConsultaTitulos.Items.Clear()

            Sql = "  SELECT Clientes.Cliente_Id as Codigo, Clientes.Endereco_Id, Clientes.Reduzido, Clientes.Nome, Clientes.Cidade, Clientes.Estado " & vbCrLf & _
                  " FROM   GruposXEmpresas INNER JOIN" & vbCrLf & _
                  " Clientes ON GruposXEmpresas.Cliente_Id = Clientes.Cliente_Id AND GruposXEmpresas.EndCliente_Id = Clientes.Endereco_Id" & vbCrLf & _
                  " Where GruposXEmpresas.Empresa_Id = '" & DdlUnidadeConsultaTitulos.SelectedValue & "' " & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                Codigo = Dr("Codigo") & "-" & CStr(Dr("Endereco_Id"))

                Cnpj = Funcoes.FormatarCpfCnpj(Dr("Codigo"))
                Cnpj = Funcoes.AlinharEsquerda(Cnpj, 18, ".")

                Nome = Funcoes.AlinharEsquerda(Dr("Nome"), 28, ".")
                Cidade = Funcoes.AlinharEsquerda(Dr("Cidade"), 20, ".")
                Descricao = Nome & " - " & Cidade & " " & Dr("Estado") & " " & Cnpj & "-" & CStr(Dr("Endereco_Id")) & "-" & Dr("Reduzido")

                DdlEmpresaConsultaTitulos.Items.Add(New ListItem(Descricao, Codigo))

            Next

            DdlEmpresaConsultaTitulos.Items.Insert(0, "")
            DdlEmpresaConsultaTitulos.SelectedIndex = 0
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    'Protected Sub Button1_Click(ByVal sender As Object, ByVal e As System.EventArgs)
    '    Try
    '        ScriptManager.RegisterClientScriptBlock(Me, Me.TabContainer1.GetType(), Guid.NewGuid().ToString(), "alert('Teste OK aba 1');", True)
    '    Catch ex As Exception
    '        MsgBox(Me.Page, ex.Message)
    '    End Try
    'End Sub

    Protected Sub DdlBancoFilial_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BancoCliente()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub DdlBancoPagador_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BancoPagador()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub BancoCliente()
        Dim Cliente As String
        Dim Campo() As String
        Dim Conta As String

        Cliente = DdlEmpresaCliente.SelectedValue
        Campo = Cliente.Split("-")
        ddlContaFilial.Items.Clear()

        Sql = "SELECT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes " & vbCrLf & _
              " FROM BancosXContas" & vbCrLf & _
              " Where Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
              " and EndEmpresa_Id  = " & Campo(1) & vbCrLf & _
              " and Banco_Id  = " & DdlBancoFilial.SelectedValue & vbCrLf

        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
            Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("ContaContabil")
            Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "  C/C " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes")
            ddlContaFilial.Items.Add(New ListItem(Descricao, Conta))
        Next

        ddlContaFilial.Items.Insert(0, "")
        ddlContaFilial.SelectedIndex = 0
    End Sub

    Private Sub BancoPagador()
        Dim Cliente As String
        Dim Campo() As String
        Dim Conta As String

        Cliente = DdlEmpresaPagadora.SelectedValue
        Campo = Cliente.Split("-")
        DdlContaPagadora.Items.Clear()

        If Campo(0) <> "" And DdlBancoPagador.Text <> "" Then
            Sql = "SELECT Agencia_Id, DigitoAgencia_Id, Conta_Id, DigitoConta_Id, ContaContabil, Observacoes " & vbCrLf & _
                  " FROM BancosXContas" & vbCrLf & _
                  " Where Empresa_Id = '" & Campo(0) & "'" & vbCrLf & _
                  " and EndEmpresa_Id  = " & Campo(1) & vbCrLf & _
                  " and Banco_Id  = " & DdlBancoPagador.SelectedValue & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "BancosXContas").Tables(0).Rows
                Conta = Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "-" & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & "-" & Dr("ContaContabil")
                Descricao = "AG " & Dr("Agencia_Id") & "-" & Dr("DigitoAgencia_Id") & "  C/C " & Dr("Conta_Id") & "-" & Dr("DigitoConta_Id") & " - " & Dr("Observacoes")
                DdlContaPagadora.Items.Add(New ListItem(Descricao, Conta))
            Next

            DdlContaPagadora.Items.Insert(0, "")
            DdlContaPagadora.SelectedIndex = 0
        End If
    End Sub

    Protected Sub cmdAtualizarObservacoes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If txtRegistro.Text <> "" Then

                Sql = " UPDATE MovimentacoesFinanceiras"
                Sql &= " SET Observacoes = '" & txtObservacoes.Text & "'"

                Sql &= ", UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" 'Usuario que Alterou
                Sql &= ", UsuarioAlteracaoData = '" & Today.ToSqlDate() & "'"   'Data da Alteracao

                Sql &= " WHERE Registro_ID = " & CInt(txtRegistro.Text)
                SqlArray.Add(Sql)

                If Banco.GravaBanco(SqlArray) = False Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                Else
                    Mensagem = "Registro <" & Registro & "> Atualizado com sucesso."
                    MsgBox(Me.Page, Mensagem)
                End If
            Else
                MsgBox(Me.Page, "Informe o número do registro para atualização.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Function ValidaData(ByVal Data As String, ByVal Tipo As String, ByVal Empresa As String, ByVal EndEmpresa As String) As String

        If IsDate(Data) Then
        Else
            Mensagem = " Data de " & Tipo & " inválida."
            Return False
        End If
        If CDate(Data).DayOfWeek = 6 Then
            Mensagem = "Sábado - Data Inválida para " & Tipo & "..."
            Return False
        End If
        If CDate(Data).DayOfWeek = 0 Then
            Mensagem = "Domingo - Data Inválida para " & Tipo & "..."
            Return False
        End If

        Sql = "  SELECT * From DatasNaoProgramaveis" & vbCrLf & _
              " Where Empresa_Id = '99999999999999' And EndEmpresa_ID = 0 And Data_ID = '" & Data.ToSqlDate() & "'" & vbCrLf
        For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
            Mensagem = "Data de " & Tipo & " não programável, Feriado Nacional > " & Dr("Descricao")
            Return False
        Next
        If Empresa <> "" Then
            Sql = "  SELECT * From DatasNaoProgramaveis" & vbCrLf & _
                  " Where Empresa_Id = '" & Empresa & "' And EndEmpresa_ID = " & EndEmpresa & " And Data_ID = '" & Data.ToSqlDate() & "'" & vbCrLf
            For Each Dr As DataRow In Banco.ConsultaDataSet(Sql, "Clientes").Tables(0).Rows
                Mensagem = "Data de " & Tipo & " não programável, Feriado Municipal > " & Dr("Descricao")
                Return False
            Next
        End If

        If Tipo = "Movimento" Then
            If Funcoes.VerificaAcesso(Empresa, EndEmpresa, Data, "Financeiro") Then
            Else
                Mensagem = "Movimento já fechado para esta data."
                Return False
            End If
        End If

        Return True
    End Function

    Protected Sub DdlEmpresaCliente_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BancoSolicitante()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub DdlEmpresaPagadora_SelectedIndexChanged1(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            BancoPagadora()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Function DolarizaBaixa(ByVal Valor As String) As String
        Dim Calculo As Decimal

        Try
            Sqla = "     SELECT   * " & vbCrLf & _
                           "   FROM Cotacoes" & vbCrLf & _
                           "   WHERE Data_Id = '" & Data.ToSqlDate() & "' AND Indexador_Id = 3" & vbCrLf

            For Each Dr As DataRow In Banco.ConsultaDataSet(Sqla, "Encargos").Tables(0).Rows
                Calculo = CDec(Valor) / Dr("Indice")
                Calculo = CDec(FormatNumber(Calculo, 2))
            Next
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try

        Return Calculo.ToString
    End Function

    Protected Sub RbGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbAtivo_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbBaixado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbDiaGeral_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbFilialDiario_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Protected Sub RbCarteiraDia_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)

    End Sub

    Private Sub DiaGeral()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        'Dim dra As DataRow
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        'Dim Inconsist As String
        'Dim Sitban As Integer
        Dim Valdia As Decimal
        'Dim ValdiaDolar As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        'Dim registro As String

        lista = "Geral"

        sql = "  SELECT  MovimentacoesFinanceiras.Registro_Id AS Registro, " & vbCrLf & _
              " convert(varchar(10),MovimentacoesFinanceiras.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              " MovimentacoesFinanceiras.Prorrogacao,103) as Posterga, Clientes.Nome AS Cliente, Historico, " & vbCrLf & _
              " isnull(MovimentacoesFinanceiras.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              " MovimentacoesFinanceiras.ValorLiquido AS Valor, " & vbCrLf & _
              " UsuarioLiberacao as Liberado, " & vbCrLf & _
              " MovimentacoesFinanceiras.Carteira As Carteira, " & vbCrLf & _
              " ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              " MovimentacoesFinanceiras.Situacao as Situacao, " & vbCrLf & _
              " MovimentacoesFinanceiras.Empresa as Empresa, " & vbCrLf & _
              " Empresa.Reduzido as Reduzido, " & vbCrLf & _
              " MovimentacoesFinanceiras.UsuarioBaixa as UsuarioBaixa" & vbCrLf & _
              " FROM MovimentacoesFinanceiras INNER JOIN" & vbCrLf & _
              " Clientes ON MovimentacoesFinanceiras.Cliente = Clientes.Cliente_Id AND MovimentacoesFinanceiras.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
              " INNER JOIN" & vbCrLf & _
              " ComprasXProdutos ON MovimentacoesFinanceiras.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              " INNER JOIN" & vbCrLf & _
              " Clientes as Empresa ON MovimentacoesFinanceiras.Empresa = Empresa.Cliente_Id AND MovimentacoesFinanceiras.EndEmpresa = Empresa.Endereco_Id" & vbCrLf
        ''sql &= " WHERE MovimentacoesFinanceiras.Provisao <> 1 and  MovimentacoesFinanceiras.Situacao = 1"
        sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1  "
        ''If RbAtivo.Checked Then
        '' sql &= " and Provisao <> 1 "
        ''End If
        ''If RbBaixado.Checked Then
        '' sql &= " and Provisao = 1 "
        ''End If
        If RbGeral.Checked = True Then
            '' Nao ira fazer nada pois vai listar todos 
            lista = "Geral"
        End If

        If RbAtivo.Checked = True Then
            ''sql &= " and MovimentacoesFinanceiras.usuariobaixa = '' "
            sql &= " and Provisao <> 1 "
            lista = "Ativos"
        End If

        If RbBaixado.Checked = True Then
            ''sql &= " and MovimentacoesFinanceiras.usuariobaixa <> '' "
            sql &= " and Provisao = 1 "
            lista = "Baixados"
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= " and MovimentacoesFinanceiras.UnidadeDeNegocio = '" & Cliente & "' "
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= " and MovimentacoesFinanceiras.Empresa = '" & Campo(0) & "'"   'Empresa
            sql &= " and MovimentacoesFinanceiras.EndEmpresa = " & Campo(1)       'Endereco da Empresa
        End If

        ' ''Cliente = DdlClienteConsultaTitulos.SelectedValue
        'Campo = Cliente.Split("-")
        'If Campo(0) <> "" Then
        '    sql &= " and MovimentacoesFinanceiras.Cliente = '" & Campo(0) & "'"   'Cliente
        '    sql &= " and MovimentacoesFinanceiras.EndCliente = " & Campo(1)       'Cliente da Empresa
        'End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= " and Vencimento between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
        End If

        '' If ChkAutozizado.Checked = True Then
        '' sql &= " and MovimentacoesFinanceiras.UsuarioLiberacao <> ''"    'Autorizados
        ''End If

        sql &= " ORDER BY MovimentacoesFinanceiras.Vencimento, MovimentacoesFinanceiras.Empresa"

        ''DS = Banco.ConsultaDataSet(Sql, "Contas")

        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf & _
                 "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf & _
                 "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf & _
                 "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        'Cabeçalho Padrao
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        Else
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        Else
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        End If

        ' ''Cliente = DdlClienteConsultaTitulos.SelectedValue
        'Campo = Cliente.Split("-")
        'If Campo(0) <> "" Then
        '    linha &= "<TR>"
        '    linha &= "<TD >" & " Cliente : " & DdlClienteConsultaTitulos.SelectedValue & " </TD>"
        '    linha &= "</TR>"
        'Else
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        ''End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & " Tipo de relatorio Diario Geral - Totalizacao por Data - Registros impressos : " & lista & " </TD>" & vbCrLf & _
                 "</TR>" & vbCrLf & _
                 "<table width= '370' cellpadding='0' cellspacing='0' Border=0>" & vbCrLf & _
                 "<TR>" & vbCrLf & _
                 "<TD><B>Registro</B></TD>" & vbCrLf & _
                 "<TD><B>Cliente/Fornecedor</B></TD>" & vbCrLf & _
                 "<TD><B>Historico</B></TD>" & vbCrLf & _
                 "<TD><B>Receber R$</B></TD>" & vbCrLf & _
                 "<TD><B>Receber US$</B></TD>" & vbCrLf & _
                 "<TD><B>Empresa</B></TD>" & vbCrLf & _
                 "<TD><B>Vencimento Original</B></TD>" & vbCrLf & _
                 "<TD><B>Prorrogacao</B></TD>" & vbCrLf & _
                 "<TD><B>Carteira</B></TD>" & vbCrLf & _
                 "<TD><B>Situacao</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        ''
        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If Datvenctr <> "" And Datvenctr <> dr("Vencimento") Then
                    linha &= "<TR>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & " Total diario: " & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "</TR>" & vbCrLf
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")

                Datvenctr = dr("Vencimento")
                linha &= "<TR>" & vbCrLf & _
                         "<TD>" & dr("Registro") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Cliente") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Historico") & "</TD>" & vbCrLf & _
                         "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>" & vbCrLf & _
                         "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Reduzido") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Vencimento") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Posterga") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("DescricaoCarteira") & "</TD>" & vbCrLf
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If
        linha &= "<TR>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & " Total diario: " & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf & _
                 "<TR>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & " Total neste processamento: " & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtot.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        If contreg = 0 Then
            MsgBox(Me.Page, "Não existem registros corretos para o período.")
            ScriptManager.GetCurrent(Me.Page).SetFocus(Me.txtPeriodoFinalConsultaTitulos)
        Else
            MsgBox(Me.Page, "Movimento com registros processados.")
            ScriptManager.GetCurrent(Me.Page).SetFocus(Me.txtPeriodoFinalConsultaTitulos)
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Finally
            strm.Close()
        End Try
        '' rotina utilizada nas procuracoes fim - html 
        '' rotina de geracao do relatorio . 
    End Sub

    Private Sub FilialGeral()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        'Dim dra As DataRow
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        'Dim Inconsist As String
        'Dim Sitban As Integer
        Dim Valdia As Decimal
        'Dim ValdiaDolar As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        Dim Empresa As String
        'Dim registro As String

        lista = "Geral"

        sql = "  SELECT  MovimentacoesFinanceiras.Registro_Id AS Registro, " & vbCrLf & _
              " convert(varchar(10),MovimentacoesFinanceiras.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              " MovimentacoesFinanceiras.Prorrogacao,103) as Posterga, Clientes.Nome AS Cliente, Historico, " & vbCrLf & _
              " isnull(MovimentacoesFinanceiras.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              " MovimentacoesFinanceiras.ValorLiquido AS Valor, " & vbCrLf & _
              " UsuarioLiberacao as Liberado, " & vbCrLf & _
              " MovimentacoesFinanceiras.Carteira As Carteira, " & vbCrLf & _
              " ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              " MovimentacoesFinanceiras.Situacao as Situacao, " & vbCrLf & _
              " MovimentacoesFinanceiras.Empresa as Empresa, " & vbCrLf & _
              " Empresa.Reduzido as Reduzido, " & vbCrLf & _
              " MovimentacoesFinanceiras.UsuarioBaixa as UsuarioBaixa" & vbCrLf & _
              " FROM MovimentacoesFinanceiras INNER JOIN" & vbCrLf & _
              " Clientes ON MovimentacoesFinanceiras.Cliente = Clientes.Cliente_Id AND MovimentacoesFinanceiras.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
              " INNER JOIN" & vbCrLf & _
              " ComprasXProdutos ON MovimentacoesFinanceiras.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              " INNER JOIN" & vbCrLf & _
              " Clientes as Empresa ON MovimentacoesFinanceiras.Empresa = Empresa.Cliente_Id AND MovimentacoesFinanceiras.EndEmpresa = Empresa.Endereco_Id" & vbCrLf & _
              " WHERE MovimentacoesFinanceiras.Provisao <> 1 and  MovimentacoesFinanceiras.Situacao = 1" & vbCrLf

        If RbGeral.Checked = True Then
            '' Nao ira fazer nada pois vai listar todos 
            lista = "Geral"
        End If
        If RbAtivo.Checked = True Then
            sql &= " and MovimentacoesFinanceiras.usuariobaixa = '' "
            lista = "Ativos"
        End If
        If RbBaixado.Checked = True Then
            sql &= " and MovimentacoesFinanceiras.usuariobaixa <> '' "
            lista = "Baixados"
        End If
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= " and MovimentacoesFinanceiras.UnidadeDeNegocio = '" & Cliente & "' "
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= " and MovimentacoesFinanceiras.Empresa = '" & Campo(0) & "'"   'Empresa
            sql &= " and MovimentacoesFinanceiras.EndEmpresa = " & Campo(1)       'Endereco da Empresa
        End If

        'Cliente = DdlClienteConsultaTitulos.SelectedValue
        'Campo = Cliente.Split("-")
        'If Campo(0) <> "" Then
        '    sql &= " and MovimentacoesFinanceiras.Cliente = '" & Campo(0) & "'"   'Cliente
        '    sql &= " and MovimentacoesFinanceiras.EndCliente = " & Campo(1)       'Cliente da Empresa
        'End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= " and Vencimento between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
        End If

        ''If ChkAutozizado.Checked = True Then
        ''sql &= " and MovimentacoesFinanceiras.UsuarioLiberacao <> ''"    'Autorizados
        ''End If

        sql &= " ORDER BY MovimentacoesFinanceiras.Empresa, MovimentacoesFinanceiras.Vencimento "

        ''DS = Banco.ConsultaDataSet(Sql, "Contas")

        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf & _
                 "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf & _
                 "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf & _
                 "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        'Cabeçalho Padrao
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        Else
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        Else
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        End If

        'Cliente = DdlClienteConsultaTitulos.SelectedValue
        'Campo = Cliente.Split("-")
        'If Campo(0) <> "" Then
        '    linha &= "<TR>"
        '    linha &= "<TD >" & " Cliente : " & DdlClienteConsultaTitulos.SelectedValue & " </TD>"
        '    linha &= "</TR>"
        'Else
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        'End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & " Tipo de relatorio Filial(Empresa)Geral - Totalizacao por Filial - Registros impressos : " & lista & " </TD>" & vbCrLf & _
                 "</TR>" & vbCrLf & _
                 "<table width= '370' cellpadding='0' cellspacing='0' Border=0>" & vbCrLf & _
                 "<TR>" & vbCrLf & _
                 "<TD><B>Registro</B></TD>" & vbCrLf & _
                 "<TD><B>Cliente/Fornecedor</B></TD>" & vbCrLf & _
                 "<TD><B>Historico</B></TD>" & vbCrLf & _
                 "<TD><B>Receber R$</B></TD>" & vbCrLf & _
                 "<TD><B>Receber US$</B></TD>" & vbCrLf & _
                 "<TD><B>Empresa</B></TD>" & vbCrLf & _
                 "<TD><B>Vencimento Original</B></TD>" & vbCrLf & _
                 "<TD><B>Prorrogacao</B></TD>" & vbCrLf & _
                 "<TD><B>Carteira</B></TD>" & vbCrLf & _
                 "<TD><B>Situacao</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf

        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        Empresa = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If Empresa <> "" And Empresa <> dr("Empresa") Then
                    linha &= "<TR>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & " Total Filial: " & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "</TR>" & vbCrLf
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")
                Empresa = dr("Empresa")
                Datvenctr = dr("Vencimento")
                linha &= "<TR>" & vbCrLf & _
                         "<TD>" & dr("Registro") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Cliente") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Historico") & "</TD>" & vbCrLf & _
                         "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>" & vbCrLf & _
                         "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Reduzido") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Vencimento") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Posterga") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("DescricaoCarteira") & "</TD>" & vbCrLf
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If
        linha &= "<TR>>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & " Total Filial: " & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf & _
                 "<TR>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & " Total neste processamento: " & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtot.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        If contreg = 0 Then
            MsgBox(Me.Page, "Não existem registros corretos para o período")
            ScriptManager.GetCurrent(Me.Page).SetFocus(Me.txtPeriodoFinalConsultaTitulos)
        Else
            MsgBox(Me.Page, "Movimento com registros processados.")
            ScriptManager.GetCurrent(Me.Page).SetFocus(Me.txtPeriodoFinalConsultaTitulos)
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Finally
            strm.Close()
        End Try
        '' rotina utilizada nas procuracoes fim - html 
        '' rotina de geracao do relatorio . 
    End Sub

    Private Sub CarteiraDia()
        Dim Cliente As String
        Dim Campo() As String
        Dim i As Integer = 0
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue
        'Dim dra As DataRow
        Dim sql As String
        Dim lista As String
        Dim NomeArquivo As String = "Files/" & Funcoes.GeraNomeArquivo & ".html"
        Dim arquivo As String = HttpContext.Current.Server.MapPath(NomeArquivo)
        Dim strm As IO.StreamWriter
        If Dir(arquivo).Length > 0 Then Kill(arquivo)
        Dim dsRelatorio As New DataSet
        Dim dr As DataRow
        Dim linha As String
        Dim dataproc As Date
        Dim conterro As Integer
        Dim contreg As Integer
        dataproc = Date.Today
        Dim ds As New DataSet
        'Dim Inconsist As String
        'Dim Sitban As Integer
        Dim Valdia As Decimal
        'Dim ValdiaDolar As Decimal
        Dim Valtot As Decimal
        Dim Valtotdolar As Decimal
        Dim Valtotdia As Decimal
        Dim Valtotdiadolar As Decimal
        Dim Valind As Decimal
        Dim Datvenctr As String
        Dim Empresa As String
        Dim carteira As String
        'Dim registro As String

        lista = "Geral"

        sql = "  SELECT  MovimentacoesFinanceiras.Registro_Id AS Registro, " & vbCrLf & _
              " convert(varchar(10),MovimentacoesFinanceiras.Vencimento,103) as Vencimento,convert(varchar(10)," & vbCrLf & _
              " MovimentacoesFinanceiras.Prorrogacao,103) as Posterga, Clientes.Nome AS Cliente, Historico, " & vbCrLf & _
              " isnull(MovimentacoesFinanceiras.MoedaValorLiquido, 0) AS Dolar, " & vbCrLf & _
              " MovimentacoesFinanceiras.ValorLiquido AS Valor, " & vbCrLf & _
              " UsuarioLiberacao as Liberado, " & vbCrLf & _
              " MovimentacoesFinanceiras.Carteira As Carteira, " & vbCrLf & _
              " ComprasXProdutos.Descricao As DescricaoCarteira, " & vbCrLf & _
              " MovimentacoesFinanceiras.Situacao as Situacao, " & vbCrLf & _
              " MovimentacoesFinanceiras.Empresa as Empresa, " & vbCrLf & _
              " Empresa.Reduzido as Reduzido, " & vbCrLf & _
              " MovimentacoesFinanceiras.UsuarioBaixa as UsuarioBaixa" & vbCrLf & _
              " FROM MovimentacoesFinanceiras INNER JOIN" & vbCrLf & _
              " Clientes ON MovimentacoesFinanceiras.Cliente = Clientes.Cliente_Id AND MovimentacoesFinanceiras.EndCliente = Clientes.Endereco_Id" & vbCrLf & _
              " INNER JOIN" & vbCrLf & _
              " ComprasXProdutos ON MovimentacoesFinanceiras.Carteira = ComprasXProdutos.Produto_id " & vbCrLf & _
              " INNER JOIN" & vbCrLf & _
              " Clientes as Empresa ON MovimentacoesFinanceiras.Empresa = Empresa.Cliente_Id AND MovimentacoesFinanceiras.EndEmpresa = Empresa.Endereco_Id" & vbCrLf & _
              " WHERE MovimentacoesFinanceiras.Provisao <> 1 and  MovimentacoesFinanceiras.Situacao = 1" & vbCrLf

        If RbGeral.Checked = True Then
            '' Nao ira fazer nada pois vai listar todos 
            lista = "Geral"
        End If

        If RbAtivo.Checked = True Then
            sql &= " and MovimentacoesFinanceiras.usuariobaixa = '' "
            lista = "Ativos"
        End If

        If RbBaixado.Checked = True Then
            sql &= " and MovimentacoesFinanceiras.usuariobaixa <> '' "
            lista = "Baixados"
        End If

        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            sql &= " and MovimentacoesFinanceiras.UnidadeDeNegocio = '" & Cliente & "' "
        End If

        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            sql &= " and MovimentacoesFinanceiras.Empresa = '" & Campo(0) & "'"   'Empresa
            sql &= " and MovimentacoesFinanceiras.EndEmpresa = " & Campo(1)       'Endereco da Empresa
        End If

        'Cliente = DdlClienteConsultaTitulos.SelectedValue
        'Campo = Cliente.Split("-")
        'If Campo(0) <> "" Then
        '    sql &= " and MovimentacoesFinanceiras.Cliente = '" & Campo(0) & "'"   'Cliente
        '    sql &= " and MovimentacoesFinanceiras.EndCliente = " & Campo(1)       'Cliente da Empresa
        'End If

        If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
            sql &= " and Vencimento between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
        End If

        'If ChkAutozizado.Checked = True Then
        ' sql &= " and MovimentacoesFinanceiras.UsuarioLiberacao <> ''"    'Autorizados
        'End If

        sql &= " ORDER BY MovimentacoesFinanceiras.Carteira, MovimentacoesFinanceiras.Vencimento "

        ''DS = Banco.ConsultaDataSet(Sql, "Contas")

        dsRelatorio = Banco.ConsultaDataSet(sql, "Contas")
        linha = "<HTML>" & vbCrLf
        ''<HEAD>
        linha &= "<HEAD>" & vbCrLf & _
                 "<META HTTP-EQUIV='Content-Type' CONTENT='text/html; charset=iso-8859-1'>" & vbCrLf & _
                 "<TITLE> Posicao de titulos processados em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TITLE>" & vbCrLf & _
                 "</HEAD>" & vbCrLf
        '<BODY>
        linha &= "<BODY>" & vbCrLf
        'Cabeçalho Padrao
        linha &= "<TR>"
        linha &= "<TD >" & "Posicao de titulos em  " & dataproc & " .Referente ao periodo de " & txtPeriodoInicialConsultaTitulos.Text & " A " & txtPeriodoFinalConsultaTitulos.Text & "!!!" & "</TD>"
        linha &= "</TR>"
        ''**************************** Controle de cabecalho
        Cliente = DdlUnidadeConsultaTitulos.SelectedValue
        If Cliente <> "" Then
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Unidade : " & DdlUnidadeConsultaTitulos.SelectedItem.Text & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        Else
            linha &= "<TR>" & vbCrLf & _
                    "<TD >" & " Unidade : " & "Todas as Unidades !!!" & " </TD>" & vbCrLf & _
                    "</TR>" & vbCrLf
        End If
        Cliente = DdlEmpresaConsultaTitulos.SelectedValue
        Campo = Cliente.Split("-")
        If Campo(0) <> "" Then
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Empresa : " & DdlEmpresaConsultaTitulos.SelectedItem.Text & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        Else
            linha &= "<TR>" & vbCrLf & _
                     "<TD >" & " Empresa : " & "Todas as Empresas !!!" & " </TD>" & vbCrLf & _
                     "</TR>" & vbCrLf
        End If

        'Cliente = DdlClienteConsultaTitulos.SelectedValue
        'Campo = Cliente.Split("-")
        'If Campo(0) <> "" Then
        '    linha &= "<TR>"
        '    linha &= "<TD >" & " Cliente : " & DdlClienteConsultaTitulos.SelectedValue & " </TD>"
        '    linha &= "</TR>"
        'Else
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & " Cliente : " & "Todos os Clientes" & " </TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        'End If
        ''**************************** Controle de cabecalho fim 
        linha &= "<TR>" & vbCrLf & _
                 "<TD >" & " Tipo de relatorio Filial(Empresa)Geral - Totalizacao por Filial - Registros impressos : " & lista & " </TD>" & vbCrLf & _
                 "</TR>" & vbCrLf & _
                 "<table width= '370' cellpadding='0' cellspacing='0' Border=0>" & vbCrLf & _
                 "<TR>" & vbCrLf & _
                 "<TD ><B>Registro</B></TD>" & vbCrLf & _
                 "<TD ><B>Cliente/Fornecedor</B></TD>" & vbCrLf & _
                 "<TD ><B>Historico</B></TD>" & vbCrLf & _
                 "<TD ><B>Receber R$</B></TD>" & vbCrLf & _
                 "<TD ><B>Receber US$</B></TD>" & vbCrLf & _
                 "<TD ><B>Empresa</B></TD>" & vbCrLf & _
                 "<TD ><B>Vencimento Original</B></TD>" & vbCrLf & _
                 "<TD ><B>Prorrogacao</B></TD>" & vbCrLf & _
                 "<TD ><B>Carteira<B></TD>" & vbCrLf & _
                 "<TD ><B>Situacao</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf

        conterro = 0
        contreg = 0
        Valind = 0
        Valtot = 0
        Valdia = 0
        Datvenctr = ""
        Empresa = ""
        carteira = ""
        If dsRelatorio.Tables(0).Rows.Count > 0 Then
            For Each dr In dsRelatorio.Tables(0).Rows
                contreg = contreg + 1
                If carteira <> "" And carteira <> dr("DescricaoCarteira") Then

                    linha &= "<TR>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & " Total Da Carteira: " & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                             "</TR>" & vbCrLf
                    Valtotdia = 0
                    Valtotdiadolar = 0
                End If
                Valtot = Valtot + dr("Valor")
                Valtotdolar = Valtotdolar + dr("Dolar")
                Valtotdia = Valtotdia + dr("Valor")
                Valtotdiadolar = Valtotdiadolar + dr("Dolar")
                Empresa = dr("Empresa")
                Datvenctr = dr("Vencimento")
                carteira = dr("DescricaoCarteira")

                linha &= "<TR>" & vbCrLf & _
                         "<TD>" & dr("Registro") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Cliente") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Historico") & "</TD>" & vbCrLf & _
                         "<TD>" & Convert.ToDouble(dr("Valor")).ToString("n2") & "</TD>" & vbCrLf & _
                         "<TD>" & Convert.ToDouble(dr("Dolar")).ToString("n2") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Reduzido") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Vencimento") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Posterga") & "</TD>" & vbCrLf & _
                         "<TD>" & dr("Carteira") & " - " & dr("DescricaoCarteira") & "</TD>" & vbCrLf
                If dr("UsuarioBaixa") = "" Then
                    linha &= "<TD>" & "Ativo" & "</TD>"
                Else
                    linha &= "<TD>" & "Baixado" & "</TD>"
                End If
                linha &= "</TR>"
            Next
        End If

        linha &= "<TR>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & " Total Da Carteira: " & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdia.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdiadolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf & _
                 "<TR>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & " Total neste processamento: " & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtot.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & Valtotdolar.ToString("n2") & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "<TD><B>" & "." & "</B></TD>" & vbCrLf & _
                 "</TR>" & vbCrLf
        If contreg = 0 Then
            MsgBox(Me.Page, "Não existem registros corretos para o período.")
            ScriptManager.GetCurrent(Me.Page).SetFocus(Me.txtPeriodoFinalConsultaTitulos)
        Else
            MsgBox(Me.Page, "Movimento com registros processados.")
            ScriptManager.GetCurrent(Me.Page).SetFocus(Me.txtPeriodoFinalConsultaTitulos)
        End If

        strm = New IO.StreamWriter(arquivo, True)
        Try
            strm.WriteLine(linha)
            strm.Close()
            ScriptManager.RegisterClientScriptBlock(Me.Page, Me.Page.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo & "');", True)
        Finally
            strm.Close()
        End Try
        ' rotina utilizada nas procuracoes fim - html 
        ' rotina de geracao do relatorio . 
    End Sub

    'TITULO
    Protected Sub BtnSlip_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        '' rotina de imprimir slip por registros checados. .'.
        Dim xextenso As String = ""
        Dim yextenso As String = ""
        Dim dsEmitir As New DataSet
        'Dim dtEmitir As DataTable
        Dim row As DataRow
        'Dim i As Integer
        Dim j As Integer
        Dim k As Integer
        Dim mes As Integer
        Dim MESX As String = ""
        'Dim NumeroDoCheque As Integer
        'Dim NumeroDoChequeX As String
        Dim Historico As String = ""
        Dim RegistroI As String = ""
        Dim RegistroS As String = ""
        Dim SqlArray As New ArrayList
        'Dim Sqlupdate As String
        'Dim Campo() As String
        'Dim Mensagem As String
        ''*Definicao de campos da empresa
        Dim Empresa As String = ""
        Dim EndEmpresa As String = ""
        Dim ENome As String = ""
        Dim EEndereco As String = ""
        Dim ECep As String = ""
        Dim ECidade As String = ""
        Dim EEstado As String = ""
        Dim ECnpj As String = ""
        Dim EInscricao As String = ""
        Dim Efone As String = ""
        Dim EBairro As String = ""
        Dim EComplemento As String = ""
        Dim ENumero As Integer
        ''*Definicao de campos da empresa Fim 
        ''* DAdos do cliente inicio
        Dim Cliente As String = ""
        Dim EndCliente As String = ""
        Dim CNome As String = ""
        Dim CEndereco As String = ""
        Dim CCep As String = ""
        Dim CCidade As String = ""
        Dim CEstado As String = ""
        Dim CCnpj As String = ""
        Dim CInscricao As String = ""
        Dim Cfone As String = ""
        Dim CBairro As String = ""
        Dim CComplemento As String = ""
        Dim CNumero As Integer

        Dim CBancoCliente As String = ""
        Dim CAgenciaCliente As String = ""
        Dim CDigitoAgenciaCliente As String = ""
        Dim CCcontaCliente As String = ""
        Dim CDigitoContaCliente As String = ""

        ''* Dados do Cliente fim 
        ''* Campos do DAta Set + titulos inicio
        'Dim myDataRow As DataRow
        '' Cria data Set que vai ser utilizado no relatorio 
        Dim dsRecibo As New DataSet
        Dim dtRecibo As DataTable
        Dim TnumeroDoCheque As Integer
        ''Dim row As DataRow
        ''Dim RegistroI As String
        ''Dim RegistroS As String
        Try
            dtRecibo = New DataTable("ReciboAPagar")
            '' campos da empresa 
            dtRecibo.Columns.Add("ENome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEndereco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECep", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ECnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EInscricao", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EFone", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EBairro", Type.GetType("System.String"))
            dtRecibo.Columns.Add("EComplemento", Type.GetType("System.String"))
            dtRecibo.Columns.Add("ENumero", Type.GetType("System.Int32")).DefaultValue = 0
            ' ''campos do cliente
            dtRecibo.Columns.Add("CNome", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CEndereco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCep", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCidade", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CEstado", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CCnpj", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CInscricao", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CFone", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CBairro", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CComplemento", Type.GetType("System.String"))
            dtRecibo.Columns.Add("CNumero", Type.GetType("System.Int32")).DefaultValue = 0
            '' campos to titulo 
            dtRecibo.Columns.Add("TNumtit", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TValor", Type.GetType("System.Decimal"))
            dtRecibo.Columns.Add("TExtenso", Type.GetType("System.String"))
            dtRecibo.Columns.Add("THistorico", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TDia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TMes", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TAno", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TFormaPagto", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TBanco", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TAgencia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TConta", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TVencimento", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TBaixa", Type.GetType("System.DateTime"))
            dtRecibo.Columns.Add("TRecibo", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TRegistro", Type.GetType("System.Int32")).DefaultValue = 0
            dtRecibo.Columns.Add("TNumeroDoCheque", Type.GetType("System.Int32")).DefaultValue = 0
            '' Campos novos do cheque 05/05/2010
            dtRecibo.Columns.Add("TValorDoDocumento", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDescontos", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDeducoes", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TJuros", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TAcrescimos", Type.GetType("System.Decimal")).DefaultValue = 0
            dtRecibo.Columns.Add("TDigito", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TDigitoAgencia", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TContaContabilPagadora", Type.GetType("System.String"))
            dtRecibo.Columns.Add("TMovimento", Type.GetType("System.String"))

            '' Campos novos do cheque 05/05/2010
            dsRecibo.Tables.Add(dtRecibo)
            Dim ValorCobrado As Decimal
            Dim ValorDoDocumento As Decimal
            Dim Juros As Decimal
            Dim Acrescimos As Decimal
            Dim Descontos As Decimal
            Dim deducoes As Decimal
            Dim TipoPagto As Integer
            Dim FormaDePagamento As String = ""
            '' Campos novos do cheque 05/05/2010
            Dim TvalorDoDocumento As Decimal
            Dim Tdescontos As Decimal
            Dim Tdeducoes As Decimal
            Dim TJuros As Decimal
            Dim TAcrescimos As Decimal
            Dim Tdigito As String = ""
            '' Campos novos do cheque 05/05/2010
            ''* Campos do Data Set + titulos fim.


            'row = dtEmitir.NewRow()
            'dtEmitir.Rows.Add(row)
            '' Rotina para leitura de registro buscando a data da baixa. 
            Dim DataBaixa As String = ""
            Dim dataVencimento As String = ""
            Dim Tmovimento As String = ""
            Dim TContaContabilPagadora As String = ""
            ''Dim Dr2 As New DataRow
            Registro = GridConsultaTitulos.SelectedRow.Cells(1).Text()
            ''txtRegistro.Text = Registro

            Sql = " SELECT Registro_Id, Sequencia_Id, Provisao, Carteira, Tributo, Indexador, Moeda, TipoPagto, Situacao, Lote, Movimento, Vencimento, Prorrogacao, DataMoeda, Baixa, " & vbCrLf & _
                  "       UnidadeDeNegocio, Empresa, EndEmpresa, Cliente, EndCliente, BancoCliente, AgenciaCliente, DigitoAgenciaCliente, ContaCliente, DigitoContaCliente, " & vbCrLf & _
                  "       ContaContabilCliente, EmpresaPagadora, EndEmpresaPagadora, BancoPagador, AgenciaPagadora, DigitoAgenciaPagadora, ContaPagadora, DigitoContaPagadora, " & vbCrLf & _
                  "       ContaContabilPagadora, Cheque, Slips, Recibo, Aviso, ReciboDeposito, EmpresaPedido, EndEmpresaPedido, Pedido, PedidoFixacao, Procuracao, ValorDoDocumento, " & vbCrLf & _
                  "       Descontos, Deducoes, Juros, Acrescimos, ValorLiquido, ISNULL(MoedaValorDoDocumento, 0) AS MoedaValorDoDocumento, ISNULL(MoedaDescontos, 0) " & vbCrLf & _
                  "       AS MoedaDescontos, ISNULL(MoedaDeducoes, 0) AS MoedaDeducoes, ISNULL(MoedaJuros, 0) AS MoedaJuros, ISNULL(MoedaAcrescimos, 0) AS MoedaAcrescimos, " & vbCrLf & _
                  "       ISNULL(MoedaValorLiquido, 0) AS MoedaValorLiquido, Historico, CodigoDeBarras, CodigoDigitado, Destinatario, EndDestinatario, NomeDoDestinatario, Destinacao, " & vbCrLf & _
                  "       solicitacao, UsuarioInclusao, UsuarioInclusaoData, UsuarioAlteracao, UsuarioAlteracaoData, UsuarioCancelamento, UsuarioCancelamentoData, UsuarioLiberacao, " & vbCrLf & _
                  "       UsuarioLiberacaoData, UsuarioBaixa, UsuarioBaixaData, Grupado, isnull(RegistroMestre, 0) as RegistroMestre, Observacoes, SituacaoBancaria,T.Descricao, ISNULL(NumeroDoCheque, 0) AS NumeroDoCheque  " & vbCrLf & _
                  "  FROM MovimentacoesFinanceiras " & vbCrLf & _
                  " INNER JOIN TIPOSDEPAGAMENTOS AS T ON T.TipoDePagamento_ID = TipoPagto " & vbCrLf & _
                  " WHERE Registro_Id = " & Registro & vbCrLf

            For Each Dr1 As DataRow In Banco.ConsultaDataSet(Sql, "ContasAPagar").Tables(0).Rows
                DataBaixa = CDate(Dr1("Baixa")).ToSqlDate()
                dataVencimento = CDate(Dr1("Prorrogacao")).ToSqlDate()
                Empresa = Dr1("Empresa")
                EndEmpresa = Dr1("EndEmpresa")
                Cliente = Dr1("Cliente")
                EndCliente = Dr1("EndCliente")
                ValorCobrado = 0
                ValorDoDocumento = Dr1("ValorDoDocumento")
                Juros = Dr1("Juros")
                Acrescimos = Dr1("Acrescimos")
                Descontos = Dr1("Descontos")
                deducoes = Dr1("Deducoes")
                Historico = Dr1("Historico")
                CBancoCliente = Dr1("BancoCliente")
                CAgenciaCliente = Dr1("AgenciaCliente")
                CDigitoAgenciaCliente = Dr1("DigitoAgenciaCliente")
                CCcontaCliente = Dr1("ContaCliente")
                CDigitoContaCliente = Dr1("DigitoContaCliente")
                ValorCobrado = ValorDoDocumento + Juros + Acrescimos - Descontos - deducoes
                TipoPagto = Dr1("TipoPagto")
                FormaDePagamento = Dr1("Descricao")
                TnumeroDoCheque = Dr1("NumeroDoCheque")
                Tdescontos = Dr1("Descontos")
                Tdeducoes = Dr1("Deducoes")
                TJuros = Dr1("Juros")
                TAcrescimos = Dr1("Acrescimos")
                TvalorDoDocumento = Dr1("ValorDodocumento")
                Tdigito = Dr1("DigitoContaCliente")
                Tmovimento = CDate(Dr1("Movimento")).ToSqlDate()
                TContaContabilPagadora = Dr1("ContaContabilPagadora")
            Next

            Dim dr As DataRow
            '' Dados da empresa - fim 
            '' Consultado empresa 

            Sql = "  SELECT Clientes.Cliente_Id ," & vbCrLf & _
                  " Clientes.Nome, Clientes.Cidade," & vbCrLf & _
                  " Clientes.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                  " Clientes.Endereco , Clientes.Cep," & vbCrLf & _
                  " Clientes.Inscricao, Clientes.Telefone," & vbCrLf & _
                  " Clientes.Bairro, Clientes.Complemento," & vbCrLf & _
                  " Clientes.Numero " & vbCrLf & _
                  " FROM Clientes " & vbCrLf & _
                  " WHERE Clientes.Cliente_Id = '" & Empresa & "'" & vbCrLf
            DS = Banco.ConsultaDataSet(Sql, "Clientes")

            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    ENome = dr("Nome")
                    EEndereco = dr("Endereco")
                    ECep = dr("Cep")
                    ECidade = dr("Cidade")
                    EEstado = dr("Estado")
                    ECnpj = dr("Cliente_id")
                    EInscricao = dr("Inscricao")
                    Efone = dr("Telefone")
                    EBairro = dr("Bairro")
                    EComplemento = dr("Complemento")
                    ENumero = dr("Numero")
                    Exit For
                Next
            End If
            '' Consultando Empresa - fim 
            '' Consultado Cliente 

            Sql = "  SELECT Clientes.Cliente_Id ," & vbCrLf & _
                  " Clientes.Nome, Clientes.Cidade," & vbCrLf & _
                  " Clientes.Estado, 'CONSOLIDADO' as Descricao," & vbCrLf & _
                  " Clientes.Endereco , Clientes.Cep," & vbCrLf & _
                  " Clientes.Inscricao, Clientes.Telefone," & vbCrLf & _
                  " Clientes.Bairro, Clientes.Complemento," & vbCrLf & _
                  " Clientes.Numero " & vbCrLf & _
                  " FROM Clientes " & vbCrLf & _
                  " WHERE Clientes.Cliente_Id = '" & Cliente & "'and Clientes.Endereco_id = " & EndCliente & vbCrLf
            DS = Banco.ConsultaDataSet(Sql, "Clientes")

            If DS.Tables(0).Rows.Count > 0 Then
                For Each dr In DS.Tables(0).Rows
                    CNome = dr("Nome")
                    CEndereco = dr("Endereco")
                    CCep = dr("Cep")
                    CCidade = dr("Cidade")
                    CEstado = dr("Estado")
                    CCnpj = dr("Cliente_id")
                    CInscricao = dr("Inscricao")
                    Cfone = dr("Telefone")
                    CBairro = dr("Bairro")
                    CComplemento = dr("Complemento")
                    CNumero = dr("Numero")
                    Exit For
                Next
            End If

            ' Cria Data Sete que vai ser utilizado no relatorio
            ' Move campos para o Data Set.
            ' Move campos da Empresa
            row = dtRecibo.NewRow()
            row("ENome") = ENome
            row("EEndereco") = EEndereco
            row("ECep") = ECep
            row("ECidade") = ECidade
            row("EEstado") = EEstado
            row("ECnpj") = Funcoes.FormatarCpfCnpj(ECnpj)
            row("EInscricao") = EInscricao
            row("EFone") = Efone
            row("EBairro") = EBairro
            row("EComplemento") = EComplemento
            row("ENumero") = ENumero
            '' Move campos do Cliente / fornecedor
            row("CNome") = CNome
            row("CEndereco") = CEndereco
            row("CCep") = CCep
            row("CCidade") = CCidade
            row("CEstado") = CEstado
            row("CCnpj") = Funcoes.FormatarCpfCnpj(CCnpj)
            row("CInscricao") = CInscricao
            row("CFone") = Cfone
            row("CBairro") = CBairro
            row("CComplemento") = CComplemento
            row("CNumero") = CNumero
            '' Move campos do Titulo
            row("Tnumtit") = Registro
            '' feito
            row("TValor") = ValorCobrado
            row("TNumeroDoCheque") = TnumeroDoCheque
            '' campos novos 05/05/10
            row("TValorDoDocumento") = TvalorDoDocumento
            row("TDescontos") = Tdescontos
            row("TDeducoes") = Tdeducoes
            row("TJuros") = TJuros
            row("TAcrescimos") = TAcrescimos
            row("TDigito") = Tdigito
            row("TDigitoAgencia") = CDigitoAgenciaCliente

            '' campos novos 05/05/10

            Dim valcobradostr As String
            valcobradostr = CStr(ValorCobrado)

            Valor = Replace(valcobradostr, ".", "")

            ''* Rotina de extenso inicio
            yextenso = "("
            yextenso &= UCase(Funcoes.Extenso(valcobradostr, "Real", "Reais"))
            yextenso &= " *"
            xextenso = yextenso
            For j = 1 To (120 - Len(xextenso))
                xextenso &= " *"
            Next
            xextenso &= ")"
            row("TExtenso") = xextenso
            ''* Rotina de extenso fim 
            row("THistorico") = Historico
            row("TDia") = Day(DataBaixa)
            row("TAno") = Year(DataBaixa)

            row("TMes") = Month(DataBaixa)
            ''* Rotina do Mes inicio
            mes = Month(DataBaixa)
            If mes = 1 Then
                MESX = "JANEIRO"
            End If
            If mes = 2 Then
                MESX = "FEVEREIRO"
            End If
            If mes = 3 Then
                MESX = "MARCO"
            End If
            If mes = 4 Then
                MESX = "ABRIL"
            End If
            If mes = 5 Then
                MESX = "MAIO"
            End If
            If mes = 6 Then
                MESX = "JUNHO"
            End If
            If mes = 7 Then
                MESX = "JULHO"
            End If
            If mes = 8 Then
                MESX = "AGOSTO"
            End If
            If mes = 9 Then
                MESX = "SETEMBRO"
            End If
            If mes = 10 Then
                MESX = "OUTUBRO"
            End If
            If mes = 11 Then
                MESX = "NOVEMBRO"
            End If
            If mes = 12 Then
                MESX = "DEZEMBRO"
            End If

            row("TMes") = MESX
            ''* rotina do mes fim 


            row("TFormaPagto") = FormaDePagamento
            row("TBanco") = CBancoCliente
            row("TAgencia") = CAgenciaCliente
            row("TConta") = CCcontaCliente
            row("TVencimento") = dataVencimento
            row("TBaixa") = DataBaixa

            '' calculo de registro . 
            RegistroI = Registro
            RegistroS = " "
            RegistroS = Trim(RegistroS)
            If Len(RegistroS) < 6 Then
                For k = 1 To (6 - Len(Trim(RegistroI)))
                    RegistroS &= "0"
                Next
            End If
            row("Trecibo") = Trim(RegistroS) & (Trim(RegistroI))
            row("TRegistro") = Trim(RegistroS) & (Trim(RegistroI))

            row("TContaContabilPagadora") = TContaContabilPagadora
            row("TMovimento") = Tmovimento

            dtRecibo.Rows.Add(row)

            'Imagem
            Dim dtImagem As New DataTable("Images")
            dtImagem.Columns.Add("path", GetType(String))
            dtImagem.Columns.Add("image", GetType(System.Byte()))

            Dim drImagem As DataRow = dtImagem.NewRow()
            Dim strCaminhoImagem As String = Server.MapPath("~/Images/" & HttpContext.Current.Session("ssImagemEmpresa"))

            drImagem("path") = strCaminhoImagem
            drImagem("image") = [Lib].Negocio.Conversoes.ConverterImagemEmByte(strCaminhoImagem)
            dtImagem.Rows.Add(drImagem)

            dsRecibo.Tables.Add(dtImagem)

            '' Consultando Empresa - fim 
            '' Emissao de relatorio 
            Dim crpt As New ReportDocument()

            Try
                crpt.FileName = Server.MapPath("~/Reports/Cr_slip.rpt")
                crpt.Load(crpt.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

                Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
                Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
                Dim arquivo As String = NomeArquivo

                crpt.SetDataSource(dsRecibo)

                Dim crparametervalues As ParameterValues
                Dim crparameterdiscretevalue As ParameterDiscreteValue
                Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
                Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

                crparameterfielddefinitions = crpt.DataDefinition.ParameterFields()

                crparameterfielddefinition = crparameterfielddefinitions.Item("XNome")
                crparametervalues = crparameterfielddefinition.CurrentValues
                crparameterdiscretevalue = New ParameterDiscreteValue
                crparameterdiscretevalue.Value = ENome
                crparametervalues.Add(crparameterdiscretevalue)
                crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                crpt.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)
                ScriptManager.RegisterClientScriptBlock(Me, Me.GetType(), Guid.NewGuid().ToString(), "window.open('" & NomeArquivo2 & "');", True)
            Catch ex As Exception
                MsgBox(Me.Page, ex.Message)
            Finally
                crpt.Close()
                crpt.Dispose()
            End Try
            '' rotina de imprimir slip  por registros checados fim .
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("MovimentacoesFinanceiras", "GRAVAR") Then
                GravaTitulo()
            Else
                MsgBox(Me.Page, "usuário sem permissão para gravar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Dim SqlArray As New ArrayList

        Try
            If Funcoes.VerificaPermissao("MovimentacoesFinanceiras", "EXCLUIR") Then
                If txtRegistro.Text <> "" Then
                    Registro = txtRegistro.Text

                    Sql = " UPDATE MovimentacoesFinanceiras" & vbCrLf & _
                          " SET Situacao = 3" & vbCrLf

                    Sql &= ", UsuarioCancelamento = '" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf & _
                           ", UsuarioCancelamentoData = '" & Today.ToSqlDate() & "'" & vbCrLf

                    Sql &= " WHERE Registro_ID = " & CStr(Registro)
                    SqlArray.Add(Sql)

                    Registro = CInt(txtRegistro.Text)

                    Sql = "DELETE FROM razao" & vbCrLf & _
                          " WHERE Titulo = " & CInt(txtRegistro.Text) & vbCrLf
                    SqlArray.Add(Sql)

                    If Banco.GravaBanco(SqlArray) = False Then
                        Mensagem = Replace(HttpContext.Current.Session("ssMessage"), "'", "")
                    Else
                        Limpar()
                        Mensagem = "Registro < " & Registro & " > Excluído com sucesso."
                        MsgBox(Me.Page, Mensagem)
                    End If
                Else
                    MsgBox(Me.Page, "Informe o número do registro para excluir.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            ConsultaMovimentacoesFinanceiras()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "MovimentacoesFinanceiras")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda")
        End Try
    End Sub

    'CONSULTA TITULO
    Protected Sub lnkConsultarCT_Click(sender As Object, e As EventArgs) Handles lnkConsultarCT.Click
        Dim Cliente As String
        Dim Campo() As String
        Dim Unidade As String = DdlUnidadeConsultaTitulos.SelectedValue

        Try
            Sql = "  SELECT  MovimentacoesFinanceiras.Registro_Id AS Registro, convert(varchar(10),MovimentacoesFinanceiras.Vencimento,103) as Vencimento, Clientes.Nome AS Cliente, MovimentacoesFinanceiras.ValorDoDocumento AS Valor"

            Sql &= " FROM MovimentacoesFinanceiras INNER JOIN" & vbCrLf & _
                   " Clientes ON MovimentacoesFinanceiras.Empresa = Clientes.Cliente_Id AND MovimentacoesFinanceiras.EndEmpresa = Clientes.Endereco_Id" & vbCrLf
            ' ROTINA ANTIGA
            'Sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1 "
            Sql &= " WHERE MovimentacoesFinanceiras.Situacao = 1 "
            If RbAtivo.Checked Then
                Sql &= " and Provisao <> 1 "
            End If
            If RbBaixado.Checked Then
                Sql &= " and Provisao = 1 "
            End If
            Cliente = DdlUnidadeConsultaTitulos.SelectedValue
            If Cliente <> "" Then
                Sql &= " and MovimentacoesFinanceiras.UnidadeDeNegocio = '" & Cliente & "' "
            End If

            Cliente = DdlEmpresaConsultaTitulos.SelectedValue
            Campo = Cliente.Split("-")
            If Campo(0) <> "" Then
                Sql &= " and MovimentacoesFinanceiras.Empresa = '" & Campo(0) & "'" & vbCrLf & _
                       " and MovimentacoesFinanceiras.EndEmpresa = " & Campo(1) & vbCrLf
            End If

            If txtPeriodoInicialConsultaTitulos.Text <> "" And txtPeriodoFinalConsultaTitulos.Text <> "" Then
                Sql &= " and Vencimento between '" & txtPeriodoInicialConsultaTitulos.Text.ToSqlDate() & "' and '" & txtPeriodoFinalConsultaTitulos.Text.ToSqlDate() & "'"
            End If

            Sql &= " ORDER BY MovimentacoesFinanceiras.Vencimento, Clientes.Nome"
            DS = Banco.ConsultaDataSet(Sql, "Contas")

            If DS Is Nothing Then
                MsgBox(Me.Page, "Nenhum registro encontrado.")
            ElseIf DS.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Nenhum registro encontrado.")
            Else
                GridConsultaTitulos.DataSource = Banco.ConsultaDataSet(Sql, "Contas")
                GridConsultaTitulos.DataBind()
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimparCT_Click(sender As Object, e As EventArgs) Handles lnkLimparCT.Click
        Try
            txtPeriodoInicialConsultaTitulos.Text = ""
            txtPeriodoFinalConsultaTitulos.Text = ""

            GridConsultaTitulos.DataSource = Nothing
            GridConsultaTitulos.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkRelatorioCT_Click(sender As Object, e As EventArgs) Handles lnkRelatorioCT.Click
        Try
            If Funcoes.VerificaPermissao("MovimentacoesFinanceiras", "RELATORIO") Then
                If RbDiaGeral.Checked = True Then
                    DiaGeral()
                End If
                If RbFilialDiario.Checked = True Then
                    FilialGeral()
                End If
                If RbCarteiraDia.Checked = True Then
                    CarteiraDia()
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

End Class