Imports System.Data
Imports System.IO
Imports System.Web
Imports System.Net
Imports System.Net.Mail
Imports System.Data.SqlClient
Imports Microsoft.VisualBasic
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class PagamentoDeTitulos
    Inherits BasePage

    Private objCliente As [Lib].Negocio.Cliente

    Dim dsCarteira As DataSet
    Dim Pagador As BoletoNet.Cedente
    Dim Parametros As String
    Dim SequenciaDePagamentoRecuperado As Integer = 0
    Dim SequenciaDeRegistroDePagamentoRecuperado As Integer = 0

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Financeiro)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("PagamentoDeTitulos", "ACESSAR") Then
                    If Not Directory.Exists(Server.MapPath("RemessaBancaria")) Then Directory.CreateDirectory(Server.MapPath("RemessaBancaria"))
                    CargaEmpresas()
                    Funcoes.VerificaEmpresa(ddlEmpresa)
                    CargaCarteiras()
                    CargaSituacao()
                    Limpar()
                    SetarData(txtDataInicial, txtDataFinal, False)
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar para essa página.", "~/Financeiro.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub CargaEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.ClientesXEmpresas, "", True)
    End Sub

    Private Sub CarregarReceber()
        Dim Empresa As String() = ddlEmpresa.SelectedValue.ToString.Split("-")
        Dim objReceber As New [Lib].Negocio.ListBancosXContas(True, Empresa(0))
        Dim agencia As String = ""
        Dim conta As String = ""

        ddlReceber.Items.Clear()

        For Each row As [Lib].Negocio.BancosXContas In objReceber
            agencia = IIf(row.DigitoAgencia.Length > 0, Funcoes.AlinharDireita(row.Agencia, 4, " ") & "-" & row.DigitoAgencia, Funcoes.AlinharDireita(row.Agencia, 6, "0"))
            conta = IIf(row.DigitoConta.Length > 0, Funcoes.AlinharDireita(row.Conta, 10, " ") & "-" & row.DigitoConta, Funcoes.AlinharDireita(row.Conta, 10, " "))
            ddlReceber.Items.Add(New ListItem(Funcoes.AlinharEsquerda(row.NomeBanco, 20, ".") & " - AG: " & agencia & " - CTA: " & conta, _
                                 row.CodigoBanco & ";" & row.Agencia & ";" & row.DigitoAgencia & ";" & row.Conta & ";" & row.DigitoConta & ";" & row.TipoConta))
        Next

        Funcoes.InserirLinhaEmBranco(ddlReceber)
    End Sub

    Private Sub CargaSituacao()
        Dim objSituacao As New [Lib].Negocio.ListSituacao(True)

        ddlSituacao.Items.Clear()

        For Each row As [Lib].Negocio.Situacao In objSituacao
            ddlSituacao.Items.Add(New ListItem(Funcoes.AlinharDireita(row.Codigo, 3, "0") & " - " & row.Descricao, row.Codigo))
        Next

        Funcoes.InserirLinhaEmBranco(ddlSituacao)
    End Sub

    Private Sub CargaCarteiras()
        Dim objCarteira As New [Lib].Negocio.ListCarteiraDoTitulo
        ddlCarteira.Items.Clear()
        For Each row As [Lib].Negocio.CarteiraDoTitulo In objCarteira
            ddlCarteira.Items.Add(New ListItem(Funcoes.AlinharDireita(row.Codigo, 3, "0") & " - " & row.Descricao, row.Codigo))
        Next
        Funcoes.InserirLinhaEmBranco(ddlCarteira)
    End Sub

    Private Sub Limpar()
        Session.Remove("objTitulos")
        gridPagamentoDeTitulos.DataBind()
        txtNumValor.Text = 0
        ucConsultaDadosBancarios.SetarHID(HID.Value)
        lblQtdeEValorTotal.Text = String.Empty
        btnDownload.Visible = False
        ddlCarteira.ClearSelection()
        ddlReceber.ClearSelection()
        ddlSituacao.ClearSelection()
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
        LiberaEmpresa()
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Protected Sub ddlEmpresa_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If ddlEmpresa.SelectedIndex > 0 Then CarregarReceber()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        Try
            If Not Session("objClientePAGTIT" & HID.Value) Is Nothing Then
                Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClientePAGTIT" & HID.Value), [Lib].Negocio.Cliente))
                txtNomeCliente.Text = itemCliente.Text
                txtCodigoCliente.Value = itemCliente.Value
                Session.Remove("objClientePAGTIT" & HID.Value)
            ElseIf Not Session("objClientePgtoTitulo" & HID.Value) Is Nothing Then
                objCliente = CType(Session("objClientePgtoTitulo" & HID.Value), [Lib].Negocio.Cliente)
                txtNomeCliente.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoCliente.Value = objCliente.Codigo & ";" & objCliente.CodigoEndereco
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub Consultar()
        CarregarDataGrid()
        If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
            MsgBox(Me.Page, "Não existem dados para esta coleção de parametros")
        End If
        txtNumValor.Text = SomaTitulosRemessa().ToString
    End Sub

    Public Function SomaTitulosRemessa() As Decimal
        Dim Total As Decimal = 0
        Dim Quantidade As Integer = 0
        For Each row As DataRow In CType(Session("objTitulos"), DataSet).Tables(0).Rows
            'Se estiver Marcado pra ser enviado
            If row.Item("Enviar").ToString.ToUpper = "TRUE" AndAlso row("TAG") <> "MT-PE" AndAlso row("TAG") <> "PE" Then
                Try
                    Total += row.Item("OficialProgramado") + row.Item("Acrescimos") + row.Item("Juros") - row.Item("Deducoes") - row.Item("Descontos")
                    Quantidade = Quantidade + 1
                Catch ex As Exception
                    Return 0
                End Try
            End If
        Next
        lblQtdeEValorTotal.Text = Quantidade & " Registro(s) selecionado(s) no valor total de  " & String.Format("{0:N2}", Total)
        txtNumValor.Text = Total
        Return Total
    End Function

#Region "Acesso ao DataGrid"
    ' Carrega o DataGrid com as informacoes do banco de dados
    Private Sub CarregarDataGrid()
        If FinanceiroNovo Then
            dsCarteira = Banco.ConsultaDataSet(GetSqlFinanceiroNovo(), "Titulos")
        Else
            dsCarteira = Banco.ConsultaDataSet(GetSqlPreencheDataGrid(), "Titulos")
        End If
        Session("objTitulos") = dsCarteira
        gridPagamentoDeTitulos.DataSource = Session("objTitulos")
        gridPagamentoDeTitulos.DataBind()
        'Ordena()
    End Sub

#End Region

#Region "SQL"

    ' Monta o Select para preencher o DataGrid
    Private Function GetSqlPreencheDataGrid() As String
        Dim sql As String
        Parametros = ""

        sql = "   SELECT Case When CP.DataEnvio is null and CP.TipoPagto not in (1,2) Then 'TRUE' Else 'FALSE' end as Enviar, " & vbCrLf

        'TED E DOC SEM DADOS BANCARIOS
        sql &= "  isnull(Case When (CP.DataEnvio is null and CP.RegistroMestre is not null and CP.TipoPagto in (4,5,6)) " & vbCrLf & _
               "  and (CP.BancoPagador IS NULL OR CP.AgenciaPagadora IS NULL OR CP.DigitoAgenciaPagadora IS NULL OR CP.ContaPagadora IS NULL OR CP.DigitoContaPagadora IS NULL)" & vbCrLf & _
               "  Then 'MT-PE' " & vbCrLf & _
               "  When (CP.DataEnvio is null and CP.TipoPagto in (4,5,6)) and (CP.BancoPagador IS NULL OR CP.AgenciaPagadora IS NULL " & vbCrLf & _
               "  OR CP.DigitoAgenciaPagadora IS NULL OR CP.ContaPagadora IS NULL OR CP.DigitoContaPagadora IS NULL)" & vbCrLf & _
               "  Then 'PE'" & vbCrLf

        'BOLETO SEM CODIGO DE BARRA = STATUS PE 
        sql &= "  When CP.DataEnvio is null and CP.RegistroMestre is not null and CP.TipoPagto in (3,7,8) and CP.CodigoDeBarras is null" & vbCrLf & _
               "  Then 'MT-PE'" & vbCrLf & _
               "  When CP.DataEnvio is null and CP.TipoPagto in (3,7,8) and CP.CodigoDeBarras is null" & vbCrLf & _
               "  Then 'PE'" & vbCrLf & _
               "  When CP.DataEnvio is null and CP.RegistroMestre is not null" & vbCrLf & _
               "  Then 'MT'" & vbCrLf & _
               "  When CP.DataEnvio is not null and CP.RegistroMestre is not null" & vbCrLf & _
               "  Then 'RM-MT'" & vbCrLf & _
               "  When CP.DataEnvio is not null" & vbCrLf & _
               "  Then 'RM'" & vbCrLf & _
               "  end,'  ') as TAG," & vbCrLf & _
               "  CP.Registro_Id as Titulo_Id," & vbCrLf & _
               "  CP.Provisao," & vbCrLf & _
               "  P.Descricao as ProvisaoDesc," & vbCrLf & _
               "  isnull(CP.CarteiraDoTitulo,0) as Carteira," & vbCrLf & _
               "  C.Descricao as DescCarteira," & vbCrLf & _
               "  CP.Indexador," & vbCrLf & _
               "  i.Descricao as IndexadorDesc," & vbCrLf & _
               "  CP.TipoPagto," & vbCrLf & _
               "  Tp.Descricao as TipoPagtoDesc," & vbCrLf & _
               "  CP.Situacao," & vbCrLf & _
               "  S.Descricao as SituacaoDesc," & vbCrLf & _
               "  CP.Baixa AS Movimento ," & vbCrLf & _
               "  CP.Vencimento," & vbCrLf & _
               "  CP.Prorrogacao," & vbCrLf & _
               "  CP.DataMoeda," & vbCrLf & _
               "  CP.Baixa," & vbCrLf & _
               "  Origem.Reduzido as ReduzidoOrigem," & vbCrLf & _
               " Origem.Endereco AS EnderecoOrigem," & vbCrLf & _
               " Origem.Numero AS NumeroOrigem," & vbCrLf & _
               " Origem.Complemento AS ComplementoOrigem," & vbCrLf & _
               " Origem.Bairro AS BairroOrigem," & vbCrLf & _
               " Origem.Cep AS CepOrigem," & vbCrLf & _
               " Origem.Cidade AS CidadeOrigem," & vbCrLf & _
               " Origem.Estado AS EstadoOrigem," & vbCrLf & _
               "  CP.Empresa as EmpresaOrigem," & vbCrLf & _
               "  isnull(CP.EndEmpresa,0) as EnderecoOrigem," & vbCrLf & _
               "  Origem.Nome as EmpresaOrigemNome," & vbCrLf & _
               "  CP.Destinatario," & vbCrLf & _
               "  isnull(CP.EndDestinatario,0) as EnderecoDestino," & vbCrLf & _
               "  CP.EmpresaPagadora," & vbCrLf & _
               "  isnull(CP.EndEmpresaPagadora,0) as EnderecoDebito," & vbCrLf & _
               "  CP.ContaContabilPagadora," & vbCrLf & _
               "  CP.Cliente as ClienteDebito," & vbCrLf & _
               "  isnull(CP.EndCliente,0) as EndClienteDebito," & vbCrLf & _
               "  isnull(CP.EmpresaPagadora,'') as EmpresaCredito," & vbCrLf & _
               "  isnull(CP.EndEmpresaPagadora,0) as EnderecoCredito," & vbCrLf & _
               "  isnull(CP.ContaContabilCliente,'') as ContaCredito," & vbCrLf & _
               "  isnull(CP.Cliente,'') as ClienteCredito," & vbCrLf & _
               "  isnull(CP.EndCliente,0) as EndClienteCredito," & vbCrLf & _
               "  2 as ReceberPagar," & vbCrLf & _
               "  case     when CP.Cheque is null or CP.Cheque = 'N'" & vbCrLf & _
               "  then 'FALSE'			     Else 'TRUE'" & vbCrLf & _
               "  end Cheque," & vbCrLf & _
               "  case     when CP.Slips is null or CP.Slips = 'N'" & vbCrLf & _
               "  then 'FALSE'			     Else 'TRUE'" & vbCrLf & _
               "  end Slips," & vbCrLf & _
               "  case    when CP.Recibo is null or CP.Recibo = 'N'" & vbCrLf & _
               "  then 'FALSE'			     Else 'TRUE'" & vbCrLf & _
               "  end Recibo," & vbCrLf & _
               "  case   when CP.Aviso is null or CP.Aviso = 'N' then 'FALSE'" & vbCrLf & _
               "  Else 'TRUE'		     end Aviso," & vbCrLf & _
               "  case when CP.ReciboDeposito is null or CP.ReciboDeposito = 'N'  then 'FALSE'" & vbCrLf & _
               "  Else 'TRUE'		     end ReciboDeposito," & vbCrLf & _
               "  isnull(CP.EmpresaPedido,'') as EmpresaPedido," & vbCrLf & _
               "  isnull(CP.EndEmpresaPedido,-1) as EndEmpresaPedido," & vbCrLf & _
               "  isnull(CP.Pedido,'') as Pedido," & vbCrLf & _
               "  isnull(CP.PedidoFixacao,'') as PedidoFixacao," & vbCrLf & _
               "  isnull(CP.ValorDoDocumento,0) as OficialProgramado," & vbCrLf & _
               "  isnull(CP.ValorLiquido,0) as OficialBaixado," & vbCrLf & _
               "  isnull(CP.MoedaValorDoDocumento,0) as MoedaProgramado," & vbCrLf & _
               "  isnull(CP.MoedaValorLiquido,0) as MoedaBaixado," & vbCrLf & _
               "  isnull(CP.Acrescimos,0) as Acrescimos," & vbCrLf & _
               "  isnull(CP.Descontos,0) as Descontos," & vbCrLf & _
               "  isnull(CP.Juros,0) as Juros," & vbCrLf & _
               "  isnull(CP.Deducoes,0) as Deducoes," & vbCrLf & _
               "  isnull(CP.ValorDoDocumento,0) + isnull(CP.Acrescimos,0) +  isnull(CP.Juros,0) - isnull(CP.Deducoes,0) - isnull(CP.Descontos,0)  as ValorParaPagamento," & vbCrLf & _
               "  CP.Historico," & vbCrLf & _
               "  Case  When len(CP.Destinatario)=14 " & vbCrLf & _
               "  then 2   else 1 		 end TipoInscricaoFornecedor," & vbCrLf & _
               "  Destinatario.Reduzido as ReduzidoFornecedor," & vbCrLf & _
               "  CP.Destinatario as Fornecedor, " & vbCrLf & _
               "  CP.EndDestinatario as EndFornecedor, " & vbCrLf & _
               "  Destinatario.Nome as NomeFornecedor," & vbCrLf & _
               "  Destinatario.Cidade as CidadeFornecedor," & vbCrLf & _
               "  Destinatario.Estado as EstadoFornecedor," & vbCrLf & _
               "  Destinatario.Numero as NumeroFornecedor," & vbCrLf & _
               "  Destinatario.Complemento as ComplementoFornecedor," & vbCrLf & _
               "  Destinatario.CEP as CEPFornecedor," & vbCrLf & _
               "  Destinatario.Bairro as BairroFornecedor," & vbCrLf & _
               "  Destinatario.Endereco as EnderecoFornecedor," & vbCrLf

        'Dados Recebedor 
        sql &= "  CP.BancoCliente,CP.AgenciaCliente, CP.DigitoAgenciaCliente, CP.ContaCliente, CP.DigitoContaCliente,  isnull(CP.TipoContaCliente,'C') AS TipoContaCliente," & vbCrLf & _
               "  'Banco:' + str(CP.BancoCliente)+ ' AG:' + CP.AgenciaCliente + ' ' + CP.DigitoAgenciaCliente + ' CONTA:' + CP.ContaCliente + ' ' + CP.DigitoContaCliente + ' ' + isnull(CP.TipoContaCliente,'C') AS DadosBancariosFornc,  " & vbCrLf

        'Dados Pagador
        sql &= "  isnull(CP.BancoPagador,0) as Banco," & vbCrLf & _
               "  B.Descricao as BancoDesc," & vbCrLf & _
               "  isnull(CP.AgenciaPagadora,0) as Agencia," & vbCrLf & _
               "  isnull(CP.DigitoAgenciaPagadora,'') as DigitoAgencia," & vbCrLf & _
               "  isnull(CP.ContaPagadora,'') as ContaCorrente," & vbCrLf & _
               "  isnull(CP.DigitoContaPagadora,'') as DigitoConta," & vbCrLf & _
               "  isnull(CP.Destinacao,'') as Destinacao," & vbCrLf & _
               "  'BANCO:' +  str(CP.BancoPagador) + ' AG:' + CP.AgenciaPagadora + ' ' + CP.DigitoAgenciaPagadora + ' CONTA:' + CP.ContaPagadora + ' ' + CP.DigitoContaPagadora AS DadosBancariosPagadora, " & vbCrLf & _
               "  CP.Moeda, " & vbCrLf & _
               "  M.Descricao as MoedaDesc," & vbCrLf & _
               "  isnull(Clientes.Reduzido,'') as ReduzidoDebito," & vbCrLf & _
               "  CP.ModalidadeDePagamento," & vbCrLf & _
               "  CP.NossoNumero," & vbCrLf & _
               "  CP.DigitoNossoNumero," & vbCrLf & _
               "  isnull(CP.NumeroDaRemessa,'') as NumeroDaRemessa," & vbCrLf & _
               "  isnull(CP.CodigoDeBarras,'') as CodigoDeBarras," & vbCrLf & _
               "  isnull(LEFT(CodigoDeBarras,3),0) as BancoNoCodigoDeBarras," & vbCrLf & _
               "  case when isnull(CP.CodigoDigitado,'FALSE')= 'S' then 'TRUE' ELSE 'FALSE' end as CodigoDigitado  " & vbCrLf & _
               "  FROM ContasAPagar as CP " & vbCrLf & _
               "  Left Join  Clientes		ON CP.EmpresaPagadora = Clientes.Cliente_Id	" & vbCrLf & _
               "  AND CP.EndEmpresaPagadora = Clientes.Endereco_Id " & vbCrLf & _
               "  left Join  Clientes as Origem " & vbCrLf & _
               "  ON CP.Empresa  = Origem.Cliente_Id " & vbCrLf & _
               "  AND CP.EndEmpresa = Origem.Endereco_Id	" & vbCrLf & _
               "  inner Join  Clientes as Destinatario " & vbCrLf & _
               "  ON CP.Destinatario    = Destinatario.Cliente_Id " & vbCrLf & _
               "  AND CP.EndDestinatario = Destinatario.Endereco_Id " & vbCrLf & _
               "  Left Join Provisoes as P " & vbCrLf & _
               "  on P.Provisao_Id = CP.Provisao " & vbCrLf & _
               "  Left Join Carteira as C " & vbCrLf & _
               "  on C.Carteira_id = isnull(CP.CarteiraDoTitulo,0) " & vbCrLf & _
               "  Left Join Indexadores as i " & vbCrLf & _
               "  on i.Indexador_id = CP.Indexador " & vbCrLf & _
               "  Left Join TiposDePagamentos as Tp " & vbCrLf & _
               "  on Tp.TipoDePagamento_id = CP.TipoPagto " & vbCrLf & _
               "  Left Join Situacoes as S " & vbCrLf & _
               "  on S.Situacao_id = CP.Situacao	" & vbCrLf & _
               "  Left Join Moedas as M " & vbCrLf & _
               "  on M.Moeda_id = CP.Moeda " & vbCrLf & _
               "  Left Join Bancos as B" & vbCrLf & _
               "  on B.Banco_id = CP.BancoPagador" & vbCrLf

        'Carrega Somente os titulos mestres ou os Independentes
        sql &= "  WHERE 1=1 " & vbCrLf & _
               "  AND (CP.RegistroMestre = CP.Registro_Id Or CP.RegistroMestre Is null)" & vbCrLf

        If ddlEmpresa.SelectedIndex > 0 Then
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            sql &= " AND CP.EmpresaPagadora = '" & strEmpresa(0) & "' " & vbCrLf & _
                   " AND CP.EndEmpresaPagadora = " & strEmpresa(1) & " " & vbCrLf
            Parametros &= "Pagador: " & ddlEmpresa.Text
        End If

        If txtCodigoCliente.Value.Length > 0 Then
            Dim strCliente As String() = txtCodigoCliente.Value.Split(";")
            sql &= " AND CP.Destinatario = '" & strCliente(0) & "' " & vbCrLf & _
                   " AND CP.EndDestinatario = " & strCliente(1) & " " & vbCrLf
            Parametros &= "Pagador: " & txtNomeCliente.Text
        End If

        If ddlCarteira.SelectedIndex > 0 Then
            sql &= "   and CP.CarteiraDoTitulo = '" & ddlCarteira.SelectedValue & "'" & vbCrLf
        End If

        If ddlSituacao.SelectedIndex <> 0 Then
            If ddlSituacao.SelectedIndex = 1 Then
                sql &= "  AND CP.Situacao  not in (2,3,4,5,6,7,8,9)" & vbCrLf
            Else
                Dim strSituacao As String() = ddlSituacao.SelectedValue.Split("-")
                sql &= " AND CP.Situacao = " & strSituacao(0) & " " & vbCrLf
            End If
        End If

        If rdEmissao.Checked Then
            sql &= "AND CP.Baixa " & vbCrLf
        Else
            sql &= "AND CP.Prorrogacao " & vbCrLf
        End If

        sql &= " BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

        Return sql
    End Function

    Private Function GetSqlFinanceiroNovo() As String
        Dim strSQL As String

        strSQL = " SELECT " & vbCrLf & _
                " CASE WHEN TTB.DataEnvio IS NULL and T.TipoPagto not in (1,2)  " & vbCrLf & _
                " THEN 'TRUE' ELSE 'FALSE' END Enviar, " & vbCrLf & _
                " ISNULL(CASE	WHEN (TTB.DataEnvio IS NULL AND T.RegistroMestre IS NOT NULL AND T.TipoPagto IN(4,5,6)) " & vbCrLf & _
                " AND (T.BancoCliFor IS NULL OR T.AgenciaCliFor IS NULL OR T.DigitoAgenciaCliFor IS NULL OR T.ContaCliFor IS NULL OR T.DigitoContaCliFor IS NULL) " & vbCrLf & _
                " THEN 'MT-PE' " & vbCrLf & _
                " WHEN (TTB.DataEnvio IS NULL AND T.TipoPagto IN(4,5,6)) AND (T.BancoCliFor IS NULL OR T.AgenciaCliFor IS NULL " & vbCrLf & _
                " OR T.DigitoAgenciaCliFor IS NULL OR T.ContaCliFor IS NULL OR T.DigitoContaCliFor IS NULL) " & vbCrLf & _
                " THEN 'PE' " & vbCrLf & _
                " WHEN TTB.DataEnvio IS NULL AND T.RegistroMestre IS NOT NULL AND T.TipoPagto IN(3,7,8) AND T.CodigoDeBarra IS NULL " & vbCrLf & _
                " THEN 'MT-PE' " & vbCrLf & _
                " WHEN TTB.DataEnvio IS NULL AND T.TipoPagto IN(3,7,8) AND T.CodigoDeBarra IS NULL " & vbCrLf & _
                " THEN 'PE' " & vbCrLf & _
                " WHEN TTB.DataEnvio IS NULL AND T.RegistroMestre IS NOT NULL " & vbCrLf & _
                " THEN 'MT' " & vbCrLf & _
                " WHEN TTB.DataEnvio IS NOT NULL AND T.RegistroMestre IS NOT NULL " & vbCrLf & _
                " THEN 'RM-MT'  " & vbCrLf & _
                " WHEN TTB.DataEnvio IS NOT NULL " & vbCrLf & _
                " THEN 'RM' " & vbCrLf & _
                " END,'  ') AS TAG, " & vbCrLf & _
                " T.Titulo_Id, " & vbCrLf & _
                " T.Provisao, " & vbCrLf & _
                " P.Descricao AS ProvisaoDesc, " & vbCrLf & _
                " isnull(T.CarteiraDoTitulo, 0) as Carteira, " & vbCrLf & _
                " C.Descricao as DescCarteira, " & vbCrLf & _
                " T.Indexador, " & vbCrLf & _
                " I.Descricao AS IndexadorDesc, " & vbCrLf & _
                " T.TipoPagto, " & vbCrLf & _
                " TP.Descricao AS TipoPagtoDesc, " & vbCrLf & _
                " T.Situacao, " & vbCrLf & _
                " S.Descricao as SituacaoDesc, " & vbCrLf & _
                " T.Baixa as Movimento, " & vbCrLf & _
                " T.Vencimento, " & vbCrLf & _
                " T.Reprogramacao AS Prorrogacao, " & vbCrLf & _
                " T.DataMoeda, " & vbCrLf & _
                " T.Baixa, " & vbCrLf & _
                " ORIGEM.Reduzido as ReduzidoOrigem, " & vbCrLf & _
                " ORIGEM.Endereco AS EnderecoOrigem, " & vbCrLf & _
                " ORIGEM.Numero AS NumeroOrigem, " & vbCrLf & _
                " ORIGEM.Complemento AS ComplementoOrigem, " & vbCrLf & _
                " ORIGEM.Bairro AS BairroOrigem, " & vbCrLf & _
                " ORIGEM.Cep AS CepOrigem, " & vbCrLf & _
                " ORIGEM.Cidade AS CidadeOrigem, " & vbCrLf & _
                " ORIGEM.Estado AS EstadoOrigem, " & vbCrLf & _
                " T.Empresa as EmpresaOrigem, " & vbCrLf & _
                " isnull(T.EndEmpresa, 0) as EnderecoOrigem, " & vbCrLf & _
                " Origem.Nome as EmpresaOrigemNome, " & vbCrLf & _
                " T.CliFor AS Destinatario, " & vbCrLf & _
                " ISNULL(T.EnderecoCliFor, 0) AS EnderecoDestino, " & vbCrLf & _
                " T.EmpresaRecPag AS EmpresaPagadora, " & vbCrLf & _
                " ISNULL(T.EndEmpresaRecPag, 0) AS EnderecoDebito, " & vbCrLf & _
                " T.ContaContabilRecPag AS ContaContabilPagadora, " & vbCrLf & _
                " T.CliFor AS ClienteDebito, " & vbCrLf & _
                " ISNULL(T.EnderecoCliFor, 0) AS EndClienteDebito, " & vbCrLf & _
                " ISNULL(T.EmpresaRecPag, '') AS EmpresaCredito, " & vbCrLf & _
                " ISNULL(T.EndEmpresaRecPag, 0) AS EnderecoCredito, " & vbCrLf & _
                " ISNULL(T.ContaContabilCliFor,'') as ContaCredito, " & vbCrLf & _
                " ISNULL(T.CliFor,'') as ClienteCredito, " & vbCrLf & _
                " ISNULL(T.EnderecoCliFor ,0) as EndClienteCredito, " & vbCrLf & _
                " 2 AS ReceberPagar, " & vbCrLf & _
                " CASE WHEN T.Cheque is null or T.Cheque = 1 " & vbCrLf & _
                " THEN 'FALSE' ELSE 'TRUE' END Cheque, " & vbCrLf & _
                " CASE WHEN T.Slips IS NULL OR T.Slips = 1 " & vbCrLf & _
                " THEN 'FALSE' ELSE 'TRUE' END Slips, " & vbCrLf & _
                " CASE WHEN T.Recibo IS NULL OR T.Recibo = 1 " & vbCrLf & _
                " THEN 'FALSE' ELSE 'TRUE' END Recibo, " & vbCrLf & _
                "'FALSE' AS Aviso, " & vbCrLf & _
                " CASE WHEN T.Recibo IS NULL OR T.Recibo = 1 " & vbCrLf & _
                " THEN 'FALSE' ELSE 'TRUE' END ReciboDeposito, " & vbCrLf & _
                " isnull(T.Empresa, '') as EmpresaPedido, " & vbCrLf & _
                " isnull(T.EndEmpresa, -1) as EndEmpresaPedido, " & vbCrLf & _
                " isnull(T.Pedido,'') as Pedido, " & vbCrLf & _
                " isnull(T.Pedido,'') as PedidoFixacao, " & vbCrLf & _
                " CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf & _
                "	(SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C' )  " & vbCrLf & _
                " ELSE " & vbCrLf & _
                "	(SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D') " & vbCrLf & _
                " END AS OficialProgramado, " & vbCrLf & _
                " CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf & _
                "	(SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C') " & vbCrLf & _
                " ELSE " & vbCrLf & _
                "	(SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D') " & vbCrLf & _
                " END AS MoedaProgramado, " & vbCrLf & _
                " CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf & _
                "	(SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C' AND Conta_Id = T.ContaContabilRecPag) " & vbCrLf & _
                " ELSE " & vbCrLf & _
                "	(SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D' AND Conta_Id = T.ContaContabilCliFor) " & vbCrLf & _
                " END AS OficialBaixado, " & vbCrLf & _
                " CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf & _
                "	(SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C' AND Conta_Id = T.ContaContabilRecPag) " & vbCrLf & _
                " ELSE " & vbCrLf & _
                "	(SELECT SUM(ValorMoeda) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D' AND Conta_Id = T.ContaContabilCliFor) " & vbCrLf & _
                " END AS MoedaBaixado, " & vbCrLf & _
                " 0 AS Acrescimos, " & vbCrLf & _
                " 0 AS Descontos, " & vbCrLf & _
                " 0 AS Juros, " & vbCrLf & _
                " 0 AS Deducoes, " & vbCrLf & _
                " CASE WHEN T.RecPag = 'P' THEN  " & vbCrLf & _
                "	(SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'C' AND Conta_Id = T.ContaContabilRecPag) " & vbCrLf & _
                " ELSE " & vbCrLf & _
                "	(SELECT SUM(ValorOficial) FROM TitulosxContaContabil WHERE Titulo_Id = T.Titulo_Id AND DC_Id = 'D' AND Conta_Id = T.ContaContabilCliFor) " & vbCrLf & _
                " END AS ValorParaPagamento, " & vbCrLf & _
                " T.Historico, " & vbCrLf & _
                " CASE WHEN LEN(T.CliFor) = 14 " & vbCrLf & _
                " THEN 2 ELSE 1 END TipoInscricaoFornecedor, " & vbCrLf & _
                " Destinatario.Reduzido as ReduzidoFornecedor, " & vbCrLf & _
                " T.CliFor as Fornecedor, " & vbCrLf & _
                " T.EnderecoCliFor as EndFornecedor, " & vbCrLf & _
                " Destinatario.Nome as NomeFornecedor, " & vbCrLf & _
                " Destinatario.Cidade as CidadeFornecedor, " & vbCrLf & _
                " Destinatario.Estado as EstadoFornecedor, " & vbCrLf & _
                " Destinatario.Numero as NumeroFornecedor, " & vbCrLf & _
                " Destinatario.Complemento as ComplementoFornecedor, " & vbCrLf & _
                " Destinatario.CEP as CEPFornecedor, " & vbCrLf & _
                " Destinatario.Bairro as BairroFornecedor, " & vbCrLf & _
                " Destinatario.Endereco as EnderecoFornecedor, " & vbCrLf & _
                " T.BancoCliFor AS BancoCliente, " & vbCrLf & _
                " T.AgenciaCliFor AS AgenciaCliente, " & vbCrLf & _
                " T.DigitoAgenciaCliFor AS DigitoAgenciaCliente, " & vbCrLf & _
                " T.ContaCliFor AS ContaCliente, " & vbCrLf & _
                " T.DigitoContaCliFor AS DigitoContaCliente,  " & vbCrLf & _
                " 'C' AS TipoContaCliente, " & vbCrLf & _
                " 'Banco:' + str(T.BancoCliFor)+ ' AG:' + T.AgenciaCliFor + ' ' + T.DigitoAgenciaCliFor + ' CONTA:' + T.ContaCliFor + ' ' + T.DigitoContaCliFor + ' ' + 'C' AS DadosBancariosFornc,  " & vbCrLf & _
                " ISNULL(BC.Banco_Id,0) AS Banco, " & vbCrLf & _
                " B.Descricao AS BancoDesc, " & vbCrLf & _
                " isnull(BC.Agencia_Id, 0) AS Agencia, " & vbCrLf & _
                " isnull(BC.DigitoAgencia_Id,'') AS DigitoAgencia, " & vbCrLf & _
                " isnull(BC.Conta_Id,'') AS ContaCorrente, " & vbCrLf & _
                " isnull(BC.DigitoConta_Id,'') AS DigitoConta, " & vbCrLf & _
                " '' AS Destinacao, " & vbCrLf & _
                " 'BANCO:' + str(BC.Banco_Id) + ' AG:' + BC.Agencia_Id + ' ' + BC.DigitoAgencia_Id + ' CONTA:' + BC.Conta_Id + ' ' + BC.DigitoConta_Id AS DadosBancariosPagadora, " & vbCrLf & _
                " T.Moeda, " & vbCrLf & _
                " M.Descricao as MoedaDesc, " & vbCrLf & _
                " ISNULL(Clientes.Reduzido,'') as ReduzidoDebito, " & vbCrLf & _
                " TTB.ModalidadeDePagamento, " & vbCrLf & _
                " TTB.NossoNumero, " & vbCrLf & _
                " TTB.DigitoNossoNumero, " & vbCrLf & _
                " isnull(TTB.NumeroDaRemessa,'') as NumeroDaRemessa, " & vbCrLf & _
                " isnull(T.CodigoDeBarra,'') as CodigoDeBarras, " & vbCrLf & _
                " isnull(LEFT(T.CodigoDeBarra,3),0) as BancoNoCodigoDeBarras, " & vbCrLf & _
                " CASE WHEN ISNULL(T.CodigoDeBarraDigitado, 0)= 1 THEN 'TRUE' ELSE 'FALSE' END AS CodigoDigitado " & vbCrLf & _
                " FROM Titulos AS T " & vbCrLf & _
                " INNER JOIN Clientes AS ORIGEM " & vbCrLf & _
                " ON T.Empresa = ORIGEM.Cliente_Id  " & vbCrLf & _
                " AND T.EndEmpresa = ORIGEM.Endereco_Id " & vbCrLf & _
                " INNER JOIN Clientes AS Destinatario " & vbCrLf & _
                " ON T.CliFor = Destinatario.Cliente_Id " & vbCrLf & _
                " AND T.EnderecoCliFor = Destinatario.Endereco_Id " & vbCrLf & _
                " INNER JOIN Clientes AS Clientes " & vbCrLf & _
                " ON T.EmpresaRecPag = Clientes.Cliente_Id " & vbCrLf & _
                " AND T.EndEmpresaRecPag = Clientes.Endereco_Id " & vbCrLf & _
                " LEFT JOIN TitulosXTransacaoBancaria AS TTB " & vbCrLf & _
                " ON T.Titulo_Id = TTB.Titulo_Id " & vbCrLf & _
                " LEFT JOIN Provisoes as P " & vbCrLf & _
                " ON P.Provisao_Id = T.Provisao " & vbCrLf & _
                " LEFT JOIN Carteira as C " & vbCrLf & _
                " ON C.Carteira_id = ISNULL(T.CarteiraDoTitulo,0) " & vbCrLf & _
                " LEFT JOIN Indexadores as I " & vbCrLf & _
                " ON I.Indexador_id = T.Indexador " & vbCrLf & _
                " LEFT JOIN TiposDePagamentos AS TP " & vbCrLf & _
                " ON TP.TipoDePagamento_id = T.TipoPagto " & vbCrLf & _
                " LEFT JOIN Situacoes as S " & vbCrLf & _
                " ON S.Situacao_id = T.Situacao " & vbCrLf & _
                " LEFT JOIN BancosXContas AS BC " & vbCrLf & _
                " ON BC.ContaContabil = T.ContaContabilRecPag " & vbCrLf & _
                " LEFT JOIN Bancos AS B " & vbCrLf & _
                " ON B.Banco_Id = BC.Banco_Id " & vbCrLf & _
                " LEFT JOIN Moedas as M " & vbCrLf & _
                " ON M.Moeda_id = T.Moeda " & vbCrLf & _
                " WHERE " & vbCrLf & _
                " (T.RegistroMestre = T.Titulo_Id OR T.RegistroMestre IS NULL OR T.RegistroMestre = '') " & vbCrLf

        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            strSQL &= " AND T.EmpresaRecPag = '" & strEmpresa(0) & "' " & vbCrLf
            strSQL &= " AND T.EndEmpresaRecPag = " & strEmpresa(1) & " " & vbCrLf
            Parametros &= "Pagador: " & ddlEmpresa.Text
        End If

        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            Dim strCliente As String() = txtCodigoCliente.Value.Split(";")
            strSQL &= " AND T.CliFor = '" & strCliente(0) & "' " & vbCrLf
            strSQL &= " AND T.EnderecoCliFor = " & strCliente(1) & " " & vbCrLf
            Parametros &= "Pagador: " & txtNomeCliente.Text
        End If

        If Not String.IsNullOrWhiteSpace(ddlCarteira.SelectedValue) Then
            strSQL &= "   AND T.CarteiraDoTitulo = '" & ddlCarteira.SelectedValue & "'" & vbCrLf
            'Parametros &= "Carteira: " & ddlCarteira.Text
        End If

        If ddlSituacao.SelectedIndex <> 0 Then
            If ddlSituacao.SelectedIndex = 1 Then
                strSQL &= "  AND T.Situacao = 1 " & vbCrLf
                'Parametros &= "Situação: Aceitos e Rejeitados pelo Banco"
            Else
                Dim strSituacao As String() = ddlSituacao.SelectedValue.Split("-")
                strSQL &= " AND T.Situacao = " & strSituacao(0) & " " & vbCrLf
                'Parametros &= "Situação: " & CType(CB_Situacao.SelectedItem, BLL.Situacao).Situacao_Id.ToString & " - " & CType(CB_Situacao.SelectedItem, BLL.Situacao).Descricao & MicrosofCP.VisualBasic.Chr(13)
            End If
        End If

        If rdEmissao.Checked Then
            strSQL &= "AND T.Baixa " & vbCrLf
            'Parametros &= "Periodo baseado na Data de Movimento "
        Else
            strSQL &= "AND T.Reprogramacao " & vbCrLf
            'Parametros &= "Periodo baseado na Data de Vencimento "
        End If

        strSQL &= " BETWEEN '" & txtDataInicial.Text.ToSqlDate() & "' AND '" & txtDataFinal.Text.ToSqlDate() & "' " & vbCrLf

        Return strSQL
    End Function

#End Region

    Protected Sub btnRelatorio_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            'Dim arquivo As New BoletoNet.ArquivoRemessa(BoletoNet.TipoArquivo.CNAB400)
            Dim strEmpresa As String() = ddlEmpresa.SelectedValue.Split("-")
            If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não há registros carregados")
                Exit Sub
            End If
            Relatorio(CType(Session("objTitulos"), DataSet).Tables(0).DataSet, txtDataInicial.Text & " à " & txtDataFinal.Text, strEmpresa(0), strEmpresa(1))
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Public Sub Relatorio(ByVal Registros As DataSet, ByVal Periodo As String, ByVal CNPJ As String, ByVal Endereco As String)
        Dim Ds_TitulosEnviados As New DataSet
        Dim dvTitulosEnviados As New DataView
        dvTitulosEnviados = Registros.Tables(0).DefaultView
        dvTitulosEnviados.RowFilter = "Enviar = TRUE"
        Ds_TitulosEnviados = Registros


        Dim crptRelatorio As New ReportDocument()
        crptRelatorio.FileName = Server.MapPath("~/Reports/Cr_TitulosEnviadosRemessaBancaria.rpt")
        crptRelatorio.Load(crptRelatorio.FileName, CrystalDecisions.Shared.OpenReportMethod.OpenReportByDefault)

        Dim NomeArquivo2 As String = "Files/" & Funcoes.GeraNomeArquivo & ".PDF"
        Dim NomeArquivo As String = Server.MapPath(NomeArquivo2)
        Dim arquivo As String = NomeArquivo

        crptRelatorio.SetDataSource(dvTitulosEnviados)

        Try
            Dim crparametervalues As ParameterValues
            Dim crparameterdiscretevalue As ParameterDiscreteValue
            Dim crparameterfielddefinitions As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinitions
            Dim crparameterfielddefinition As CrystalDecisions.CrystalReports.Engine.ParameterFieldDefinition

            crparameterfielddefinitions = crptRelatorio.DataDefinition.ParameterFields()

            crparameterfielddefinition = crparameterfielddefinitions.Item("Titulo")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Titulos Enviados na Remessa"
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Nome")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssNomeEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Cidade")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New ParameterDiscreteValue
            crparameterdiscretevalue.Value = HttpContext.Current.Session("ssCidadeEmpresa") & " - " & HttpContext.Current.Session("ssEstadoEmpresa")
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("Periodo")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Período de " & Format(CDate(txtDataInicial.Text), "dd/MM/yyyy") & " Á " & Format(CDate(txtDataFinal.Text), "dd/MM/yyyy") & ""
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            crparameterfielddefinition = crparameterfielddefinitions.Item("DataHora")
            crparametervalues = crparameterfielddefinition.CurrentValues
            crparameterdiscretevalue = New CrystalDecisions.Shared.ParameterDiscreteValue
            crparameterdiscretevalue.Value = "Emissão:"
            crparametervalues.Add(crparameterdiscretevalue)
            crparameterfielddefinition.ApplyCurrentValues(crparametervalues)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            crptRelatorio.ExportToDisk(ExportFormatType.PortableDocFormat, arquivo)

            If IO.File.Exists(arquivo) Then
                Funcoes.AbrirArquivo(Me.Page, NomeArquivo2)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, "Erro: " & Funcoes.EliminarCaracteresEspeciais(ex.Message.ToString))
        Finally
            crptRelatorio.Close()
            crptRelatorio.Dispose()
        End Try
    End Sub

    Protected Sub btnRemessa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Trim(txtNumValor.Text) <> "" And txtNumValor.Text > 0 Then 'tem saldo para remessa.
                If ddlCarteira.SelectedIndex > 0 Then
                    SomaTitulosRemessa()
                    'forca carga pagador
                    Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
                    Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))

                    Dim NomeArquivo As String
                    Dim strline As String
                    Dim strm As StreamWriter

                    Dim NomeArquivo2 As String = "RemessaBancaria/" & "PG" & Date.Now.Day.ToString("00") & Date.Now.Month.ToString("00") & Date.Now.Hour.ToString("00") & ".REM"
                    NomeArquivo = Server.MapPath(NomeArquivo2)

                    If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                    strm = New StreamWriter(NomeArquivo, True)

                    If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
                        MsgBox(Me.Page, "Não há registros carregados")
                        Exit Sub
                    End If

                    Try

                        Dim x As Integer = 1
                        Dim Valor As Double = 0
                        Dim Sql As String
                        Dim ListSql As New ArrayList


                        strline = HeaderPG()
                        strm.WriteLine(strline)

                        Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
                        Dim Objconta As New ContaTituloPgto(strEmpresaPagadora(0), strEmpresaPagadora(1), strConta(1), strConta(2), strConta(3), strConta(4), strConta(5), SequenciaDePagamentoRecuperado, SequenciaDeRegistroDePagamentoRecuperado)

                        For Each row As DataRow In CType(Session("objTitulos"), DataSet).Tables(0).Rows
                            'Se estiver Marcado pra ser enviado

                            If row.Item("Enviar").ToString.ToUpper = "TRUE" AndAlso row("TAG") <> "MT-PE" AndAlso row("TAG") <> "PE" Then 'AndAlso row("TAG") <> "RM" AndAlso row("TAG") <> "RM-MT"
                                'Verifica se a forma de pagamento pode ser enviada ao banco
                                If row("TipoPagto") = 1 Or row("TipoPagto") = 2 Then
                                    MsgBox(Me.Page, "Pagamentos em espécie / cheques não são enviados ao banco verifique o titulo: " & row("Titulo_Id").ToString)
                                    row("Enviar") = "FALSE"
                                Else
                                    'Verifica se os bancos são iguais caso a forma de pagamento seja transferencia entre contas
                                    If strConta(0) <> row("Banco") And row("TipoPagto") = 3 Then
                                        If (row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")) > 4999 Then '2999
                                            row("TipoPagto") = 6 'ted
                                        Else
                                            row("TipoPagto") = 11 'doc
                                        End If
                                    End If


                                    'If strConta(0) = row("Banco") And (row("TipoPagto") = 5 Or row("TipoPagto") = 6) Then   'Real Time tranferencia entre contas tem tarifa mais cara, respeita oq tiver no contas a pagar.
                                    '    row("TipoPagto") = 4
                                    'End If

                                    x += 1
                                    strline = Transacao(row, x, Objconta)
                                    strm.WriteLine(strline)
                                    Valor += CDbl((row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")))

                                    Sql = "Update ContasAPagar set "
                                    Sql &= "    DataEnvio       ='" & Date.Now.ToSqlDate() & "'"
                                    'altera a data de prorrogação e baixa para a mesma data de Movimento.
                                    If (Not CDate(row("prorrogacao")).Equals(CDate(row("Movimento")))) Then
                                        Sql &= "   ,Prorrogacao          ='" & CDate(row("Movimento")).ToSqlDate() & "'"
                                        Sql &= "   ,Baixa           ='" & CDate(row("Movimento")).ToSqlDate() & "'"
                                    Else
                                        Sql &= "   ,Baixa           ='" & CDate(row("prorrogacao")).ToSqlDate() & "'"
                                    End If

                                    Sql &= "   ,ValorLiquido  = isnull(ValorDoDocumento,0) -  isnull(Descontos,0) - isnull(Deducoes,0) + isnull(Juros,0) + isnull(Acrescimos,0)"
                                    Sql &= "   ,EmpresaPagadora   ='" & Pagador.CPFCNPJ.ToString & "'"
                                    'Sql &= "   ,EndEmpresaPagadora = " & Pagador.EndCnpj.ToString
                                    'Sql &= "   ,EmpresaCredito  ='" & Pagador.CPFCNPJ.ToString & "'"
                                    'Sql &= "   ,EnderecoCredito = " & Pagador.EndCnpj.ToString
                                    'Sql &= "   ,ContaContabilCliente    = " & Pagador.ContaContabil.ToString
                                    Sql &= "   ,SituacaoRemessaBancaria     =  case when RegistroMestre is null or RegistroMestre = '' then 201  when Registro_Id = RegistroMestre then 201 else 205 end"
                                    'Sql &= "   ,Situacao_Id     = 201"
                                    Sql &= "   ,NumeroDaRemessa = " & SequenciaDePagamentoRecuperado
                                    Sql &= "   ,CarteiraDoTitulo     = '" & ddlCarteira.SelectedValue.ToString & "'"
                                    Sql &= "   ,TipoPagto    = " & row("TipoPagto")
                                    Sql &= "   ,BancoPagador    = " & strConta(0)
                                    Sql &= "   ,AgenciaPagadora    = " & Objconta.Agencia
                                    Sql &= "   ,DigitoAgenciaPagadora    = " & Objconta.DigitoAgencia
                                    Sql &= "   ,ContaPagadora    = " & Objconta.Conta_Id
                                    Sql &= "   ,DigitoContaPagadora    = " & Objconta.DigitoConta_Id
                                    Sql &= "   ,BancoCliente    = " & row("BancoCliente")
                                    Sql &= "   ,AgenciaCliente    = '" & row("AgenciaCliente").ToString.Trim & "'"
                                    Sql &= "   ,DigitoAgenciaCliente    = '" & row("DigitoAgenciaCliente").ToString.Trim & "'"
                                    Sql &= "   ,ContaCliente    = '" & row("ContaCliente").ToString.Trim & "'"
                                    Sql &= "   ,DigitoContaCliente    = '" & row("DigitoContaCliente").ToString.Trim & "'"
                                    Sql &= "   ,TipoContaCliente    = '" & row("TipoContaCliente").ToString.Trim & "'"
                                    Sql &= " Where (Registro_Id = '" & row("Titulo_id") & "' or RegistroMestre = '" & row("Titulo_id") & "')"
                                    ListSql.Add(Sql)
                                End If
                            Else
                                row("Enviar") = "FALSE"
                            End If
                        Next
                        strline = Trailler(x + 1, Valor)
                        strm.WriteLine(strline)
                        strm.Close()
                        Sql = " Update ParametrosDaCarteiraDeCobranca set"
                        Sql &= "  SequenciaDePagamentos          = " & SequenciaDePagamentoRecuperado + 1.ToString
                        Sql &= "  ,SequenciaDeRegistroDePagamento = " & SequenciaDeRegistroDePagamentoRecuperado + x.ToString
                        Sql &= " Where Empresa_Id         ='" & Pagador.CPFCNPJ & "'"
                        'Sql &= "   and EndEmpresa_id      = " & Pagador.EndCnpj
                        Sql &= "   and Agencia_id         = " & Objconta.Agencia.ToString
                        Sql &= "   and DigitoDaAgencia_id ='" & Objconta.DigitoAgencia & "'"
                        Sql &= "   and Conta_id           = " & Objconta.Conta_Id
                        Sql &= "   and DigitoConta_Id     ='" & Objconta.DigitoConta_Id & "'"
                        ListSql.Add(Sql)

                        If x = 1 Then
                            strm.Close()
                            MsgBox(Me.Page, "Nenhum registro gerado para o arquivo de Pagamento")
                        Else
                            If Banco.GravaBanco(ListSql) Then
                                MsgBox(Me.Page, "Arquivo gerado com Sucesso, Nº da Remessa: " & SequenciaDePagamentoRecuperado.ToString, eTitulo.Sucess)
                                btnRelatorio_Click(Nothing, Nothing)
                                Session("NomeArquivoDownload") = NomeArquivo2
                                btnDownload.Visible = True
                            Else
                                strm.Close()
                                MsgBox(Me.Page, "Erro Durante a Atualizacao dos Titulos na Base de Dados")
                            End If
                        End If
                    Catch ex As Exception
                        If NomeArquivo2.Length > 0 Then
                            strm.Close()
                            Kill(NomeArquivo2)
                        End If
                        MsgBox(Me.Page, ex.ToString())
                    Finally
                        Consultar()
                        btnTotalTitulos_Click(Nothing, Nothing)
                    End Try
                Else
                    MsgBox(Me.Page, "Selecione Uma Carteira!")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

#Region "Arquivo Remessa Pagamento"

    Function HeaderPG() As String
        Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))
        Dim linha As String

        Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
        Dim objconta As New ContaTituloPgto(strEmpresaPagadora(0), strEmpresaPagadora(1), strConta(1), strConta(2), strConta(3), strConta(4), strConta(5))

        '001 a 001 - 001 - Identificacao do Registro - "0"
        linha = "0"
        '002 a 009 - 008 - Identificacao da Empresa no Banco
        linha &= FitStringLength(Pagador.Codigo.ToString(), 8, 8, "0", 0, True, True, True)
        '010 a 010 - 001 -Tipo de Incricao da Empresa Pagadora - 1 = CPF, 2 = CNPJ, 3 = Outros
        linha &= "2"
        '011 a 019 CNPJ/CPF - 009
        '020 a 023 Filial   - 004
        '024 a 025 Controle - 002
        linha &= FitStringLength(strEmpresaPagadora(0), 15, 15, "0", 0, True, True, True)
        '026 a 065 Nome da Empresa Pagadora - 040
        'linha &= FitStringLength(Pagador.NomeDaEmpresaNoBanco.ToString(), 40, 40, " ", 0, True, True, False)
        '066 a 067 - 002 - Tipo De Seviço 20 Pagto Fornecedor
        linha &= "20"
        '068 a 068 - 001 - Origem do Arquivo,  1 = Cliente 2 = Banco
        linha &= "1"
        '069 a 073 - 005 - Numero sequencial da remessa
        linha &= FitStringLength(objconta.SequenciaDePagamento.ToString, 5, 5, "0", 0, True, True, True)
        SequenciaDePagamentoRecuperado = objconta.SequenciaDePagamento
        SequenciaDeRegistroDePagamentoRecuperado = objconta.SequenciaDeRegistroDePagamento
        '074 a 078 - 005 - Numero do retorno 
        linha &= New String("0", 5)
        '079 a 086 - 008 - Data Gravacao Arquivo AAAAMMDD
        linha &= Date.Now.ToString("yyyyMMdd")
        '087 a 092 - 006 - Hora da Gravacao do Arquivo HHMMSS
        linha &= Format(Date.Now, "HHmmss")
        '093 a 097 - 005 Desidade de Gravação no Arquivo
        linha &= New String(" ", 5)
        '098 a 100 - 003 Unidade de Desidade de Gravação no Arquivo
        linha &= New String(" ", 3)
        '101 a 105 - 005 Identificacao Módulo Micro
        linha &= New String(" ", 5)
        '106 a 106 - 001 Tipo de Processamento
        linha &= "0"
        '107 a 180 - 074 Para Uso da Empresa
        linha &= New String(" ", 74)
        '181 a 260 - 080 Brancos
        linha &= New String(" ", 80)
        '261 a 477 - 217 Brancos
        linha &= New String(" ", 217)
        '478 a 486 - 009 Numero da Lista de Debito
        linha &= New String(" ", 9)
        '487 a 494 - 008 Brancos
        linha &= New String(" ", 8)
        '495 a 500 - 006 Numero sequencial do Registro
        linha &= "000001"
        Return linha
    End Function

    Function Transacao(ByVal row As DataRow, ByVal Numero As Integer, Objconta As ContaTituloPgto) As String
        Dim Tp As Integer
        Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), CBool(row("CodigoDigitado")))
        Select Case row("TipoPagto")
            Case 4 : Tp = 31
            Case 3 : Tp = 1
            Case 11 : Tp = 3
            Case 6 : Tp = 8
            Case 5 : Tp = 31
            Case 12 : Tp = 31
            Case Else
                MsgBox(Me.Page, "Erro inesperado, tipo de Pagamento não aceito... 264 a 265")
        End Select
        Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
        'Dim conta As New Conta(strEmpresaPagadora(0), strEmpresaPagadora(1), strConta(1), strConta(2), strConta(3), strConta(4), SequenciaDePagamentoRecuperado, SequenciaDeRegistroDePagamentoRecuperado)

        Dim NumRegCob As Integer = SequenciaDeRegistroDePagamentoRecuperado + Numero - 1
        Dim linha As String

        '001 a 001 - 001 - Identificacao Fixo "1"
        linha = "1"
        '002 a 002 - 001 - Tipo de Inscricao do Fornecedor 1 = CPF, 2 = CNPJ, 3 = Outros
        linha &= row("TipoInscricaoFornecedor").ToString()
        '003 a 011 - 009 - Base CNPJ Fornecedor 
        '012 a 015 - 004 - Filial
        '016 a 017 - 002 - Controle
        If row("TipoInscricaoFornecedor") = 2 Then
            linha &= FitStringLength(row("Fornecedor").ToString, 15, 15, "0", 0, True, True, True)
        Else
            linha &= FitStringLength(row("Fornecedor").ToString.Substring(0, row("Fornecedor").ToString.Length - 2), 9, 9, "0", 0, True, True, True)
            linha &= New String("0", 4)
            linha &= FitStringLength(row("Fornecedor").ToString.Substring(row("Fornecedor").ToString.Length - 2, 2), 2, 2, "0", 0, True, True, True)
        End If

        '018 a 047 - 030 - Nome Do Fornecedor
        linha &= FitStringLength(row("NomeFornecedor").ToString, 30, 30, " ", 0, True, True, False)
        '048 a 087 - 040 Endereço do Fornecedor 
        linha &= FitStringLength(row("EnderecoFornecedor").ToString, 40, 40, " ", 0, True, True, False)
        '088 a 092 - 005 - Numero do CEP do Fornecedor
        '093 a 095 - 003 - Complemento do CEP
        linha &= FitStringLength(row("CEPFornecedor").ToString, 8, 8, "0", 0, True, True, True)
        '096 a 098 - 003 - Codigo do Banco - Fornecedor
        '099 a 103 - 005 - Código da Agencia - Fornecedor
        '104 a 104 - 001 - Dígito da Agencia - Fornecedor
        '105 a 117 - 013 - Conta Corrente - Fornecedor
        '118 a 119 - 002 - Digito da Conta - Fornecedor
        If Tp = 31 AndAlso CInt(CodBarras.CodBanco) <> 237 Then
            'linha &= FitStringLength(row("BancoCliente").ToString, 3, 3, "0", 0, True, True, True)
            linha &= FitStringLength(CodBarras.CodBanco.ToString, 3, 3, "0", 0, True, True, True)
            linha &= New String("0", 5)
            linha &= " "
            linha &= New String("0", 13)
            linha &= New String(" ", 2)
            'If CodBarras.CampoLivre.Length <> 0 Then
            '    linha &= FitStringLength(CodBarras.CodBanco.ToString, 3, 3, "0", 0, True, True, True)
            '    linha &= FitStringLength(CodBarras.Agencia.ToString, 5, 5, "0", 0, True, True, True)
            '    linha &= FitStringLength(CodBarras.DigAgencia.ToString, 1, 1, "0", 0, True, True, True)
            '    linha &= FitStringLength(CodBarras.ContaCorrente.ToString, 13, 13, "0", 0, True, True, True)
            '    linha &= FitStringLength(CodBarras.DigContaCorrente.ToString, 2, 2, " ", 0, True, True, False)

            'End If
        Else

            If Tp = 31 And CodBarras.CampoLivre.Length <> 0 Then
                linha &= FitStringLength(CodBarras.CodBanco.ToString, 3, 3, "0", 0, True, True, True)
                linha &= FitStringLength(CodBarras.Agencia.ToString, 5, 5, "0", 0, True, True, True)
                linha &= FitStringLength(CodBarras.DigAgencia.ToString, 1, 1, "0", 0, True, True, True)
                linha &= FitStringLength(CodBarras.ContaCorrente.ToString, 13, 13, "0", 0, True, True, True)
                linha &= FitStringLength(CodBarras.DigContaCorrente.ToString, 2, 2, " ", 0, True, True, False)
            Else
                linha &= FitStringLength(row("BancoCliente").ToString, 3, 3, "0", 0, True, True, True)
                linha &= FitStringLength(row("AgenciaCliente").ToString, 5, 5, "0", 0, True, True, True)
                linha &= FitStringLength(row("DigitoAgenciaCliente").ToString, 1, 1, "0", 0, True, True, False)
                linha &= FitStringLength(row("ContaCliente").ToString, 13, 13, "0", 0, True, True, True)
                linha &= FitStringLength(row("DigitoContaCliente").ToString, 2, 2, " ", 0, True, True, False)
            End If

        End If

        '120 a 135 - 016 - Numero do Pagamento
        linha &= FitStringLength(NumRegCob.ToString, 16, 16, "0", 0, True, True, True)
        '136 a 138 - 003 - Carteira
        Select Case Tp.ToString + row("BancoCliente")
            Case 30237 : linha &= "030"
            Case 31237 : linha &= "031"
            Case Else : linha &= "000"
        End Select
        '139 a 150 - 012 - Nosso Numero
        Select Case Tp.ToString + row("BancoCliente")
            Case 31237 : linha &= FitStringLength(row("NossoNumero").ToString + row("DigitoNossoNumero").ToString, 12, 12, "0", 0, True, True, True)
            Case Else : linha &= New String("0", 12)
        End Select
        '151 a 165 - 015 Seu Numero
        linha &= New String(" ", 15)
        '166 a 173 - 008 Data De Vencimento AAAAMMDD
        linha &= CDate(row("Movimento")).ToString("yyyyMMdd")
        '174 a 181 - 008 Data De Emissao do Documento AAAAMMDD
        linha &= CDate(row("Prorrogacao")).ToString("yyyyMMdd")

        '182 a 189 - 008 Data Limite para Desconto
        linha &= New String("0", 8)
        'If row("Descontos") = 0 And row("Deducoes") = 0 Then
        '    linha &= New String("0", 8)
        'Else
        '    linha &= Format(Date.Now, "yyyyMMdd")
        'End If
        '190 a 190 - 001 - Fixos Zeros
        linha &= "0"
        '191 a 194 - 004 - Fator de Vencimento 
        linha &= New String("0", 4)

        '195 a 204 - 010 - Valor do Documento
        If row("TipoPagto") = 4 Or row("TipoPagto") = 5 Or row("TipoPagto") = 12 Then
            If CodBarras.EnviarVlrDocumento Then
                'linha &= FitStringLength(CInt(row("OficialProgramado") * 100).ToString, 10, 10, "0", 0, True, True, True)
                linha &= FitStringLength(CInt((row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")) * 100).ToString, 10, 10, "0", 0, True, True, True)
            Else
                linha &= New String("0", 10)
            End If
        Else
            'linha &= FitStringLength(CInt(row("OficialProgramado") * 100).ToString, 10, 10, "0", 0, True, True, True)
            linha &= FitStringLength(CInt((row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")) * 100).ToString, 10, 10, "0", 0, True, True, True)
        End If

        '205 a 219 - 015 - Valor do Pagamento
        linha &= FitStringLength(CInt((row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")) * 100).ToString, 15, 15, "0", 0, True, True, True)
        '220 a 234 - 015 - Valor do Desconto
        'linha &= FitStringLength(CInt((row("Deducoes") + row("Descontos")) * 100).ToString, 15, 15, "0", 0, True, True, True)
        linha &= FitStringLength("0", 15, 15, "0", 0, True, True, True)
        '235 a 249 - Valor do Acrescimo
        linha &= FitStringLength("0", 15, 15, "0", 0, True, True, True)
        '250 a 251 - 002 - Tipo do Documento
        '01 nf/fatura (numero nota obrigatorio)
        '02 fatura (numero nota obrigatorio)
        '03 nota fiscal (numero nota obrigatorio)
        '04 duplicata
        '05 outros
        linha &= "05"

        '252 a 261 - 010 - Numero Nota Fiscal / Fatura Duplicada
        'linha &= New String("0", 10)
        linha &= FitStringLength(row("Titulo_Id"), 10, 10, "0", 0, True, True, True)

        '262 a 263 - 002 - Serie Documento
        linha &= New String(" ", 2)

        '264 a 265 - 002 Modalidade de Pagamento
        Select Case row("TipoPagto")
            Case 4
                linha &= "31"
                Tp = 31
            Case 3 : linha &= "01"
                Tp = 1
            Case 11 : linha &= "03"
                Tp = 3
            Case 6 : linha &= "08"
                Tp = 8
            Case 5 : linha &= "31"
                Tp = 31
            Case 12 : linha &= "31"
                Tp = 31
            Case Else
                MsgBox(Me.Page, "Erro inesperado, tipo de Pagamento não aceito... 264 a 265")
                linha &= "99" 'ERRO
        End Select

        '266 a 273 - 008 Data Para Efetivacao do Pagamento
        'linha &= New String("0", 8) Format(row("Vencimento"), "yyyyMMdd")
        linha &= CDate(row("Movimento")).ToString("yyyyMMdd")
        '274 a 276 - 003 - Moeda
        linha &= New String(" ", 3)
        '277 a 278 - 002 - Situacao do Agendamento
        linha &= "01"
        '279 a 288 - Fixo Branco - utilizado somente no retorno
        linha &= New String(" ", 10)
        '289 a 289 - Tipo Movimento - 0 Inclusao, 5 Alteração, 9 exclusao
        Select Case row("TAG")
            Case "  " : linha &= "0"
            Case "MT" : linha &= "0"
            Case "DEL" : linha &= "9"
            Case Else : linha &= "0"
        End Select
        '290 a 291 - 002 Codigo do movimento 00 - Agenda 25 - desautoriza agendamento
        linha &= "00"
        '292 a 295 - 004 Horario consulta saldo
        linha &= New String(" ", 4)
        '296 a 310 - 015 - Saldo Disponivel no momento da consulta - retorno
        linha &= New String(" ", 15)
        '311 a 325 - 015 - Valor da taxa Pré funding - retorno
        linha &= New String(" ", 15)
        '326 a 331 - 006 Reserva
        linha &= New String(" ", 6)
        '332 a 371 - 040 - Dados Complementares 
        linha &= New String(" ", 40)
        '372 a 372 - 001 - Fixo Branco - Reserva
        linha &= " "
        '373 a 373 - 001 - Fixo Branco - Retorno
        linha &= " "
        '374 a 413 - 040 
        Select Case Tp
            Case 1
                If row("TipoContaCliente").ToString <> "P" Then
                    linha &= New String(" ", 40)
                    '252 a 261 - 010 - Numero Nota Fiscal / Fatura Duplicada
                    linha = linha.Substring(0, 251) & FitStringLength(row("Titulo_Id"), 10, 10, "0", 0, True, True, True) & linha.Substring(261, linha.Length - 261)
                Else
                    If Pagador.CPFCNPJ.Substring(0, 8) = row("Fornecedor").ToString.Substring(0, 8) Then
                        linha &= "D" 'Mesma Titularidade
                        linha &= New String("0", 6) 'Nr do DOC/TED - Deverá ser informado com Zeros
                        'linha &= "01" 'Codigo da Finalidade 01 = Credito em Conta Corrente
                        linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "11", "01"), 2, 2, " ", 0, True, True, False) 'Finalidade
                        linha &= New String(" ", 29)
                        linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "02", "01"), 2, 2, " ", 0, True, True, False) 'Tipo Conta
                    Else
                        linha &= "C" 'Titularidade Diferente
                        linha &= New String("0", 6) 'Nr do DOC/TED - Deverá ser informado com Zeros
                        'linha &= "01" 'Codigo da Finalidade 01 = Credito em Conta Corrente
                        linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "11", "01"), 2, 2, " ", 0, True, True, False) 'Finalidade
                        linha &= New String(" ", 29)
                        linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "02", "01"), 2, 2, " ", 0, True, True, False) 'Tipo Conta
                    End If
                End If
            Case 5 : linha &= New String(" ", 40)
            Case 3, 8
                If Pagador.CPFCNPJ.Substring(0, 8) = row("Fornecedor").ToString.Substring(0, 8) Then
                    linha &= "D" 'Mesma Titularidade
                    linha &= New String("0", 6) 'Nr do DOC/TED - Deverá ser informado com Zeros
                    'linha &= "01" 'Codigo da Finalidade 01 = Credito em Conta Corrente
                    linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "11", "01"), 2, 2, " ", 0, True, True, False) 'Finalidade
                    linha &= New String(" ", 29)
                    linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "02", "01"), 2, 2, " ", 0, True, True, False) 'Tipo Conta
                Else
                    linha &= "C" 'Titularidade Diferente
                    linha &= New String("0", 6) 'Nr do DOC/TED - Deverá ser informado com Zeros
                    'linha &= "01" 'Codigo da Finalidade 01 = Credito em Conta Corrente
                    linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "11", "01"), 2, 2, " ", 0, True, True, False) 'Finalidade
                    linha &= New String(" ", 29)
                    linha &= FitStringLength(IIf(row("TipoContaCliente").ToString = "P", "02", "01"), 2, 2, " ", 0, True, True, False) 'Tipo Conta
                End If

            Case 31
                If CodBarras.CampoLivre.Length <> 0 Then
                    linha &= CodBarras.CampoLivre
                    linha &= CodBarras.DigitoVerificador
                    linha &= CodBarras.CodMoeda
                    linha &= New String(" ", 13)
                    linha = linha.Substring(0, 190) & FitStringLength(CodBarras.FatorDeVencimento, 4, 4, "0", 0, True, True, True) & linha.Substring(194, linha.Length - 194)
                Else
                    linha &= New String(" ", 40)
                End If
            Case Else
                linha &= New String("9", 40)

        End Select

        '414 a 415 - 002 - Codigo de area na empresa
        linha &= "00"
        '416 a 450 - 035 - Campos para uso da Empresa
        linha &= New String(" ", 35)
        '451 a 472 - 022 - Fixo Brancos
        linha &= New String(" ", 22)
        '473 a 477 - 005 - Codigo de Lancamento
        linha &= New String("0", 5)
        '478 a 478 - 001 - Reserva Branco
        linha &= " "
        '479 a 479 - 001 - Tipo de conta do Fornecedor
        If Objconta.TipoConta = "C" Then
            linha &= "1"
        Else
            linha &= "2"
        End If
        '480 a 486 - 007 - Conta Complementar
        'linha &= FitStringLength(Pagador.ContaBancaria.Conta, 7, 7, "0", 0, True, True, True)
        Dim strContaPagador As String() = ddlReceber.SelectedValue.Split(";")
        'Dim objconta As New ContaTituloPgto(strEmpresaPagadora(0), strEmpresaPagadora(1), strContaPagador(1), strContaPagador(2), strContaPagador(3), strContaPagador(4), strContaPagador(5))
        'linha &= FitStringLength(IIf(objconta.Conta_Id = "40155", "0040156", IIf(objconta.Conta_Id = "40156", "0040155", "0000000")), 7, 7, "0", 0, True, True, True) 'fixo conta complementar
        linha &= FitStringLength(Objconta.Conta_Id, 7, 7, "0", 0, True, True, True)
        '487 a 494 - 008 - Reserva
        linha &= New String(" ", 8)
        '495 a 500 - 006 Numero sequencial do registro
        linha &= FitStringLength(Numero.ToString(), 6, 6, "0", 0, True, True, True)
        Return linha
    End Function

    Function Trailler(ByVal Numero As Integer, ByVal Valor As Double) As String
        Dim linha As String
        '001 a 001 - 001 - Identificaçao do registro 9
        linha = "9"
        '002 a 007 - 006 - Total de registros no Arquivo
        linha &= FitStringLength(Numero.ToString(), 6, 6, "0", 0, True, True, True)
        '008 a 024 - 017 - Somatorio do conteudo do valor de pagamento
        linha &= FitStringLength(Convert.ToString(Valor * 100), 17, 17, "0", 0, True, True, True)
        '025 a 494 - 470 Brancos
        linha &= New String(" ", 470)
        '495 a 500 - 006 - Sequencial Crescente no Arquivo
        linha &= FitStringLength(Numero.ToString(), 6, 6, "0", 0, True, True, True)
        Return linha
    End Function

    Public Shared Function FitStringLength(ByVal SringToBeFit As String, ByVal maxLength As Integer, ByVal minLength As Integer, _
                                           ByVal FitChar As Char, ByVal maxStartPosition As Integer, ByVal maxTest As Boolean, _
                                           ByVal minTest As Boolean, ByVal isNumber As Boolean) As String
        Try
            Dim result As String = ""

            If maxTest = True Then
                ' max
                If SringToBeFit.Length > maxLength Then
                    result += SringToBeFit.Substring(maxStartPosition, maxLength)
                End If
            End If

            If minTest = True Then
                ' min
                If SringToBeFit.Length <= minLength Then
                    If isNumber = True Then
                        result += DirectCast((New String(FitChar, (minLength - SringToBeFit.Length)) + SringToBeFit), String)
                    Else
                        result += DirectCast((SringToBeFit + New String(FitChar, (minLength - SringToBeFit.Length))), String)
                    End If
                End If
            End If
            Return result
        Catch ex As Exception
            Dim tmpEx As New Exception("Problemas ao Formatar a string. String = " & SringToBeFit, ex)
            Throw tmpEx
        End Try
    End Function

#End Region

#Region "Safra"

    Function HeaderSafra(ByVal sequencia As Integer) As String
        Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))
        Dim linha As String

        Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
        Dim objconta As New ContaTituloPgto(strEmpresaPagadora(0), strEmpresaPagadora(1), "0035", "0", "24551", "7", "C")

        '001 a 001 - 001 - Identificacao Registro Header - "0"
        linha = "0"
        '002 a 002 - 002 - Identificacao arquivo remessa - "1"
        linha &= "1"
        '003 a 009 - 003 - Identificacao arquivo por extenso "REMESSA"
        linha &= "REMESSA"
        '010 a 011 - 004 - Código de Identificação do serviço "11"
        linha &= "11"
        '012 a 026 - 005 - Identificação do serviço "PAGTOS FORNECED"
        linha &= "PAGTOS FORNECED"
        '027 a 034 - 006 - Identificação da Empresa CONTA "9999999"
        linha &= "00245517"
        '035 a 035 - 007 - Validação do trailer ("S" ou "N")"
        linha &= "S"
        '036 a 037 - 008 - Campo sem preenchimento (Brancos)"
        linha &= New String(" ", 2)
        '038 a 044 - 009 - Código da agência do cliente: AG 0099999"
        linha &= "0003500"
        '045 a 046 - 010 - Campo sem preenchimento (Brancos)"
        linha &= New String(" ", 2)

        '047 a 076 - 011 - Nome do cliente por extenso"
        Dim nomedaempresa = "CURTUME PANORAMA LTDA"
        linha &= FitStringLength(nomedaempresa, 30, 30, " ", 0, True, True, False)

        '077 a 079 - 012 - Código do Banco ( FIXO "422" )"
        linha &= "422"
        '080 a 094 - 013 - Nome do banco por extenso ( "BANCO SAFRA S/A" )"
        linha &= "BANCO SAFRA S/A"
        '095 a 100 - 014 - Data da gravação do arquivo"
        linha &= Now.ToString("ddMMyy")
        '101 a 101 - 015 - Orgiem do arquivo (preenchimento do banco)"
        linha &= New String(" ", 1)
        '102 a 103 - 016 - Código do terceiro (preenchimento do banco)"
        linha &= New String(" ", 2)
        '104 a 387 - 017 - Campo sem preenchimento (Brancos)"
        linha &= New String(" ", 284)
        '388 a 388 - 018 - Valida CNPJ/CPF no arquivo detalhe ( "S" ou "N" )"
        linha &= "S"
        '389 a 394 - 019 - Número sequencial do arquivo remessa"
        linha &= FitStringLength(1, 6, 6, "0", 0, True, True, True)
        '395 a 400 - 020 - Número sequencial do registro no arquivo ( "000001" )"
        linha &= FitStringLength(sequencia, 6, 6, "0", 0, True, True, True)

        '069 a 073 - 005 - Numero sequencial da remessa
        'linha &= FitStringLength(objconta.SequenciaDePagamento.ToString, 5, 5, "0", 0, True, True, True)
        'SequenciaDePagamentoRecuperado = objconta.SequenciaDePagamento
        'SequenciaDeRegistroDePagamentoRecuperado = objconta.SequenciaDeRegistroDePagamento

        Return linha
    End Function

    Function BoletoSafra(ByVal row As DataRowView, ByVal sequencia As Integer) As String
        Dim linha As String = String.Empty

        'Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), True)
        Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), CBool(row("CodigoDigitado")))

        '001 a 001 - 001 Identificação registro compromisso ( FIXO "1" )
        linha = "1"
        '002 a 003 - 002 Tipo de inscrição da empresa (CNPJ "01"; CPF "02")
        linha &= "01"
        '004 a 017 - 003 Número de inscrição da Empresa
        linha &= FitStringLength(row("EmpresaOrigem"), 14, 14, "0", 0, True, True, True)
        '018 a 025 - 004 Identificação Empresa: CONTA 09999999
        ' linha &= FitStringLength(row("ContaCorrente"), 8, 8, "0", 0, True, True, True)
        linha &= "00245517"
        '026 a 028 - 005 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 3)
        '029 a 035 - 006 Código da agência do cliente: AG 0099999
        'linha &= FitStringLength(row("Agencia"), 7, 7, "0", 0, True, True, True)
        linha &= "0003500"
        '036 a 037 - 007 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 2)
        '038 a 062 - 008 Uso exclusivo da Empresa
        linha &= New String(" ", 25)
        '063 a 076 - 009 Código de inscrição do fornecedor
        linha &= FitStringLength(row("Fornecedor"), 14, 14, "0", 0, True, True, True)


        '077 a 079 - 010 Tipo de Documento - Deve ser preenchido de acordo com o pagamento que será feito e de acordo com as ocorrências disponíveis no layout (DUP- duplicata, NF- nota fiscal, REC- recibo,  BLQ- bloqueto, OUT- outros, NP- nota promissória)
        linha &= "BLQ"
        '080 a 089 - 011 Número de compromisso no banco ( Deve ser preenchido com o número do documento que será pago )
        linha &= FitStringLength(row("Titulo_Id"), 10, 10, "0", 0, True, True, True)
        '090 a 090 - 012 Sequência de compromisso ( Deve ser preenchido caso existam parcelas no pagamento )
        linha &= New String(" ", 1)


        '091 a 107 - 013 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 17)

        '108 a 108 - 014 Código da Operação ( FIXO "C" )
        linha &= "C"


        '109 a 110 - 015 Código da Ocorrência (Deve ser preenchido de acordo com as ocorrências disponíveis no layout: 01 - inclusão, 02 - alteração, 03 - exclusão, 04 - autorização, 05 - bloqueio ou 06 - desbloqueio)
        linha &= "01"

        '111 a 120 - 016 Número do compromisso para cliente - Deve ser preenchido com o número do documento que será pago
        linha &= FitStringLength(row("Titulo_Id"), 10, 10, "0", 0, True, True, True)


        '121 a 126 - 017 Data de vencimento"
        linha &= Format(row("Vencimento"), "ddMMyy")
        '127 a 139 - 018 Valor para pagamento"
        linha &= FitStringLength(row("OficialProgramado").ToString.Replace(",", ""), 13, 13, "0", 0, True, True, True)
        '140 a 142 - 019 Tipo de pagamento ( FIXO "COB" )
        linha &= "COB"
        '143 a 145 - 020 Cód. do Banco Destino
        linha &= FitStringLength(row("BancoNoCodigoDeBarras"), 3, 3, "0", 0, True, True, True)
        '146 a 152 - 021 Agência Destino (Opcional)
        linha &= New String("0", 7)
        '153 a 155 - 022 Zeros
        linha &= New String("0", 3)
        '156 a 165 - 023 Conta Corrente Destino (Opciona)
        linha &= New String("0", 10)
        '166 a 185 - 024 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 20)


        '186 a 192 - 025 Agência para pagamento
        'linha &= "0024551"
        linha &= "0003500"


        '193 a 208 - 026 Nosso número (opcional)
        linha &= New String(" ", 16)
        '209 a 247 - 027 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 39)


        '248 a 250 - 028 Banco portador bloquete
        linha &= FitStringLength(CodBarras.CodBanco, 3, 3, "0", 0, True, True, True)

        '251 a 263 - 029 Valor abatimento
        linha &= FitStringLength((row("Descontos") + row("Deducoes")).ToString.Replace(",", ""), 13, 13, "0", 0, True, True, True)

        '264 a 293 - 030 Nome fornecedor
        linha &= FitStringLength(row("NomeFornecedor"), 30, 30, " ", 0, True, True, False)

        '294 a 303 - 031 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 10)

        Dim codigo = FitStringLength(row("CodigoDeBarras"), 44, 44, "0", 0, True, True, True)

        '304 a 347 - 032 Código de barras
        If CBool(row("CodigoDigitado")) Then
            codigo = (CodBarras.CodBanco + CodBarras.CodMoeda + CodBarras.DigitoVerificador + CodBarras.FatorDeVencimento + FitStringLength(row("ValorParaPagamento").ToString.Replace(",", ""), 10, 10, "0", 0, True, True, True) + CodBarras.CampoLivre)
        End If

        linha &= codigo

        '348 a 360 - 033 Valor de juros mora / multa
        linha &= FitStringLength((row("Acrescimos") + row("Juros")).ToString.Replace(",", ""), 13, 13, "0", 0, True, True, True)
        '361 a 366 - 034 Data pagamento"
        linha &= Format(row("Baixa"), "ddMMyy")
        '367 a 379 - 035 Valor autorizado para pagamento"
        linha &= FitStringLength(row("ValorParaPagamento").ToString.Replace(",", ""), 13, 13, "0", 0, True, True, True)
        '380 a 383 - 036 Código da moeda
        linha &= "REAL"
        '384 a 387 - 037 Tipo de carteira (opcional)
        linha &= New String(" ", 4)
        '388 a 390 - 038 Espécie de documento (opcional)
        linha &= New String(" ", 3)
        '391 a 394 - 039 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 4)
        '395 a 400 - 040 Número sequencialdo Registro (ultimo número +1)
        linha &= FitStringLength(sequencia, 6, 6, "0", 0, True, True, True)

        Return linha
    End Function

    Function TrailerSafra(ByVal sequencia As Integer, ByVal valorTotal As Decimal, ByVal valorDesconto As Decimal, ByVal valorValorAcrescimo As Decimal, ByVal valorPagamento As Decimal) As String
        Dim linha As String = String.Empty
        '001 a 001 - 001 Identificação do registro Trailler ( FIXO "9" )
        linha = "9"
        '002 a 124 - 002 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 123)
        '125 a 139 - 003 Somatório do campo valor
        linha &= FitStringLength(valorTotal.ToString.Replace(",", ""), 15, 15, "0", 0, True, True, True)
        '140 a 248 - 004 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 109)
        '249 a 263 - 005 Somatório do campo Abatimento
        linha &= FitStringLength(valorDesconto.ToString.Replace(",", ""), 15, 15, "0", 0, True, True, True)
        '264 a 345 - 006 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 82)
        '346 a 360 - 007 Somatório do Campo Juros de mora/Multa
        linha &= FitStringLength(valorValorAcrescimo.ToString.Replace(",", ""), 15, 15, "0", 0, True, True, True)
        '361 a 364 - 008 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 4)
        '365 a 379 - 009 Somatório do Campo Valor Autorizado para Pagamento
        linha &= FitStringLength(valorPagamento.ToString.Replace(",", ""), 15, 15, "0", 0, True, True, True)
        '380 a 394 - 010 Campo sem preenchimento (Brancos)
        linha &= New String(" ", 15)
        '395 a 400 - 011 Número Sequencial do Registro
        linha &= FitStringLength(sequencia, 6, 6, "0", 0, True, True, True)
        Return linha
    End Function

#End Region
    Public Function BuscaCedente(ByVal cliente_id As String, ByVal endereco As String) As BoletoNet.Cedente
        'Dim Conta As Conta = CType(CB_Conta.SelectedItem, Conta)
        Dim sql As String
        sql = "  SELECT Clientes.Cliente_Id,"
        sql &= "       Clientes.Endereco_Id,"
        sql &= "       Clientes.Nome,"
        sql &= "       ParametrosDaCarteiraDeCobranca.Agencia_Id,"
        sql &= "       ParametrosDaCarteiraDeCobranca.DigitoDaAgencia_Id,"
        sql &= "       ParametrosDaCarteiraDeCobranca.Conta_Id,"
        sql &= "       ParametrosDaCarteiraDeCobranca.DigitoConta_Id,"
        sql &= "       ParametrosDaCarteiraDeCobranca.SubConta_Id,"
        sql &= "       ParametrosDaCarteiraDeCobranca.Instrucao01,"
        sql &= "       ParametrosDaCarteiraDeCobranca.Instrucao02,"
        sql &= "       ParametrosDaCarteiraDeCobranca.MensagemDeInstrucao,"
        sql &= "       isnull(ParametrosDaCarteiraDeCobranca.DiasParaProtesto,0) as DiasParaProtesto,"
        sql &= "       isnull(ParametrosDaCarteiraDeCobranca.DiasBaixaDecursoPrazo,0) as DiasBaixaDecursoPrazo,"
        sql &= "       ParametrosDaCarteiraDeCobranca.JurosMes,"
        sql &= "       ParametrosDaCarteiraDeCobranca.MoraDiaria,"
        sql &= "       ParametrosDaCarteiraDeCobranca.SequenciaDeRemessa,"
        sql &= "       ParametrosDaCarteiraDeCobranca.CodigoDaEmpresaPagamento as CodigoDaEmpresaNoBanco,"
        sql &= "       ParametrosDaCarteiraDeCobranca.ContaContabil,"
        sql &= "       ParametrosDaCarteiraDeCobranca.NomeDaEmpresaNoBanco"
        sql &= "  FROM Clientes"
        sql &= " INNER JOIN ParametrosDaCarteiraDeCobranca"
        sql &= "    ON Clientes.Cliente_Id  = ParametrosDaCarteiraDeCobranca.Empresa_Id "
        sql &= "   AND Clientes.Endereco_Id = ParametrosDaCarteiraDeCobranca.EndEmpresa_Id"
        sql &= " Where Clientes.Cliente_id  = '" & cliente_id & "'"
        sql &= "   and Clientes.Endereco_id =  " & endereco
        'sql &= "   and ParametrosDaCarteiraDeCobranca.Agencia_id         = " & Conta.Agencia
        'sql &= "   and ParametrosDaCarteiraDeCobranca.DigitoDaAgencia_Id ='" & Conta.DigitoAgencia & "'"
        'sql &= "   and ParametrosDaCarteiraDeCobranca.Conta_Id           = " & Conta.Conta_Id
        'sql &= "   and ParametrosDaCarteiraDeCobranca.DigitoConta_Id     ='" & Conta.DigitoConta_Id & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Cedente")
        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            Return Nothing
        Else
            For Each dr As DataRow In ds.Tables(0).Rows
                Dim c As New BoletoNet.Cedente(dr("Cliente_Id"), dr("Nome"), CInt(dr("Agencia_Id")).ToString("0000"), dr("DigitoDaAgencia_Id"), dr("Conta_Id"), dr("DigitoConta_Id"))
                c.Codigo = CInt(dr("CodigoDaEmpresaNoBanco"))
                'c.NrArqRemessa = CInt(dr("SequenciaDeRemessa"))
                'c.ContaContabil = CInt(dr("ContaContabil"))
                'c.NomeDaEmpresaNoBanco = dr("NomeDaEmpresaNoBanco")
                Return c
            Next
        End If
        Return Nothing
    End Function

    Protected Sub btnTotalTitulos_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            txtNumValor.Text = SomaTitulosRemessa().ToString
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ChkGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim chkTitulo As CheckBox = CType(sender, CheckBox)
            Dim row As GridViewRow = CType(chkTitulo.NamingContainer, GridViewRow)
            If chkTitulo.Checked Then
                CType(Session("objTitulos"), DataSet).Tables(0).Rows(row.RowIndex).Item("Enviar") = "TRUE"
            Else
                CType(Session("objTitulos"), DataSet).Tables(0).Rows(row.RowIndex).Item("Enviar") = "FALSE"
            End If
            SomaTitulosRemessa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub ChkTodosGridTitulos_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim chkTodos As CheckBox = CType(gridPagamentoDeTitulos.HeaderRow.FindControl("chkTodosGridTitulos"), CheckBox)

            Dim check As Boolean = False
            If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In CType(Session("objTitulos"), DataSet).Tables(0).Rows
                    If (chkTodos.Checked) Then
                        row("Enviar") = "TRUE"
                    Else
                        row("Enviar") = "FALSE"

                    End If
                Next
            End If
            gridPagamentoDeTitulos.DataSource = Session("objTitulos")
            gridPagamentoDeTitulos.DataBind()
            chkTodos = CType(gridPagamentoDeTitulos.HeaderRow.FindControl("chkTodosGridTitulos"), CheckBox)
            If (Convert.ToBoolean(Session("Check"))) Then
                chkTodos.Checked = True
                Session("Check") = False
            Else
                chkTodos.Checked = False
                Session("Check") = True
            End If
            SomaTitulosRemessa()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub btnCliente_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objClientePgtoTitulo" & HID.Value, "txtNome")
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub gridPagamentoDeTitulos_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count > 0 Then
                Dim cliente As New [Lib].Negocio.Cliente(CType(Session("objTitulos"), DataSet).Tables(0).Rows(gridPagamentoDeTitulos.SelectedIndex)("Fornecedor").ToString(), CType(Session("objTitulos"), DataSet).Tables(0).Rows(gridPagamentoDeTitulos.SelectedIndex)("EndFornecedor").ToString())
                ucConsultaDadosBancarios.CarregaGrid(cliente.Codigo, cliente.CodigoEndereco)
                Popup.ConsultaDeDadosBancarios(Me.Page, "objContaBancariaPagDeTitulos" & HID.Value)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Private Sub Download()
        Response.ContentType = "text/plain"
        Response.AppendHeader("Content-Disposition", "attachment; filename=" & Session("NomeArquivoDownload").ToString.Split("/").ToArray(1))
        Const bufferLength As Integer = 10000
        Dim buffer As Byte() = New Byte(bufferLength - 1) {}
        Dim length As Integer = 0
        Dim download As Stream = Nothing
        Try
            download = New FileStream(Server.MapPath("~/" & Session("NomeArquivoDownload")), FileMode.Open, FileAccess.Read)
            Do
                If Response.IsClientConnected Then
                    length = download.Read(buffer, 0, bufferLength)
                    Response.OutputStream.Write(buffer, 0, length)
                    buffer = New Byte(bufferLength - 1) {}
                Else
                    length = -1
                End If
            Loop While length > 0
            Response.Flush()
            Response.End()
        Finally
            If download IsNot Nothing Then
                download.Close()
            End If
        End Try
    End Sub

    Protected Sub GerarRemessa()
        If Trim(txtNumValor.Text) <> "" And txtNumValor.Text > 0 Then 'tem saldo para remessa.
            If ddlCarteira.SelectedIndex > 0 Then

                'forca carga pagador
                Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
                Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))

                Dim NomeArquivo As String
                Dim strline As String
                Dim strm As StreamWriter
                Dim x As Integer = 1
                Dim Valor As Double = 0

                Dim NomeArquivo2 As String = "RemessaBancaria/" & "PG" & Date.Now.Day.ToString("00") & Date.Now.Month.ToString("00") & Date.Now.Hour.ToString("00") & ".REM"
                NomeArquivo = Server.MapPath(NomeArquivo2)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                strm = New StreamWriter(NomeArquivo, True)

                If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Não há registros carregados")
                    Exit Sub
                End If

                Try


                    Dim Sql As String
                    Dim ListSql As New ArrayList


                    strline = HeaderPG()
                    strm.WriteLine(strline)

                    Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
                    Dim Objconta As New ContaTituloPgto(strEmpresaPagadora(0), strEmpresaPagadora(1), strConta(1), strConta(2), strConta(3), strConta(4), strConta(5), SequenciaDePagamentoRecuperado, SequenciaDeRegistroDePagamentoRecuperado)

                    For Each row As DataRow In CType(Session("objTitulos"), DataSet).Tables(0).Rows
                        'Se estiver Marcado pra ser enviado

                        If row.Item("Enviar").ToString.ToUpper = "TRUE" AndAlso row("TAG") <> "MT-PE" AndAlso row("TAG") <> "PE" Then 'AndAlso row("TAG") <> "RM" AndAlso row("TAG") <> "RM-MT"
                            'Verifica se a forma de pagamento pode ser enviada ao banco
                            If row("TipoPagto") = 1 Or row("TipoPagto") = 2 Then
                                MsgBox(Me.Page, "Pagamentos em espécie / cheques não são enviados ao banco verifique o titulo: " & row("Titulo_Id").ToString)
                                row("Enviar") = "FALSE"
                            Else
                                'Verifica se os bancos são iguais caso a forma de pagamento seja transferencia entre contas
                                If strConta(0) <> row("Banco") And row("TipoPagto") = 3 Then
                                    If (row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")) > 4999 Then '2999
                                        row("TipoPagto") = 6 'ted
                                    Else
                                        row("TipoPagto") = 11 'doc
                                    End If
                                End If


                                'If strConta(0) = row("Banco") And (row("TipoPagto") = 5 Or row("TipoPagto") = 6) Then   'Real Time tranferencia entre contas tem tarifa mais cara, respeita oq tiver no contas a pagar.
                                '    row("TipoPagto") = 4
                                'End If

                                x += 1
                                strline = Transacao(row, x, Objconta)
                                strm.WriteLine(strline)
                                Valor += CDbl((row("OficialProgramado") + row("Acrescimos") + row("Juros") - row("Deducoes") - row("Descontos")))

                                Sql = "Update ContasAPagar set "
                                Sql &= "    DataEnvio       ='" & Date.Now.ToSqlDate() & "'"
                                'altera a data de prorrogação e baixa para a mesma data de Movimento.
                                If (Not CDate(row("prorrogacao")).Equals(CDate(row("Movimento")))) Then
                                    Sql &= "   ,Prorrogacao          ='" & Format(row("Movimento"), "yyyy-MM-dd") & "'"
                                    Sql &= "   ,Baixa           ='" & CDate(row("Movimento")).ToSqlDate() & "'"
                                Else
                                    Sql &= "   ,Baixa           ='" & CDate(row("prorrogacao")).ToSqlDate() & "'"
                                End If

                                Sql &= "   ,ValorLiquido  = isnull(ValorDoDocumento,0) -  isnull(Descontos,0) - isnull(Deducoes,0) + isnull(Juros,0) + isnull(Acrescimos,0)"
                                Sql &= "   ,EmpresaPagadora   ='" & Pagador.CPFCNPJ.ToString & "'"
                                ' Sql &= "   ,EndEmpresaPagadora = " & Pagador.EndCnpj.ToString
                                'Sql &= "   ,EmpresaCredito  ='" & Pagador.CPFCNPJ.ToString & "'"
                                'Sql &= "   ,EnderecoCredito = " & Pagador.EndCnpj.ToString
                                'Sql &= "   ,ContaContabilCliente    = " & Pagador.ContaContabil.ToString
                                Sql &= "   ,SituacaoRemessaBancaria     =  case when RegistroMestre is null or RegistroMestre = '' then 201  when Registro_Id = RegistroMestre then 201 else 205 end"
                                'Sql &= "   ,Situacao_Id     = 201"
                                Sql &= "   ,NumeroDaRemessa = " & SequenciaDePagamentoRecuperado
                                Sql &= "   ,CarteiraDoTitulo     = '" & ddlCarteira.SelectedValue.ToString & "'"
                                Sql &= "   ,TipoPagto    = " & row("TipoPagto")
                                Sql &= "   ,BancoPagador    = " & strConta(0)
                                Sql &= "   ,AgenciaPagadora    = " & Objconta.Agencia
                                Sql &= "   ,DigitoAgenciaPagadora    = " & Objconta.DigitoAgencia
                                Sql &= "   ,ContaPagadora    = " & Objconta.Conta_Id
                                Sql &= "   ,DigitoContaPagadora    = " & Objconta.DigitoConta_Id
                                Sql &= "   ,BancoCliente    = " & row("BancoCliente")
                                Sql &= "   ,AgenciaCliente    = '" & row("AgenciaCliente").ToString.Trim & "'"
                                Sql &= "   ,DigitoAgenciaCliente    = '" & row("DigitoAgenciaCliente").ToString.Trim & "'"
                                Sql &= "   ,ContaCliente    = '" & row("ContaCliente").ToString.Trim & "'"
                                Sql &= "   ,DigitoContaCliente    = '" & row("DigitoContaCliente").ToString.Trim & "'"
                                Sql &= "   ,TipoContaCliente    = '" & row("TipoContaCliente").ToString.Trim & "'"
                                Sql &= " Where (Registro_Id = '" & row("Titulo_id") & "' or RegistroMestre = '" & row("Titulo_id") & "')"
                                ListSql.Add(Sql)
                            End If
                        Else
                            row("Enviar") = "FALSE"
                        End If
                    Next
                    strline = Trailler(x + 1, Valor)
                    strm.WriteLine(strline)
                    strm.Close()
                    Sql = " Update ParametrosDaCarteiraDeCobranca set"
                    Sql &= "  SequenciaDePagamentos          = " & SequenciaDePagamentoRecuperado + 1.ToString
                    Sql &= "  ,SequenciaDeRegistroDePagamento = " & SequenciaDeRegistroDePagamentoRecuperado + x.ToString
                    Sql &= " Where Empresa_Id         ='" & Pagador.CPFCNPJ & "'"
                    'Sql &= "   and EndEmpresa_id      = " & Pagador.EndCnpj
                    Sql &= "   and Agencia_id         = " & Objconta.Agencia.ToString
                    Sql &= "   and DigitoDaAgencia_id ='" & Objconta.DigitoAgencia & "'"
                    Sql &= "   and Conta_id           = " & Objconta.Conta_Id
                    Sql &= "   and DigitoConta_Id     ='" & Objconta.DigitoConta_Id & "'"
                    ListSql.Add(Sql)

                    If x = 1 Then
                        strm.Close()
                        MsgBox(Me.Page, "Nenhum registro gerado para o arquivo de Pagamento")
                    Else
                        If Banco.GravaBanco(ListSql) Then
                            MsgBox(Me.Page, "Arquivo gerado com Sucesso, Nº da Remessa: " & SequenciaDePagamentoRecuperado.ToString, eTitulo.Sucess)
                            btnRelatorio_Click(Nothing, Nothing)
                            Session("NomeArquivoDownload") = NomeArquivo2
                            btnDownload.Visible = True
                        Else
                            strm.Close()
                            MsgBox(Me.Page, "Erro Durante a Atualizacao dos Titulos na Base de Dados")
                        End If
                    End If
                Catch ex As Exception
                    If NomeArquivo2.Length > 0 Then
                        strm.Close()
                        Kill(NomeArquivo2)
                    End If
                    MsgBox(Me.Page, ex.ToString())
                Finally
                    Consultar()
                    btnTotalTitulos_Click(btnTotalTitulos, Nothing)
                    lblQtdeEValorTotal.Text = "Foram enviados " & (x - 1) & " Registro(s) para o arquivo de remessa no valor total de  " & String.Format("{0:N2}", Valor)
                End Try
            Else
                MsgBox(Me.Page, "Selecione Uma Carteira!")
            End If
        End If
    End Sub

#Region "Banco do Brasil"
    Function RemessaBrasil() As Boolean
        Dim linha As String = String.Empty
        Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim b As New BoletoNet.Banco(341)

        If ddlCarteira.SelectedIndex > 0 Then
            If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não há registros carregados")
                Exit Function
            End If

            Dim NomeArquivo As String
            Dim strline As String = String.Empty
            Dim strm As StreamWriter
            Dim qtdeRegistro As Integer = 0
            Dim quantidadeLote As Integer = 0
            Dim quantidadeRegistroLote As Integer = 0
            Dim valorTotal As Decimal = 0
            Dim sql As String = String.Empty
            Dim remessa As String = DateTime.Now.ToString("ddMMyyyyss")
            Dim ListSql As New ArrayList

            Dim NomeArquivo2 As String = "RemessaBancaria/" & "PG" & Date.Now.Day.ToString("00") & Date.Now.Month.ToString("00") & Date.Now.Minute.ToString("00") & ".REM"
            NomeArquivo = Server.MapPath(NomeArquivo2)

            If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

            strm = New StreamWriter(NomeArquivo, True)

            Try
                quantidadeRegistroLote = quantidadeRegistroLote + 1

                'Header Arquivo
                Dim rowHeaderArquivo As DataRowView = CType(Session("objTitulos"), DataSet).Tables(0).DefaultView.Item(0)
                strline = HeaderBrasil(rowHeaderArquivo, quantidadeRegistroLote)
                strm.WriteLine(strline)

                strm.Close()

                If (Banco.GravaBanco(ListSql)) Then
                    MsgBox(Me.Page, "Arquivo gerado com Sucesso, Nº da Remessa: " & remessa, eTitulo.Sucess)
                    btnRelatorio_Click(Nothing, Nothing)
                    Session("NomeArquivoDownload") = NomeArquivo2
                    btnDownload.Visible = True
                Else
                    MsgBox(Me.Page, "Erro durante a geração do arquivo de remessa!")
                End If
            Catch ex As Exception
                If NomeArquivo2.Length > 0 Then
                    strm.Close()
                    Kill(NomeArquivo2)
                End If
                MsgBox(Me.Page, ex.ToString())
            End Try
        End If

    End Function

    Function HeaderBrasil(ByVal row As DataRowView, ByVal sequencia As Integer) As String
        Dim linha As String = String.Empty

        '001 a 003 - código do banco na compensação (001) Banco do Brasil
        linha = "001"
        '004 a 007 - código do lote 'fixo (0000)'
        linha &= "0000"
        '008 a 008 - tipo de registro 'fixo (0)'
        linha &= "0"
        '009 a 017 - brancos
        linha &= New String(" ", 9)
        '018 a 018 - tipo empresa (1)CPF, (2) CNPJ
        linha &= "2"
        '019 a 032 - CNPJ
        linha &= FitStringLength(row("EmpresaOrigem"), 14, 14, "0", 0, True, True, True)
        '033 a 052 - brancos - código do convênio do banco
        linha &= New String(" ", 20)
        '053 a 057 - Agência
        linha &= FitStringLength(row("Agencia"), 5, 5, "0", 0, True, True, True)
        '058 a 058 - brancos - dígito verificador da agência
        linha &= New String(" ", 1)
        '059 a 070 - Conta
        linha &= FitStringLength(row("ContaCorrente"), 12, 12, "0", 0, True, True, True)
        '071 a 071 - brancos - dígito verificador da conta
        linha &= New String(" ", 1)
        '072 a 072 - DAC - Digito Verificador da Agência/Conta
        linha &= row("DigitoConta")
        '073 a 102 - Nome da Empresa
        linha &= FitStringLength(row("EmpresaOrigemNome"), 30, 30, " ", 0, True, True, True)
        '103 a 132 - Nome do Banco
        linha &= FitStringLength(row("BancoDesc"), 30, 30, " ", 0, True, True, True)
        '133 a 142 - brancos
        linha &= New String(" ", 10)
        '143 a 143 - Código do arquivo (1) Remessa, (2) Retorno
        linha &= "1"
        '144 a 151 - Data de geração do arquivo (DDMMAAAA)
        linha &= DateTime.Now.ToString("ddMMyyyy")
        '152 a 157 - Hora de Geração do Arquivo (HHMMSS)
        linha &= DateTime.Now.ToString("HHmmss")
        '158 a 163 - Número sequencial do registro no arquivo ( "000001" )"
        linha &= FitStringLength(sequencia, 6, 6, "0", 0, True, True, True)
        '164 a 166 - Nº da versão do layout do arquivo
        linha &= "103"
        '167 a 171 - Densidade de Gravação do Arquivo
        linha &= "00000"
        '172 a 191 - brancos - uso do banco
        linha &= New String(" ", 20)
        '192 a 211 - brancos - uso da empresa
        linha &= New String(" ", 20)
        '212 a 240 - brancos - uso febraban
        linha &= New String(" ", 29)

        Return linha
    End Function

    Function HeaderLoteBB(ByVal row As DataRowView, ByVal sequencia As Integer) As String
        Dim linha As String = String.Empty
        Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), True)

        '001 a 001 - Código do Banco Fixo(001 Banco do Brasil)
        linha = "001"
        '004 a 007 - Código do lote
        linha &= FitStringLength(sequencia, 4, 4, "0", 0, True, True, True)
        '008 a 008 - Tipo de Registro (1) header de lote
        linha &= "1"
        '009 a 009 - Tipo da Operação (C -Crédito)
        linha &= "C"
    End Function

#End Region

#Region "Safra"
    Function RemessaSafra() As Boolean
        Dim linha As String = String.Empty
        Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim b As New BoletoNet.Banco(341)

        If Trim(txtNumValor.Text) <> "" And txtNumValor.Text > 0 Then 'tem saldo para remessa.
            If ddlCarteira.SelectedIndex > 0 Then

                'forca carga pagador
                ' Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
                ' Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))

                If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
                    MsgBox(Me.Page, "Não há registros carregados")
                    Exit Function
                End If

                Dim NomeArquivo As String
                Dim strline As String = String.Empty
                Dim strm As StreamWriter
                'Dim qtdeRegistro As Integer = 0
                'Dim quantidadeLote As Integer = 0
                Dim quantidadeRegistroLote As Integer = 0
                Dim valorTotal As Decimal = 0
                Dim valorDesconto As Decimal = 0
                Dim valorValorAcrescimo As Decimal = 0
                Dim valorPagamento As Decimal = 0
                Dim sql As String = String.Empty
                Dim remessa As String = DateTime.Now.ToString("ddMMyyyyss")
                Dim ListSql As New ArrayList

                Dim NomeArquivo2 As String = "RemessaBancaria/" & "PG" & Date.Now.Day.ToString("00") & Date.Now.Month.ToString("00") & Date.Now.Minute.ToString("00") & ".REM"
                NomeArquivo = Server.MapPath(NomeArquivo2)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                strm = New StreamWriter(NomeArquivo, True)


                Try
                    quantidadeRegistroLote = quantidadeRegistroLote + 1

                    strline = HeaderSafra(quantidadeRegistroLote)

                    strm.WriteLine(strline)

                    Dim dvBoletoOutrosBancos As DataView = CType(Session("objTitulos"), DataSet).Tables(0).DefaultView

                    For Each row As DataRowView In dvBoletoOutrosBancos
                        If row.Item("Enviar").ToString.ToUpper = "TRUE" Then
                            quantidadeRegistroLote = quantidadeRegistroLote + 1
                            strline = BoletoSafra(row, quantidadeRegistroLote)
                            strm.WriteLine(strline)
                            valorTotal = valorTotal + row("OficialProgramado")
                            valorDesconto = valorDesconto + row("Descontos") + row("Deducoes")
                            valorValorAcrescimo = valorValorAcrescimo + row("Acrescimos") + row("Juros")
                            valorPagamento = valorPagamento + row("ValorParaPagamento")
                            sql = "UPDATE ContasAPagar SET NumeroDaRemessa = " & remessa & " Where (Registro_Id = '" & row("Titulo_id") & "' or RegistroMestre = '" & row("Titulo_id") & "')"
                            ListSql.Add(sql)
                        End If
                    Next

                    If (dvBoletoOutrosBancos.Count > 0) Then
                        'Trailler 
                        quantidadeRegistroLote = quantidadeRegistroLote + 1
                        strline = TrailerSafra(quantidadeRegistroLote, valorTotal, valorDesconto, valorValorAcrescimo, valorPagamento)
                        strm.WriteLine(strline)
                    End If

                    strm.Close()

                    If (Banco.GravaBanco(ListSql)) Then
                        MsgBox(Me.Page, "Arquivo gerado com Sucesso, Nº da Remessa: " & remessa, eTitulo.Sucess)
                        btnRelatorio_Click(Nothing, Nothing)
                        Session("NomeArquivoDownload") = NomeArquivo2
                        btnDownload.Visible = True
                    Else
                        MsgBox(Me.Page, "Erro durante a geração do arquivo de remessa!")
                    End If


                Catch ex As Exception
                    If NomeArquivo2.Length > 0 Then
                        strm.Close()
                        Kill(NomeArquivo2)
                    End If
                    MsgBox(Me.Page, ex.ToString())
                Finally
                    'lblQtdeEValorTotal.Text = "Foram enviados " & (x - 1) & " Registro(s) para o arquivo de remessa no valor total de  " & String.Format("{0:N2}", Valor)
                End Try
            Else
                MsgBox(Me.Page, "Selecione Uma Carteira!")
            End If
        End If
    End Function
#End Region

#Region "Itau"

    Function RemessaItau(ByVal codBan As Integer) As Boolean
        Dim linha As String = String.Empty
        Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        Dim b As New BoletoNet.Banco(codBan)

        If Not String.IsNullOrWhiteSpace(txtNumValor.Text) AndAlso CDec(txtNumValor.Text) > 0 Then 'tem saldo para remessa.
            If ddlCarteira.SelectedIndex > 0 Then
                'forca carga pagador
                'Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))

                Dim NomeArquivo As String
                Dim strline As String = String.Empty
                Dim strm As StreamWriter
                Dim qtdeRegistro As Integer = 0
                Dim quantidadeLote As Integer = 0
                Dim quantidadeRegistroLote As Integer = 0
                Dim valorTotal As Decimal = 0
                Dim sql As String = String.Empty
                Dim remessa As String = DateTime.Now.ToString("ddMMyyyyss")
                Dim ListSql As New ArrayList

                Dim NomeArquivo2 As String = "RemessaBancaria/" & "PG" & Date.Now.Day.ToString("00") & Date.Now.Month.ToString("00") & Date.Now.Minute.ToString("00") & ".REM"
                NomeArquivo = Server.MapPath(NomeArquivo2)

                If Dir(NomeArquivo).Length > 0 Then Kill(NomeArquivo)

                strm = New StreamWriter(NomeArquivo, True)

                Try
                    'Header Arquivo
                    Dim rowHeaderArquivo As DataRowView = CType(Session("objTitulos"), DataSet).Tables(0).DefaultView.Item(0)
                    strline = HeaderItau(rowHeaderArquivo, codBan)
                    strm.WriteLine(strline)

                    'Bloquetos do Itau
                    Dim dvBoletoItau As DataView = CType(Session("objTitulos"), DataSet).Tables(0).DefaultView

                    If codBan = 1 Then
                        dvBoletoItau.RowFilter = "BancoNoCodigoDeBarras = '001' AND Enviar = 'TRUE'"
                    Else
                        dvBoletoItau.RowFilter = "BancoNoCodigoDeBarras = '341' AND Enviar = 'TRUE'"
                    End If

                    For Each row As DataRowView In dvBoletoItau
                        If row.Item("Enviar").ToString.ToUpper = "TRUE" Then
                            quantidadeRegistroLote = quantidadeRegistroLote + 1
                            If (quantidadeRegistroLote = 1) Then
                                'Header Lote
                                quantidadeLote = quantidadeLote + 1
                                strline = HeaderSegJ(row, quantidadeLote, codBan)
                                strm.WriteLine(strline)
                            End If
                            'Detalhe J
                            strline = DetalheSegJ(row, quantidadeRegistroLote, quantidadeLote, codBan)
                            strm.WriteLine(strline)
                            qtdeRegistro = qtdeRegistro + 1

                            If codBan = 1 Or codBan = 341 Then
                                quantidadeRegistroLote = quantidadeRegistroLote + 1
                                strline = DetalheSegJ52(row, quantidadeRegistroLote, quantidadeLote, codBan)
                                strm.WriteLine(strline)
                                qtdeRegistro = qtdeRegistro + 1
                            End If

                            valorTotal = valorTotal + row("ValorParaPagamento")
                            'Adiciona ao update
                            sql = "UPDATE ContasAPagar SET NumeroDaRemessa = " & remessa & " Where (Registro_Id = '" & row("Titulo_id") & "' or RegistroMestre = '" & row("Titulo_id") & "')"
                            ListSql.Add(sql)
                        End If
                    Next
                    'Trailler Seg J
                    If (dvBoletoItau.Count > 0) Then
                        strline = TrailerSegJ((quantidadeRegistroLote + 2), valorTotal, quantidadeLote, codBan)
                        strm.WriteLine(strline)
                    End If
                    quantidadeRegistroLote = 0
                    valorTotal = 0
                    'Bloquetos Outros Bancos
                    Dim dvBoletoOutrosBancos As DataView = CType(Session("objTitulos"), DataSet).Tables(0).DefaultView

                    If codBan = 1 Then
                        dvBoletoOutrosBancos.RowFilter = "BancoNoCodigoDeBarras <> '001' AND Enviar = 'TRUE'"
                    Else
                        dvBoletoOutrosBancos.RowFilter = "BancoNoCodigoDeBarras <> '341' AND Enviar = 'TRUE'"
                    End If

                    For Each row As DataRowView In dvBoletoOutrosBancos
                        If row.Item("Enviar").ToString.ToUpper = "TRUE" Then
                            quantidadeRegistroLote = quantidadeRegistroLote + 1
                            If (quantidadeRegistroLote = 1) Then
                                'Header Lote
                                quantidadeLote = quantidadeLote + 1
                                strline = HeaderSegJ(row, quantidadeLote, codBan)
                                strm.WriteLine(strline)
                            End If
                            'Detalhe J
                            strline = DetalheSegJ(row, quantidadeRegistroLote, quantidadeLote, codBan)
                            strm.WriteLine(strline)
                            qtdeRegistro = qtdeRegistro + 1

                            If codBan = 1 Or codBan = 341 Then
                                quantidadeRegistroLote = quantidadeRegistroLote + 1
                                strline = DetalheSegJ52(row, quantidadeRegistroLote, quantidadeLote, codBan)
                                strm.WriteLine(strline)
                                qtdeRegistro = qtdeRegistro + 1
                            End If

                            'Adicona o Titulo no Update
                            valorTotal = valorTotal + row("ValorParaPagamento")
                            sql = "UPDATE ContasAPagar SET NumeroDaRemessa = " & remessa & " Where (Registro_Id = '" & row("Titulo_id") & "' or RegistroMestre = '" & row("Titulo_id") & "')"
                            ListSql.Add(sql)
                        End If
                    Next

                    If (dvBoletoOutrosBancos.Count > 0) Then
                        'Trailler Seg J
                        strline = TrailerSegJ((quantidadeRegistroLote + 2), valorTotal, quantidadeLote, codBan)
                        strm.WriteLine(strline)
                    End If

                    'Trailler Remessa
                    strline = TrailerItau(quantidadeLote, (quantidadeLote * 2) + qtdeRegistro + 2, codBan)
                    strm.WriteLine(strline)
                    strm.Close()

                    If (Banco.GravaBanco(ListSql)) Then
                        MsgBox(Me.Page, "Arquivo gerado com Sucesso, Nº da Remessa: " & remessa, eTitulo.Sucess)
                        btnRelatorio_Click(Nothing, Nothing)
                        Session("NomeArquivoDownload") = NomeArquivo2
                        btnDownload.Visible = True
                    Else
                        MsgBox(Me.Page, "Erro durante a geração do arquivo de remessa!")
                    End If

                Catch ex As Exception
                    If NomeArquivo2.Length > 0 Then
                        strm.Close()
                        Kill(NomeArquivo2)
                    End If
                    MsgBox(Me.Page, ex.ToString())
                Finally
                    Consultar()
                    btnTotalTitulos_Click(Nothing, Nothing)
                End Try

            End If
        End If
        Return False
    End Function

    Function HeaderItau(ByVal row As DataRowView, ByVal codBan As Integer) As String
        Dim linha As String = String.Empty
        'Dim strEmpresaPagadora As String() = ddlEmpresa.SelectedValue.Split("-")
        'Pagador = BuscaCedente(strEmpresaPagadora(0), strEmpresaPagadora(1))
        'Dim strConta As String() = ddlReceber.SelectedValue.Split(";")
        'Dim objconta As New ContaTituloPgto(strEmpresaPagadora(0), strEmpresaPagadora(1), strConta(1), strConta(2), strConta(3), strConta(4), strConta(5))

        '001 a 003 - código do banco na compensação (341) Itaú
        If codBan = 1 Then
            linha = "001"
        Else
            linha = "341"
        End If
        '004 a 007 - código do lote 'fixo (0000)'
        linha &= "0000"
        '008 a 008 - tipo de registro 'fixo (0)'
        linha &= "0"

        If codBan = 1 Then
            '009 a 017 - brancos
            linha &= New String(" ", 9)
        Else
            '009 a 014 - brancos
            linha &= New String(" ", 6)
            '015 a 017 - Tipo de Layout do arquivo 'fixo (080)'
            linha &= "080"
        End If

        '018 a 018 - tipo empresa (1)CPF, (2) CNPJ
        linha &= "2"
        '019 a 032 - CNPJ
        linha &= FitStringLength(row("EmpresaOrigem"), 14, 14, "0", 0, True, True, True)
        '033 a 052 - brancos
        If codBan = 1 Then
            linha &= FitStringLength("769840", 9, 9, "0", 0, True, True, True)
            linha &= "0126"
            linha &= New String(" ", 7)
        Else
            linha &= New String(" ", 20)
        End If
        '053 a 057 - Agência
        linha &= FitStringLength(row("Agencia"), 5, 5, "0", 0, True, True, True)
        '058 a 058 - brancos
        If codBan = 1 Then
            linha &= row("DigitoAgencia")
        Else
            linha &= New String(" ", 1)
        End If
        '059 a 070 - Conta
        linha &= FitStringLength(row("ContaCorrente"), 12, 12, "0", 0, True, True, True)
        '071 a 071 - brancos
        If codBan = 1 Then
            linha &= row("DigitoConta")
        Else
            linha &= New String(" ", 1)
        End If
        '072 a 072 - DAC
        If codBan = 1 Then
            linha &= New String(" ", 1)
        Else
            linha &= row("DigitoConta")
        End If
        '073 a 102 - Nome da Empresa
        linha &= FitStringLength(row("EmpresaOrigemNome"), 30, 30, " ", 0, True, True, False)
        '103 a 132 - Nome do Banco
        linha &= FitStringLength(row("BancoDesc"), 30, 30, " ", 0, True, True, False)
        '133 a 142 - brancos
        linha &= New String(" ", 10)
        '143 a 143 - Código do arquivo (1) Remessa, (2) Retorno
        linha &= "1"
        '144 a 151 - Data de geração do arquivo (DDMMAAAA)
        linha &= DateTime.Now.ToString("ddMMyyyy")
        '152 a 157 - Hora de Geração do Arquivo (HHMMSS)
        linha &= DateTime.Now.ToString("HHmmss")

        If codBan = 1 Then
            '158 a 163 - Zeros
            linha &= New String("0", 6)
            'nº da versão do layout do arquivo
            linha &= "103"
        Else
            '158 a 166 - Zeros
            linha &= New String("0", 9)
        End If

        '167 a 171 - Unidade de Densidade (00000) ???
        linha &= New String("0", 5)
        '172 a 240 - brancos
        linha &= New String(" ", 69)
        Return linha
    End Function

#Region "Segmento J (Bloquetos)"
    Function HeaderSegJ(ByVal row As DataRowView, ByVal sequencia As Integer, ByVal codBan As Integer) As String
        Dim linha As String = String.Empty
        Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), True)
        '001 a 001 - Código do Banco Fixo(341 Itaú)
        If codBan = 1 Then
            linha = "001"
        Else
            linha = "341"
        End If
        '004 a 007 - Código do lote
        linha &= FitStringLength(sequencia, 4, 4, "0", 0, True, True, True)
        '008 a 008 - Tipo de Registro (1) header de lote
        linha &= "1"
        '009 a 009 - Tipo da Operação (C -Crédito)
        linha &= "C"

        '010 a 011 - Tipo Pagamento (20) Fornecedores - VER BB SE USA 03 OU 20 MESMO  ??????????????????????????
        linha &= "20"

        '012 a 013 - Forma de pagamento
        If codBan = 1 Then
            linha &= IIf(CodBarras.CodBanco.Equals("001"), "30", "31")
        Else
            linha &= IIf(CodBarras.CodBanco.Equals("341"), "30", "31")
        End If

        '014 a 016 - Layout do Lote (030)
        If codBan = 1 Then
            linha &= "040"
        Else
            linha &= "030"
        End If
        '017 a 017 - Brancos
        linha &= New String(" ", 1)
        '018 a 018 - tipo empresa (1)CPF, (2) CNPJ
        linha &= "2"
        '019 a 032 - CNPJ
        linha &= FitStringLength(row("EmpresaOrigem"), 14, 14, "0", 0, True, True, True)
        '033 a 052 - Brancos
        If codBan = 1 Then
            linha &= FitStringLength("769840", 9, 9, "0", 0, True, True, True)
            linha &= "0126"
            linha &= New String(" ", 7)
        Else
            linha &= New String(" ", 20)
        End If
        '053 a 057 - Agência
        linha &= FitStringLength(row("Agencia"), 5, 5, "0", 0, True, True, True)
        '058 a 058 - Brancos
        If codBan = 1 Then
            linha &= row("DigitoAgencia")
        Else
            linha &= New String(" ", 1)
        End If
        '059 a 070 - Conta
        linha &= FitStringLength(row("ContaCorrente"), 12, 12, "0", 0, True, True, True)
        '071 a 071 - brancos
        If codBan = 1 Then
            linha &= row("DigitoConta")
        Else
            linha &= New String(" ", 1)
        End If
        '072 a 072 - DAC
        If codBan = 1 Then
            linha &= New String(" ", 1)
        Else
            linha &= row("DigitoConta")
        End If
        '073 a 102 - Nome Empresa
        linha &= FitStringLength(row("EmpresaOrigemNome"), 30, 30, "0", 0, True, True, True)
        '103 a 132 - Finalidade do Lote
        linha &= New String(" ", 30)
        '133 a 142 - Histórico da Conta
        linha &= New String(" ", 10)
        '143 a 172 - Endereço Empresa
        linha &= FitStringLength(row("EnderecoOrigem"), 30, 30, " ", 0, True, True, False)
        '173 a 177 - Número
        linha &= FitStringLength(row("NumeroOrigem"), 5, 5, "0", 0, True, True, True)
        '178 a 192 - Complemento
        linha &= FitStringLength(row("ComplementoOrigem"), 15, 15, " ", 0, True, True, False)
        '193 a 212 - Cidade
        linha &= FitStringLength(row("CidadeOrigem"), 20, 20, " ", 0, True, True, False)
        '213 a 220 - CEP
        linha &= FitStringLength(row("CepOrigem"), 8, 8, " ", 0, True, True, False)
        '221 a 222 - Estado
        linha &= row("EstadoOrigem")
        '223 a 230 - Brancos
        linha &= New String(" ", 8)
        '231 a 240 - Ocorrências do Retorno
        linha &= New String(" ", 10)
        Return linha
    End Function

    Function DetalheSegJ(ByVal row As DataRowView, ByVal sequencia As Integer, ByVal lote As Integer, ByVal codBan As Integer) As String
        Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), CBool(row("CodigoDigitado")))
        Dim linha As String = String.Empty
        '001 a 003 - Código do Banco (341) Itau
        If codBan = 1 Then
            linha = "001"
        Else
            linha = "341"
        End If
        '004 a 007 - Código do Lote
        linha &= FitStringLength(lote, 4, 4, "0", 0, True, True, True)
        '008 a 008 - Tipo de Registro (3) detalhe
        linha &= "3"
        '009 a 013 - Numero do Registro Sequencial
        linha &= FitStringLength(sequencia, 5, 5, "0", 0, True, True, True)
        '014 a 014 - Segmento J (J)
        linha &= "J"
        '015 a 017 - Tipo de Movimento
        linha &= "000"
        '018 a 020 - CÓD. DE BARRAS – CÓDIGO BANCO FAVORECIDO
        linha &= CodBarras.CodBanco
        '021 a 021 - CÓD. DE BARRAS – CÓDIGO DA MOEDA
        linha &= CodBarras.CodMoeda
        '022 a 022 - CÓD. DE BARRAS – DÍGITO VERIF. DO CÓD. BARRAS
        linha &= CodBarras.DigitoVerificador
        '023 a 026 - CÓD. DE BARRAS – FATOR DE VENCIMENTO
        linha &= CodBarras.FatorDeVencimento
        '027 a 036 - CÓD. DE BARRAS – VALOR
        linha &= FitStringLength(row("ValorParaPagamento").ToString.Replace(",", ""), 10, 10, "0", 0, True, True, True)
        '037 a 061 - CÓD. DE BARRAS - 'CAMPO LIVRE'
        linha &= FitStringLength(CodBarras.CampoLivre, 25, 25, "0", 0, True, True, True)
        '062 a 091 - Nome Favorecido
        linha &= FitStringLength(row("NomeFornecedor"), 30, 30, " ", 0, True, True, False)
        '092 a 099 - Data de Vencimento (DDMMAAAA)
        linha &= DateTime.Now.ToString("ddMMyyyy")
        '100 a 114 - Valor do Titulo
        linha &= FitStringLength(row("ValorParaPagamento").ToString.Replace(",", ""), 15, 15, "0", 0, True, True, True)
        '115 a 129 - Valor Desconto
        linha &= FitStringLength("0", 15, 15, "0", 0, True, True, True)
        '130 a 144 - Valor Acréscimos
        linha &= FitStringLength("0", 15, 15, "0", 0, True, True, True)
        '145 a 152 - Data do pagamento (DDMMAAA)
        linha &= CDate(row("Movimento")).ToString("ddMMyyyy")
        '153 a 167 - Valor Pagamento
        linha &= FitStringLength(row("ValorParaPagamento").ToString.Replace(",", ""), 15, 15, "0", 0, True, True, True)
        '168 a 182 - Zeros
        linha &= FitStringLength("0", 15, 15, "0", 0, True, True, True)
        '183 a 202 - Nº DOCTO ATRIBUÍDO PELA EMPRESA
        linha &= FitStringLength(row("Titulo_Id"), 20, 20, " ", 0, True, True, False)
        '203 a 215 - Brancos
        linha &= New String(" ", 13)
        '216 a 230 - NÚMERO ATRIBUÍDO PELO BANCO
        linha &= New String(" ", 15)
        '231 a 240 - Ocorrências no Retorno
        linha &= New String(" ", 10)
        Return linha
    End Function

    Function DetalheSegJ52(ByVal row As DataRowView, ByVal sequencia As Integer, ByVal lote As Integer, ByVal codBan As Integer) As String
        Dim CodBarras As New DadosBoleto(row("CodigoDeBarras"), CBool(row("CodigoDigitado")))
        Dim linha As String = String.Empty
        '001 a 003 - Código do Banco (341) Itau
        If codBan = 1 Then
            linha = "001"
        Else
            linha = "341"
        End If        '004 a 007 - Código do Lote
        linha &= FitStringLength(lote, 4, 4, "0", 0, True, True, True)
        '008 a 008 - Tipo de Registro (3) detalhe
        linha &= "3"
        '009 a 013 - Numero do Registro Sequencial
        linha &= FitStringLength(sequencia, 5, 5, "0", 0, True, True, True)
        '014 a 014 - Segmento J (J)
        linha &= "J"
        '015 a 017 - Tipo de Movimento   ?????
        linha &= "000"
        '018 a 019 - Identificação Registro
        linha &= "52"
        '020 - Tipo de Inscrição do Pagador 1- CPF 2 - CNPJ 
        linha &= "2"
        '021 a 035 - NÚMERO DE INSCRIÇÃO DO PAGADOR
        linha &= FitStringLength(row("EmpresaOrigem"), 15, 15, "0", 0, True, True, True)
        '036 a 075 - NOME DO PAGADOR   
        linha &= FitStringLength(row("EmpresaOrigemNome"), 40, 40, " ", 0, True, True, False)
        '076 - Tipo de Inscrição do Beneficiário 1- CPF 2 - CNPJ ????????
        If row("TipoInscricaoFornecedor") = "1" Then
            linha &= "1"
        Else
            linha &= "2"
        End If
        '077 a 091 - NÚMERO DE INSCRIÇÃO DO BENEFICIÁRIO
        linha &= FitStringLength(row("Fornecedor"), 15, 15, "0", 0, True, True, True)
        '092 a 131 - NOME DO BENEFICIARIO
        linha &= FitStringLength(row("NomeFornecedor"), 40, 40, " ", 0, True, True, False)

        '132 - Tipo de Inscrição do SACADOR AVALISTA 1- CPF 2 - CNPJ ????????
        linha &= "0"
        '133 a 147 - NÚMERO DE INSCRIÇÃO DO SACADOR AVALISTA  ?????
        linha &= "000000000000000"
        '148 a 187 - NOME DO SACADOR AVALISTA   ????
        linha &= New String(" ", 40)

        '188 a 240 - Complemento - BRANCOS
        linha &= New String(" ", 53)

        Return linha
    End Function

    Function TrailerSegJ(ByVal quantidade As Integer, ByVal valor As Decimal, ByVal lote As Integer, ByVal codBan As Integer) As String
        Dim linha As String = String.Empty
        '001 a 003 - Código do Banco (341) - Itaú
        If codBan = 1 Then
            linha = "001"
        Else
            linha = "341"
        End If
        '004 a 007 - Lote de Serviço
        linha &= FitStringLength(lote, 4, 4, "0", 0, True, True, True)
        '008 a 008 - Tipo de Registro (5)
        linha &= "5"
        '009 a 017 - Brancos
        linha &= New String(" ", 9)
        '018 a 023 - Qtde de registro no lote
        linha &= FitStringLength(quantidade, 6, 6, "0", 0, True, True, True)
        '024 a 041 - Total valor dos pagamentos
        linha &= FitStringLength(valor.ToString.Replace(",", ""), 18, 18, "0", 0, True, True, True)
        '042 a 059 - Zeros
        linha &= New String("0", 18)
        '060 a 230 - Brancos
        linha &= New String(" ", 171)
        '231 a 240 - Ocorrências no retorno
        linha &= New String(" ", 10)
        Return linha
    End Function
#End Region

    Function TrailerItau(ByVal quantidadeLote As Integer, ByVal quantidadeRegistro As Integer, ByVal codBan As Integer) As String
        Dim linha As String = String.Empty
        '001 a 003 - Código do Banco (341) - Itaú
        If codBan = 1 Then
            linha = "001"
        Else
            linha = "341"
        End If
        '004 a 007 - Código do Lote (9999)
        linha &= "9999"
        '008 a 008 - Tipo de Registro (9) Trailler do Arquivo
        linha &= "9"
        '009 a 017 - Brancos
        linha &= New String(" ", 9)
        '018 a 023 - QTDE LOTES DO ARQUIVO
        linha &= FitStringLength(quantidadeLote, 6, 6, "0", 0, True, True, True)
        '024 a 029 - QTDE REGISTROS DO ARQUIVO
        linha &= FitStringLength(quantidadeRegistro, 6, 6, "0", 0, True, True, True)

        If codBan = 1 Then
            '030 a 35 - Brancos
            linha &= New String("0", 6)
            '036 a 240 - Brancos
            linha &= New String(" ", 205)
        Else
            '030 a 240 - Brancos
            linha &= New String(" ", 211)
        End If

        Return linha
    End Function

#End Region

#Region "Barra Botões"

    Protected Sub lnkConsultar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConsultar.Click
        Try
            Session("Check") = True
            CarregarDataGrid()
            If CType(Session("objTitulos"), DataSet).Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Não existem dados para esta coleção de parametros")
            End If
            txtNumValor.Text = SomaTitulosRemessa().ToString()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkRemessa_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkRemessa.Click
        Try
            If ddlCarteira.SelectedItem.Text.Contains("ITAÚ") Then
                RemessaItau(341) 'Banco Itau
            ElseIf ddlCarteira.SelectedItem.Text.Contains("BRADESCO") Then
                GerarRemessa() 'Banco Bradesco
            ElseIf ddlCarteira.SelectedItem.Text.Contains("SAFRA") Then
                RemessaSafra() 'Banco Safra
            ElseIf ddlCarteira.SelectedItem.Text.Contains("BRASIL") Then
                'RemessaBrasil() 'Banco do Brasil
                RemessaItau(1) 'Banco do Brasil
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.ToString())
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "PagamentoDeTitulos")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

#End Region

    Protected Sub btnDownload_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnDownload.Click
        Try
            Download()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

End Class
