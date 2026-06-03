Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis
Imports OfficeOpenXml.Drawing

Public Class ucEmitirCTe
    Inherits BaseUserControl

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If Not IsPostBack And IsConnect Then
            'ddl.Carregar(ddlSituacao, CarregarDDL.Tabela.Situacao, "Situacao_id in (1,2,3,4,5,6,50)")

            Dim TipoDeFrete As New Dictionary(Of String, String)
            TipoDeFrete.Add("   ", "----- [SELECIONE] -----")
            TipoDeFrete.Add("CIF", "POR CONTA DO REMETENTE")    'CIF - 0 - Por conta do Remetente
            TipoDeFrete.Add("FOB", "POR CONTA DO DESTINATÁRIO") 'FOB - 1 - Por conta do Destinatário
            TipoDeFrete.Add("TER", "POR CONTA DE TERCEIROS")    '      2 - Por conta de Terceiros

            ddlFrete.DataTextField = "Value"
            ddlFrete.DataValueField = "Key"
            ddlFrete.DataSource = TipoDeFrete
            ddlFrete.DataBind()

            Dim TipoCTe As New Dictionary(Of Integer, String)

            TipoCTe.Add(0, "----- [SELECIONE] -----")
            TipoCTe.Add(4, "Prestação de Serviço de Terceiros c/ Financeiro")
            TipoCTe.Add(5, "Prestação de Serviço de Terceiros s/ Financeiro")
            TipoCTe.Add(6, "Complemento de Frete")

            ddlTipoCTe.DataTextField = "Value"
            ddlTipoCTe.DataValueField = "Key"
            ddlTipoCTe.DataSource = TipoCTe
            ddlTipoCTe.DataBind()

            Me.Limpar()

        End If
    End Sub

    Private Property objConhecimento As [Lib].Negocio.NotaFiscal
        Get
            Return CType(Session("ssConhecimentoDeTransporte" & HID.Value), [Lib].Negocio.NotaFiscal)
        End Get
        Set(value As [Lib].Negocio.NotaFiscal)
            Session("ssConhecimentoDeTransporte" & HID.Value) = CType(value, [Lib].Negocio.NotaFiscal)
        End Set
    End Property

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Overrides Sub Limpar()

        Session.Remove("objNFe" & HID.Value)
        Session.Remove("objTomadorCTe" & HID.Value)
        Session.Remove("objEmbarqueCTe" & HID.Value)
        Session.Remove("objEntregaCTe" & HID.Value)
        Session.Remove("objTransportadorCTe" & HID.Value)
        Session.Remove("objPlacaCTe" & HID.Value)

        ddlFrete.SelectedIndex = 0
        ddlFrete.Enabled = False

        ddlTipoCTe.SelectedIndex = 0
        ddlTipoCTe.Enabled = False

        rdContratoSim.Checked = False
        rdoContratoNao.Checked = False
        rdPedagioSim.Checked = False
        rdPedagioNao.Checked = False

        rdContratoSim.Enabled = False
        rdoContratoNao.Enabled = False
        rdPedagioSim.Enabled = False
        rdPedagioNao.Enabled = False

        btnConsultarNFe.Enabled = True
        btnTomador.Visible = False
        btnEmbarque.Enabled = False
        btnEntrega.Enabled = False

        txtNFe.Enabled = True
        txtNFe.Text = String.Empty

        txtRemetenteNome.Text = String.Empty
        txtRemetenteEndereco.Text = String.Empty
        txtRemetenteComplemento.Text = String.Empty
        txtRemetenteBairro.Text = String.Empty
        txtRemetenteCidade.Text = String.Empty
        txtRemetenteCnpj.Text = String.Empty
        txtRemetenteInscricao.Text = String.Empty

        txtDestinatarioNome.Text = String.Empty
        txtDestinatarioEndereco.Text = String.Empty
        txtDestinatarioComplemento.Text = String.Empty
        txtDestinatarioBairro.Text = String.Empty
        txtDestinatarioCidade.Text = String.Empty
        txtDestinatarioCnpj.Text = String.Empty
        txtDestinatarioInscricao.Text = String.Empty

        txtTomadorNome.Text = String.Empty
        txtTomadorEndereco.Text = String.Empty
        txtTomadorComplemento.Text = String.Empty
        txtTomadorBairro.Text = String.Empty
        txtTomadorCidade.Text = String.Empty
        txtTomadorCnpj.Text = String.Empty
        txtTomadorInscricao.Text = String.Empty
        btnTomador.Visible = False

        txtEmbarqueNome.Text = String.Empty
        txtEmbarqueEndereco.Text = String.Empty
        txtEmbarqueComplemento.Text = String.Empty
        txtEmbarqueBairro.Text = String.Empty
        txtEmbarqueCidade.Text = String.Empty
        txtEmbarqueCnpj.Text = String.Empty
        txtEmbarqueInscricao.Text = String.Empty

        txtEntregaNome.Text = String.Empty
        txtEntregaEndereco.Text = String.Empty
        txtEntregaComplemento.Text = String.Empty
        txtEntregaBairro.Text = String.Empty
        txtEntregaCidade.Text = String.Empty
        txtEntregaCnpj.Text = String.Empty
        txtEntregaInscricao.Text = String.Empty

        txtQuantidade.Text = String.Empty
        txtUnitario.Text = String.Empty
        txtValor.Text = String.Empty

        objConhecimento.EmitindoCTe = False

        Dim dt As New DataTable("NotaFiscal")
        dt.Columns.Add("Nota", Type.GetType("System.String"))
        dt.Columns.Add("TipoDocumento", Type.GetType("System.String"))
        dt.Columns.Add("CodigoOperacao", Type.GetType("System.String"))
        dt.Columns.Add("Produto", Type.GetType("System.String"))
        dt.Columns.Add("CFOP", Type.GetType("System.String"))
        dt.Columns.Add("QuantidadeFisica", Type.GetType("System.String"))
        dt.Columns.Add("QuantidadeFiscal", Type.GetType("System.String"))
        dt.Columns.Add("ValorUnitario", Type.GetType("System.String"))
        dt.Columns.Add("ValorTotal", Type.GetType("System.String"))
        Session("objNFe" & HID.Value) = dt

        gridNFe.DataSource = Nothing
        gridNFe.DataBind()

    End Sub

    Public Overrides Sub Carregar(ByVal obj As IBaseEntity)
        Try
            If Session("objTomadorCTe" & HID.Value) IsNot Nothing Then
                Dim objTomadorCTe As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
                txtTomadorNome.Text = objTomadorCTe.Nome
                txtTomadorEndereco.Text = objTomadorCTe.Endereco
                txtTomadorComplemento.Text = objTomadorCTe.Complemento
                txtTomadorBairro.Text = objTomadorCTe.Bairro
                txtTomadorCidade.Text = objTomadorCTe.Cidade & "/" & objTomadorCTe.CodigoEstado
                txtTomadorCnpj.Text = objTomadorCTe.CodigoFormatado
                txtTomadorInscricao.Text = objTomadorCTe.InscricaoEstadual
                objConhecimento.CodigoTomador = objTomadorCTe.Codigo
                objConhecimento.EnderecoTomador = objTomadorCTe.CodigoEndereco
                Session.Remove("objTomadorCTe" & HID.Value)

            ElseIf Session("objEmbarqueCTe" & HID.Value) IsNot Nothing Then
                Dim objEmbarqueCTe As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
                txtEmbarqueNome.Text = objEmbarqueCTe.Nome
                txtEmbarqueEndereco.Text = objEmbarqueCTe.Endereco
                txtEmbarqueComplemento.Text = objEmbarqueCTe.Complemento
                txtEmbarqueBairro.Text = objEmbarqueCTe.Bairro
                txtEmbarqueCidade.Text = objEmbarqueCTe.Cidade & "/" & objEmbarqueCTe.CodigoEstado
                txtEmbarqueCnpj.Text = objEmbarqueCTe.CodigoFormatado
                txtEmbarqueInscricao.Text = objEmbarqueCTe.InscricaoEstadual
                objConhecimento.CodigoLocalEmbarque = objEmbarqueCTe.Codigo
                objConhecimento.EndLocalEmbarque = objEmbarqueCTe.CodigoEndereco
                Session.Remove("objEmbarqueCTe" & HID.Value)

            ElseIf Session("objEntregaCTe" & HID.Value) IsNot Nothing Then
                Dim objEntregaCTe As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
                txtEntregaNome.Text = objEntregaCTe.Nome
                txtEntregaEndereco.Text = objEntregaCTe.Endereco
                txtEntregaComplemento.Text = objEntregaCTe.Complemento
                txtEntregaBairro.Text = objEntregaCTe.Bairro
                txtEntregaCidade.Text = objEntregaCTe.Cidade & "/" & objEntregaCTe.CodigoEstado
                txtEntregaCnpj.Text = objEntregaCTe.CodigoFormatado
                txtEntregaInscricao.Text = objEntregaCTe.InscricaoEstadual
                objConhecimento.CodigoEntrega = objEntregaCTe.Codigo
                objConhecimento.EnderecoEntrega = objEntregaCTe.CodigoEndereco
                objConhecimento.Entrega = objEntregaCTe
                Session.Remove("objEntregaCTe" & HID.Value)

            ElseIf Session("objTransportadorCTe" & HID.Value) IsNot Nothing Then
                Dim objTransportador As [Lib].Negocio.Cliente = CType(obj, [Lib].Negocio.Cliente)
                txtTransportadorNome.Text = objTransportador.Nome
                txtTransportadorRNTRC.Text = objTransportador.RNTRCTransportador
                txtTransportadorEndereco.Text = objTransportador.Endereco & IIf(objTransportador.Numero.ToString.Length > 0, ", " & objTransportador.Numero.ToString, "")
                txtTransportadorCidade.Text = objTransportador.Cidade & "/" & objTransportador.CodigoEstado
                txtTransportadorCnpj.Text = objTransportador.CodigoFormatado
                txtTransportadorInscricao.Text = objTransportador.InscricaoEstadual
                objConhecimento.CodigoTransportador = objTransportador.Codigo
                objConhecimento.EnderecoTransportador = objTransportador.CodigoEndereco
                objConhecimento.Transportador = objTransportador
                Session.Remove("objTransportadorCTe" & HID.Value)

            ElseIf Session("objPlacaCTe" & HID.Value) IsNot Nothing Then
                Dim objPlaca As [Lib].Negocio.Placa = CType(Session("objPlacaCTe" & HID.Value), [Lib].Negocio.Placa)
                objConhecimento.PlacaTransportador = LTrim(RTrim(objPlaca.Placa01.ToUpper))

                If objConhecimento.PlacaDetalhes.Motorista IsNot Nothing Then
                    txtMotoristaNome.Text = objConhecimento.PlacaDetalhes.Motorista.Nome
                    txtMotoristaEndereco.Text = objConhecimento.PlacaDetalhes.Motorista.Endereco & IIf(objConhecimento.PlacaDetalhes.Motorista.Numero.ToString.Length > 0, ", " & objConhecimento.PlacaDetalhes.Motorista.Numero.ToString, "")
                    txtMotoristaCidade.Text = objConhecimento.PlacaDetalhes.Motorista.Cidade
                    txtMotoristaEstado.Text = objConhecimento.PlacaDetalhes.Motorista.CodigoEstado
                    txtMotoristaCPF.Text = objConhecimento.PlacaDetalhes.Motorista.CodigoFormatado
                    txtMotoristaHabilitacao.Text = objConhecimento.PlacaDetalhes.Motorista.Habilitacao

                    objConhecimento.CodigoMotorista = objConhecimento.PlacaDetalhes.Motorista.Codigo
                    objConhecimento.EnderecoMotorista = objConhecimento.PlacaDetalhes.Motorista.CodigoEndereco
                End If

                If objConhecimento.PlacaDetalhes IsNot Nothing Then

                    objConhecimento.PlacaTransportador = objConhecimento.PlacaDetalhes.Placa01

                    txtCPlaca.Text = objConhecimento.PlacaDetalhes.Placa01
                    txtPlacaCidade.Text = objConhecimento.PlacaDetalhes.CidadePlaca01
                    txtPlacaEstado.Text = objConhecimento.PlacaDetalhes.EstadoPlaca01
                    txtPlacaRNTRC.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca01
                    txtPlacaProprietario.Text = objConhecimento.PlacaDetalhes.CodigoProprietario01

                    If Not String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.Placa02) Then
                        txtCPlaca02.Text = objConhecimento.PlacaDetalhes.Placa02
                        txtPlaca2Cidade.Text = objConhecimento.PlacaDetalhes.CidadePlaca02
                        txtPlaca2Estado.Text = objConhecimento.PlacaDetalhes.EstadoPlaca02
                        txtPlaca2RNTRC.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca02
                        txtPlaca2Proprietario.Text = objConhecimento.PlacaDetalhes.CodigoProprietario02
                    End If

                    If Not String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.Placa03) Then
                        txtCPlaca03.Text = objConhecimento.PlacaDetalhes.Placa03
                        txtPlaca3Cidade.Text = objConhecimento.PlacaDetalhes.CidadePlaca03
                        txtPlaca3Estado.Text = objConhecimento.PlacaDetalhes.EstadoPlaca03
                        txtPlaca3RNTRC.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca03
                        txtPlaca3Proprietario.Text = objConhecimento.PlacaDetalhes.CodigoProprietario03
                    End If

                    If Not String.IsNullOrWhiteSpace(objConhecimento.PlacaDetalhes.Placa04) Then
                        txtCPlaca04.Text = objConhecimento.PlacaDetalhes.Placa04
                        txtPlaca4Cidade.Text = objConhecimento.PlacaDetalhes.CidadePlaca04
                        txtPlaca4Estado.Text = objConhecimento.PlacaDetalhes.EstadoPlaca04
                        txtPlaca4RNTRC.Text = objConhecimento.PlacaDetalhes.RNTRCPlaca04
                        txtPlaca4Proprietario.Text = objConhecimento.PlacaDetalhes.CodigoProprietario04
                    End If

                End If

                Session.Remove("objPlacaCTe" & HID.Value)

            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Overrides Sub Selecionar(ByVal args As String)
        Try
            'Dim cliente As String() = args.Split(";")
            'Dim objCliente As New [Lib].Negocio.Cliente(cliente(0), Convert.ToInt32(cliente(1)))
            'Session(Session("ssTipoRetorno")) = objCliente
            'If Session("ssTipoRetorno") IsNot Nothing Then
            '    If MainUserControl IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(MainUserControl.ClientID) Then
            '        Dim ucName = MainUserControl.ClientID.Split("_")
            '        Dim uc As UserControl = CType(WebHelpers.FindControlRecursive(Me.Page, ucName(ucName.Length - 1)), UserControl)
            '        CType(uc, IBaseUserControl).Carregar(objCliente)
            '    Else
            '        CType(Me.Page, IBasePage).Carregar(objCliente)
            '    End If
            '    Popup.CloseDialog(Me.Page, "divConsultaCliente")
            'End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultarNFe_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnConsultarNFe.Click
        Try

            If Not String.IsNullOrWhiteSpace(txtNFe.Text) AndAlso txtNFe.Text.Replace(".", "").Length <> 44 Then
                MsgBox(Me.Page, "Chave da nota eletrônica deve ter 44 dígitos.")
                Exit Sub
            End If

            If objConhecimento.NotasTrocaOrigem Is Nothing Then
                objConhecimento.NotasTrocaOrigem = New List(Of [Lib].Negocio.NotaFiscal)
            End If

            If Not String.IsNullOrWhiteSpace(txtNFe.Text) Then

                Dim nf As New [Lib].Negocio.NotaFiscal()
                nf.ChaveNFE = txtNFe.Text.Replace(".", "")
                nf.CarregarNotaComChaveXML(nf.ChaveNFE)

                If nf.Codigo > 0 Then
                    objConhecimento.NotasTrocaOrigem.Add(nf)
                    PreencherInformacoes()
                Else
                    MsgBox(Me.Page, "Nfe não foi encontrada.")
                End If

            End If

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConfirmar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkConfirmar.Click
        Try
            Selecionar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkLimpar.Click
        Try
            Me.Limpar()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles lnkFechar.Click
        Try
            Popup.CloseDialog(Me.Page, "divEmitirCTe")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub btnTomador_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTomador.Click
        Session("ssCampo") = "LivreClasse"

        Dim ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)

        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaClientes.MainUserControl = Me
        End If

        Popup.ConsultaDeClientes(Me.Page, "objTomadorCTe" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnEmbarque_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEmbarque.Click
        Session("ssCampo") = "LivreClasse"

        Dim ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)

        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaClientes.MainUserControl = Me
        End If

        Popup.ConsultaDeClientes(Me.Page, "objEmbarqueCTe" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnEntrega_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnEntrega.Click
        Session("ssCampo") = "LivreClasse"

        Dim ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)

        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaClientes.MainUserControl = Me
        End If

        Popup.ConsultaDeClientes(Me.Page, "objEntregaCTe" & HID.Value, "txtNome")
    End Sub

    Protected Sub btnTransportador_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnTransportador.Click
        Session("ssCampo") = "LivreClasse"

        Dim ucConsultaClientes = CType(Me.Page.FindControlRecursive("ucConsultaClientes"), ucConsultaClientes)

        If ucConsultaClientes IsNot Nothing Then
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarHID(HID.Value)
            ucConsultaClientes.MainUserControl = Me
        End If

        Popup.ConsultaDeClientes(Me.Page, "objTransportadorCTe" & HID.Value, "txtNome")
    End Sub


    Protected Sub btnPlacaTrator_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles btnPlacaTrator.Click
        Try
            Dim ucConsultaPlacas = CType(Me.Page.FindControlRecursive("ucConsultaPlacas"), ucConsultaPlacas)
            If ucConsultaPlacas IsNot Nothing Then
                ucConsultaPlacas.Limpar()
                ucConsultaPlacas.SetarHID(HID.Value)
                ucConsultaPlacas.MainUserControl = Me
            End If

            Popup.ConsultaDePlacas(Me.Page, "objPlacaCTe" & HID.Value)
            Popup.CenterDialog(Me.Page, "divConsultaPlacas")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub PreencherInformacoes()

        For Each nfe In objConhecimento.NotasTrocaOrigem

            ddlFrete.Enabled = True
            ddlTipoCTe.Enabled = True
            rdoContratoNao.Enabled = True
            rdPedagioSim.Enabled = True
            rdPedagioNao.Enabled = True
            btnEmbarque.Enabled = True
            btnEntrega.Enabled = True

            txtNFe.Enabled = False
            btnConsultarNFe.Enabled = False

            objConhecimento.ChaveNFE = txtNFe.Text.Replace(".", "")

            'NO CONHECIMENTO VAMOS USAR O PROPRIETÁTIO COMO REMENTE DA NOTA FISCAL
            txtRemetenteNome.Text = nfe.Empresa.Nome
            txtRemetenteEndereco.Text = nfe.Empresa.Endereco
            txtRemetenteComplemento.Text = nfe.Empresa.Complemento
            txtRemetenteBairro.Text = nfe.Empresa.Bairro
            txtRemetenteCidade.Text = nfe.Empresa.Cidade & "/" & nfe.Empresa.CodigoEstado
            txtRemetenteCnpj.Text = nfe.Empresa.CodigoFormatado
            txtRemetenteInscricao.Text = nfe.Empresa.InscricaoEstadual
            objConhecimento.CodigoProprietarioDaMercadoria = nfe.CodigoEmpresa
            objConhecimento.EnderecoProprietarioDaMercadoria = nfe.EnderecoEmpresa


            'NO CONHECIMENTO VAMOS USAR O CLIENTE DA NOTA FISCAL COMO DESTINATÁRIO DA MERCADORIA 
            txtDestinatarioNome.Text = nfe.Cliente.Nome
            txtDestinatarioEndereco.Text = nfe.Cliente.Endereco
            txtDestinatarioComplemento.Text = nfe.Cliente.Complemento
            txtDestinatarioBairro.Text = nfe.Cliente.Bairro
            txtDestinatarioCidade.Text = nfe.Cliente.Cidade & "/" & nfe.Cliente.CodigoEstado
            txtDestinatarioCnpj.Text = nfe.Cliente.CodigoFormatado
            txtDestinatarioInscricao.Text = nfe.Cliente.InscricaoEstadual
            objConhecimento.CodigoCliente = nfe.CodigoCliente
            objConhecimento.EnderecoCliente = nfe.EnderecoCliente


            'NO CONHECIMENTO VAMOS USAR O PROPRIETÁTIO COMO LOCAL DE EMBARQUE PODENDO TROCAR DEPOIS
            txtEmbarqueNome.Text = objConhecimento.ProprietarioDaMercadoria.Nome
            txtEmbarqueEndereco.Text = objConhecimento.ProprietarioDaMercadoria.Endereco
            txtEmbarqueComplemento.Text = objConhecimento.ProprietarioDaMercadoria.Complemento
            txtEmbarqueBairro.Text = objConhecimento.ProprietarioDaMercadoria.Bairro
            txtEmbarqueCidade.Text = objConhecimento.ProprietarioDaMercadoria.Cidade & "/" & objConhecimento.ProprietarioDaMercadoria.CodigoEstado
            txtEmbarqueCnpj.Text = objConhecimento.ProprietarioDaMercadoria.CodigoFormatado
            txtEmbarqueInscricao.Text = objConhecimento.ProprietarioDaMercadoria.InscricaoEstadual
            objConhecimento.CodigoLocalEmbarque = objConhecimento.CodigoProprietarioDaMercadoria
            objConhecimento.EndLocalEmbarque = objConhecimento.EnderecoProprietarioDaMercadoria


            'NO CONHECIMENTO VAMOS USAR O CLIENTE DA NOTA FISCAL COMO DESTINATÁRIO/ENTREGA DA MERCADORIA 
            txtEntregaNome.Text = objConhecimento.Cliente.Nome
            txtEntregaEndereco.Text = objConhecimento.Cliente.Endereco
            txtEntregaComplemento.Text = objConhecimento.Cliente.Complemento
            txtEntregaBairro.Text = objConhecimento.Cliente.Bairro
            txtEntregaCidade.Text = objConhecimento.Cliente.Cidade & "/" & objConhecimento.Cliente.CodigoEstado
            txtEntregaCnpj.Text = objConhecimento.Cliente.CodigoFormatado
            txtEntregaInscricao.Text = objConhecimento.Cliente.InscricaoEstadual
            objConhecimento.CodigoEntrega = objConhecimento.CodigoCliente
            objConhecimento.EnderecoEntrega = objConhecimento.EnderecoCliente


            If nfe.CodigoTransportador.Length > 0 Then
                txtTransportadorNome.Text = nfe.Transportador.Nome
                txtTransportadorEndereco.Text = nfe.Transportador.Endereco
                txtTransportadorCidade.Text = nfe.Transportador.Cidade & "/" & nfe.Transportador.CodigoEstado
                txtTransportadorCnpj.Text = nfe.Transportador.CodigoFormatado
                txtTransportadorInscricao.Text = nfe.Transportador.InscricaoEstadual

                objConhecimento.CodigoTransportador = nfe.CodigoTransportador
                objConhecimento.EnderecoTransportador = nfe.EnderecoTransportador
            End If

            If nfe.PlacaTransportador.Length > 0 Then

                objConhecimento.PlacaTransportador = nfe.PlacaTransportador
                txtCPlaca.Text = nfe.PlacaTransportador
                txtPlacaCidade.Text = nfe.PlacaDetalhes.CidadePlaca01
                txtPlacaEstado.Text = nfe.PlacaDetalhes.EstadoPlaca01
                txtPlacaRNTRC.Text = nfe.PlacaDetalhes.RNTRCPlaca01
                txtPlacaProprietario.Text = nfe.PlacaDetalhes.Proprietario01.CodigoFormatado & " - " & nfe.PlacaDetalhes.Proprietario01.Nome

                If nfe.PlacaDetalhes.Placa02.Length > 0 Then
                    txtCPlaca02.Text = nfe.PlacaDetalhes.Placa02
                    txtPlaca2Cidade.Text = nfe.PlacaDetalhes.CidadePlaca02
                    txtPlaca2Estado.Text = nfe.PlacaDetalhes.EstadoPlaca02
                    txtPlaca2RNTRC.Text = nfe.PlacaDetalhes.RNTRCPlaca02
                    txtPlaca2Proprietario.Text = nfe.PlacaDetalhes.Proprietario02.CodigoFormatado & " - " & nfe.PlacaDetalhes.Proprietario02.Nome
                End If

                If nfe.PlacaDetalhes.Placa03.Length > 0 Then
                    txtCPlaca03.Text = nfe.PlacaDetalhes.Placa03
                    txtPlaca3Cidade.Text = nfe.PlacaDetalhes.CidadePlaca03
                    txtPlaca3Estado.Text = nfe.PlacaDetalhes.EstadoPlaca03
                    txtPlaca3RNTRC.Text = nfe.PlacaDetalhes.RNTRCPlaca03
                    txtPlaca3Proprietario.Text = nfe.PlacaDetalhes.Proprietario03.CodigoFormatado & " - " & nfe.PlacaDetalhes.Proprietario03.Nome
                End If

                If nfe.PlacaDetalhes.Placa04.Length > 0 Then
                    txtCPlaca04.Text = nfe.PlacaDetalhes.Placa04
                    txtPlaca4Cidade.Text = nfe.PlacaDetalhes.CidadePlaca04
                    txtPlaca4Estado.Text = nfe.PlacaDetalhes.EstadoPlaca04
                    txtPlaca4RNTRC.Text = nfe.PlacaDetalhes.RNTRCPlaca04
                    txtPlaca4Proprietario.Text = nfe.PlacaDetalhes.Proprietario04.CodigoFormatado & " - " & nfe.PlacaDetalhes.Proprietario04.Nome
                End If
            End If

            Exit For

        Next

        If objConhecimento.NotasTrocaOrigem.Count > 0 Then
            txtQuantidade.Text = objConhecimento.NotasTrocaOrigem.Sum(Function(s) s.TotalQuantidadeFiscal)
            objConhecimento.EmitindoCTe = True

            'pesquisa a lista de notas selecionadas para exbir todos os produtos na gridview
            For Each item As [Lib].Negocio.NotaFiscalXItem In objConhecimento.NotasTrocaOrigem.SelectMany(Function(s) s.Itens).ToList
                Dim dr As DataRow = CType(Session("objNFe" & HID.Value), DataTable).NewRow()
                dr("Nota") = IIf(item.NotaFiscal.EntradaSaida = [Lib].Negocio.eEntradaSaida.Saida, "S", "E") & "-" & item.NotaFiscal.Codigo & "-" & item.NotaFiscal.Serie
                dr("TipoDocumento") = item.NotaFiscal.CodigoTipoDeDocumento & "-" & item.NotaFiscal.TipoDeDocumento.Descricao
                dr("CodigoOperacao") = item.CodigoOperacao & "-" & item.CodigoSubOperacao
                dr("Produto") = item.CodigoProduto & "-" & item.Produto.Nome
                dr("CFOP") = item.NotaFiscal.NaturezaDaOperacao
                dr("QuantidadeFisica") = item.QuantidadeFisica.ToString("N0")
                dr("QuantidadeFiscal") = item.QuantidadeFiscal.ToString("N0")
                dr("ValorUnitario") = item.Unitario.ToString("N2")
                dr("ValorTotal") = item.ValorTotal.ToString("N2")
                CType(Session("objNFe" & HID.Value), DataTable).Rows.Add(dr)
            Next

            gridNFe.DataSource = CType(Session("objNFe" & HID.Value), DataTable)
            gridNFe.DataBind()
        End If

    End Sub

    Protected Sub ddlFrete_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlFrete.SelectedIndexChanged
        Try
            If objConhecimento.ChaveNFE.Length = 0 Then
                MsgBox(Me.Page, "Nota Fiscal Eletrônica ainda não foi selecionada.")
                Exit Sub
            End If

            objConhecimento.CIFFOB = ([Enum].Parse(GetType(eTiposFrete), ddlFrete.SelectedValue))

            btnTomador.Visible = False

            For Each nfe In objConhecimento.NotasTrocaOrigem

                If objConhecimento.CIFFOB = eTiposFrete.CIF Then
                    txtTomadorNome.Text = nfe.Empresa.Nome
                    txtTomadorEndereco.Text = nfe.Empresa.Endereco
                    txtTomadorComplemento.Text = nfe.Empresa.Complemento
                    txtTomadorBairro.Text = nfe.Empresa.Bairro
                    txtTomadorCidade.Text = nfe.Empresa.Cidade & "/" & nfe.Empresa.CodigoEstado
                    txtTomadorCnpj.Text = nfe.Empresa.CodigoFormatado
                    txtTomadorInscricao.Text = nfe.Empresa.InscricaoEstadual

                    objConhecimento.CodigoTomador = nfe.CodigoEmpresa
                    objConhecimento.EnderecoTomador = nfe.EnderecoEmpresa
                End If

                If objConhecimento.CIFFOB = eTiposFrete.FOB Then
                    txtTomadorNome.Text = nfe.Cliente.Nome
                    txtTomadorEndereco.Text = nfe.Cliente.Endereco
                    txtTomadorComplemento.Text = nfe.Cliente.Complemento
                    txtTomadorBairro.Text = nfe.Cliente.Bairro
                    txtTomadorCidade.Text = nfe.Cliente.Cidade & "/" & nfe.Cliente.CodigoEstado
                    txtTomadorCnpj.Text = nfe.Cliente.CodigoFormatado
                    txtTomadorInscricao.Text = nfe.Cliente.InscricaoEstadual

                    objConhecimento.CodigoTomador = nfe.CodigoCliente
                    objConhecimento.EnderecoTomador = nfe.EnderecoCliente
                End If

                If objConhecimento.CIFFOB = eTiposFrete.TER Then
                    btnTomador.Visible = True

                    txtTomadorNome.Text = String.Empty
                    txtTomadorEndereco.Text = String.Empty
                    txtTomadorComplemento.Text = String.Empty
                    txtTomadorBairro.Text = String.Empty
                    txtTomadorCidade.Text = String.Empty
                    txtTomadorCnpj.Text = String.Empty
                    txtTomadorInscricao.Text = String.Empty

                    objConhecimento.CodigoTomador = String.Empty
                    objConhecimento.EnderecoTomador = 0

                End If

                Exit For
            Next

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub


    Protected Sub rdContratoSim_CheckedChanged(sender As Object, e As EventArgs) Handles rdContratoSim.CheckedChanged
        If objConhecimento.ChaveNFE.Length = 0 Then
            MsgBox(Me.Page, "Nota Fiscal Eletrônica ainda não foi selecionada.")
            rdContratoSim.Checked = False
            rdoContratoNao.Checked = False
            Exit Sub
        End If

        objConhecimento.ContratoDeFreteCTe = True
    End Sub

    Protected Sub rdoContratoNao_CheckedChanged(sender As Object, e As EventArgs) Handles rdoContratoNao.CheckedChanged
        If objConhecimento.ChaveNFE.Length = 0 Then
            MsgBox(Me.Page, "Nota Fiscal Eletrônica ainda não foi selecionada.")
            rdContratoSim.Checked = False
            rdoContratoNao.Checked = False
            Exit Sub
        End If

        objConhecimento.ContratoDeFreteCTe = False
    End Sub

    Protected Sub rdPedagioSim_CheckedChanged(sender As Object, e As EventArgs) Handles rdPedagioSim.CheckedChanged
        If objConhecimento.ChaveNFE.Length = 0 Then
            MsgBox(Me.Page, "Nota Fiscal Eletrônica ainda não foi selecionada.")
            rdPedagioSim.Checked = False
            rdPedagioNao.Checked = False
            Exit Sub
        End If

        objConhecimento.TemPedagioCTe = True
    End Sub

    Protected Sub rdPedagioNao_CheckedChanged(sender As Object, e As EventArgs) Handles rdPedagioNao.CheckedChanged
        If objConhecimento.ChaveNFE.Length = 0 Then
            MsgBox(Me.Page, "Nota Fiscal Eletrônica ainda não foi selecionada.")
            rdPedagioSim.Checked = False
            rdPedagioNao.Checked = False
            Exit Sub
        End If

        objConhecimento.TemPedagioCTe = False
    End Sub

    Protected Sub ddlTipoCTe_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlTipoCTe.SelectedIndexChanged
        Try
            If CType(CInt(ddlTipoCTe.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro Then
                MsgBox(Me.Page, "Opção ainda não está disponível.")
                ddlTipoCTe.SelectedIndex = 0
            ElseIf CType(CInt(ddlTipoCTe.SelectedValue), eTipoDeDocumentoFrete) = eTipoDeDocumentoFrete.Complemento Then
                MsgBox(Me.Page, "Opção ainda não está disponível.")
                ddlTipoCTe.SelectedIndex = 0
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

End Class