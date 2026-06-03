Imports System.Data
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class CertidaoNegativa
    Inherits BasePage

    Private sql As String
    Private lstHistorico As [Lib].Negocio.ListCertidaoNegativa
    Private lst As [Lib].Negocio.Estados
    Protected objCliente As [Lib].Negocio.Cliente

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Funcoes.VerificaPermissao("CertidaoNegativa", "ACESSAR") Then
                CarregarModelo()
                CarregarEstados()
                Limpar()
            Else
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx")
                Exit Sub
            End If
        End If
        AtribuirValores()
    End Sub

    Private Sub CarregarModelo()
        ddl.Carregar(ddlModeloCN, CarregarDDL.Tabela.ModeloCertidaoNegativa, "", False)
    End Sub

    Private Sub CarregarEstados()
        chkEstados.Items.Clear()
        chkEstados.DataValueField = "Codigo"
        chkEstados.DataTextField = "Codigo"
        lst = New [Lib].Negocio.Estados(True, "Estado_Id not in('EX','SX','DR')")
        Session("LEST") = lst
        chkEstados.DataSource = lst
        chkEstados.DataBind()
    End Sub

    Private Sub CarregarHistorico()
        Dim cli As String()
        cli = txtCodigoCliente.Value.Split(";")

        If String.IsNullOrWhiteSpace(cli(0)) Then Exit Sub

        Dim lst As New [Lib].Negocio.ListCertidaoNegativa

        If rdConsolidado.Checked Then
            lst.Historico(cli(0), True)
        Else
            lst.Historico(cli(0), False)
        End If

        gridHistorico.DataSource = lst.ToArray
        gridHistorico.DataBind()
        Session("Historico") = lst
    End Sub

    Private Sub AtribuirValores()
        If Not Session("objClienteCN") Is Nothing Then
            Dim cliente As [Lib].Negocio.Cliente = CType(Session("objClienteCN"), [Lib].Negocio.Cliente)
            With cliente
                txtCliente.Text = .CodigoFormatado & " : " & .Nome & " - " & RTrim(.Cidade) & "/" & .CodigoEstado
                txtCodigoCliente.Value = .Codigo & ";" & .CodigoEndereco
                Session.Remove("objClienteCN")
                CarregarHistorico()
            End With
        End If
    End Sub

    Function EstadosSelecionado() As String
        lst = Session("LEST")
        Dim UF As String = ""
        For i As Integer = 0 To chkEstados.Items.Count - 1
            If chkEstados.Items(i).Selected = True Then
                UF &= IIf(UF <> "", ",", "") & "'" & chkEstados.Items(i).Value & "'"
            End If
        Next
        If UF.Length > 0 Then
            Return "(" & UF & ")"
        Else
            Return ""
        End If
    End Function

    Protected Sub cmdConsultaClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            hdnControlePopup.Value = "Consulta"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objClienteCertNegativa" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnConsultaPrincipal_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim lst As New [Lib].Negocio.ListCertidaoNegativa
            Dim EmissaoInicial As String = ""
            Dim EmissaoFinal As String = ""

            If Not IsDate(txtDataEmissaoInicial.Text) Or Not IsDate(txtDataEmissaoFinal.Text) Then
                chkEmissao.Checked = False
            ElseIf chkEmissao.Checked Then
                EmissaoInicial = txtDataEmissaoInicial.Text
                EmissaoFinal = txtDataEmissaoFinal.Text
            End If

            Dim ValidadeInicial As String = ""
            Dim ValidadeFinal As String = ""
            If Not IsDate(txtDataValidadeInicial.Text) Or Not IsDate(txtDataValidadeFinal.Text) Then
                chkValidade.Checked = False
            ElseIf chkValidade.Checked Then
                ValidadeInicial = txtDataValidadeInicial.Text
                ValidadeFinal = txtDataValidadeFinal.Text
            End If

            Dim cn As New [Lib].Negocio.CertidaoNegativa
            If txtCodigoClienteConsulta.Value <> "" Then cn.CodigoCliente = txtCodigoClienteConsulta.Value.Split(";")(0)
            If txtNumero.Text <> "" Then cn.Numero = txtCodigoNumero.Text
            If txtCodigoAutenticidade.Text <> "" Then cn.CodigoAutenticidade = txtCodigoAutenticidade.Text

            lst.UltimasInformadas(cn, EmissaoInicial, EmissaoFinal, ValidadeInicial, ValidadeFinal, txtIE.Text, EstadosSelecionado)

            gridConsulta.DataSource = lst.ToArray
            gridConsulta.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnComCertidao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsNumeric(txtDias01.Text) Then
                MsgBox(Me.Page, "Informe o numero de dias para consulta")
                Exit Sub
            End If

            sql = "select (case when len(cn.cliente_id) = 14 " & vbCrLf & _
                  "               then left(cn.cliente_id,8) " & vbCrLf & _
                  "               else cn.cliente_id " & vbCrLf & _
                  "         end) as Cliente, " & vbCrLf & _
                  "       max(cn.DataValidade) as DataValidade " & vbCrLf & _
                  "  into #temp  " & vbCrLf & _
                  "  from certidaonegativa cn " & vbCrLf & _
                  " where datavalidade between getdate() and getdate() + " & txtDias01.Text & vbCrLf

            Dim uf As String = EstadosSelecionado()
            If uf.Length > 0 Then
                sql &= " and exists(select 1 from clientes cl where cl.cliente_id = CN.Cliente_id and cl.estado in " & uf & ")"
            End If

            sql &= " group by(case " & vbCrLf & _
                  "            when len(cn.cliente_id) = 14 " & vbCrLf & _
                  "              then left(cn.cliente_id,8) " & vbCrLf & _
                  "              else cn.cliente_id " & vbCrLf & _
                  "          end) " & vbCrLf

            sql &= " select t.cliente," & vbCrLf & _
                   "      (select top 1 cl.nome" & vbCrLf & _
                   "         from clientes cl" & vbCrLf & _
                   "        where case when len(t.cliente) = 8" & vbCrLf & _
                   "                     then left(cl.cliente_id,8)" & vbCrLf & _
                   "                     else cl.cliente_id" & vbCrLf & _
                   "              end = t.cliente) as nome," & vbCrLf & _
                   "       t.DataValidade" & vbCrLf & _
                   "  from #temp t" & vbCrLf & _
                   " order by 2" & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "Consulta01")

            Funcoes.BindReport(Me.Page, ds, "cr_CertidoesNegativasAVencer", eExportType.PDF)
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnSemCertidao_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            If Not IsNumeric(txtDias02.Text) Then
                MsgBox(Me.Page, "Informe o numero de dias para consulta")
                Exit Sub
            End If

            sql = "select(case when len(nf.cliente_id) = 14" & vbCrLf & _
                  "              then left(nf.cliente_id,8)" & vbCrLf & _
                  "              else nf.cliente_id" & vbCrLf & _
                  "       end) as Cliente," & vbCrLf & _
                  "      (select max(cn.datavalidade)" & vbCrLf & _
                  "         from certidaonegativa cn" & vbCrLf & _
                  "        where case when len(cn.cliente_id) = 14" & vbCrLf & _
                  "                     then left(cn.cliente_id,8)" & vbCrLf & _
                  "                     else cn.cliente_id" & vbCrLf & _
                  "              end = case when len(nf.cliente_id) = 14" & vbCrLf & _
                  "                          then left(nf.cliente_id,8)" & vbCrLf & _
                  "                          else nf.cliente_id" & vbCrLf & _
                  "                    end" & vbCrLf & _
                  "        ) as DataUltimaCN," & vbCrLf & _
                  "       max(nf.datadanota) as DataUltimaNota" & vbCrLf & _
                  "  Into #Temp" & vbCrLf & _
                  "  from notasfiscais nf" & vbCrLf & _
                  " inner Join Clientes C" & vbCrLf & _
                  "    on nf.cliente_id     = C.Cliente_id" & vbCrLf & _
                  "   and nf.endCliente_id = C.Endereco_id" & vbCrLf & _
                  " where nf.DataDaNota > getdate()- " & txtDias02.Text & vbCrLf & _
                  "   and exists (select 1 from clientesxtipos ct Where ct.cliente_id = nf.cliente_id and ct.tipo_id in(4,5))" & vbCrLf & _
                  "   and C.situacao = 1" & vbCrLf & _
                  " group by case" & vbCrLf & _
                  "            when len(nf.cliente_id) = 14" & vbCrLf & _
                  "              then left(nf.cliente_id,8)" & vbCrLf & _
                  "              else nf.cliente_id" & vbCrLf & _
                  "          end" & vbCrLf

            '"   and not exists (select 1 from certidaonegativa cn where cn.cliente_id = nf.cliente_id and cn.datavalidade > getdate())" & vbCrLf & _

            sql &= "select t.cliente," & vbCrLf & _
                   "      (select top 1 cl.nome" & vbCrLf & _
                   "         from clientes cl" & vbCrLf & _
                   "        where case when len(t.cliente) = 8" & vbCrLf & _
                   "                     then left(cl.cliente_id,8)" & vbCrLf & _
                   "                     else cl.cliente_id" & vbCrLf & _
                   "              end = t.cliente) as nome," & vbCrLf & _
                   "       t.DataUltimaCN," & vbCrLf & _
                   "       t.DataUltimaNota" & vbCrLf & _
                   "  from #Temp t" & vbCrLf & _
                   " where t.DataUltimaCN < getdate()" & vbCrLf & _
                   " Order by 2"

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(sql, "Consulta02")

            Funcoes.BindReport(Me.Page, ds, "cr_CertidoesNegativasClientesSem", eExportType.PDF)

        Catch ex As Exception
            MsgBox(Me.Page, ex.Message.ToString)
        End Try
    End Sub

    Protected Sub btnNossaEmpresa_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim lst As New [Lib].Negocio.ListCertidaoNegativa

            Dim cn As New [Lib].Negocio.CertidaoNegativa
            If txtCodigoClienteConsulta.Value <> "" Then cn.CodigoCliente = txtCodigoClienteConsulta.Value.Split(";")(0)
            If txtNumero.Text <> "" Then cn.Numero = txtCodigoNumero.Text
            If txtCodigoAutenticidade.Text <> "" Then cn.CodigoAutenticidade = txtCodigoAutenticidade.Text

            lst.UltimasInformadas(cn, "", "", "", "", "", "", False, "S")

            gridConsulta.DataSource = lst.ToArray
            gridConsulta.DataBind()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub gridHistorico_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            Dim lst As [Lib].Negocio.ListCertidaoNegativa
            lst = Session("Historico")
            Dim cn As [Lib].Negocio.CertidaoNegativa = lst(gridHistorico.SelectedIndex)

            ddlModeloCN.SelectedIndex = cn.CodigoModeloCertidao
            txtCliente.Text = cn.NomeCliente
            txtCodigoCliente.Value = cn.CodigoCliente & ";0"
            txtNumero.Text = cn.Numero
            txtDataEmissao.Text = cn.DataEmissao.ToString("dd/MM/yyyy")
            txtDataValidade.Text = cn.DataValidade.ToString("dd/MM/yyyy")
            txtCodigoAutenticidade.Text = cn.CodigoAutenticidade

            lnkNovo.Enabled = False
            lnkAtualizar.Enabled = True
            lnkExcluir.Enabled = True

            cmdClientes.Enabled = False
            txtNumero.Enabled = False
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdConsolidado_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarHistorico()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub rdIndividual_CheckedChanged(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            CarregarHistorico()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub cmdClientes_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            hdnControlePopup.Value = "Manutencao"
            ucConsultaClientes.Limpar()
            ucConsultaClientes.SetarTituloDIV("Consulta de Clientes")
            Popup.ConsultaDeClientes(Me.Page, "objClienteCertNegativa" & HID.Value.ToString)
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Public Sub Limpar()
        Session.Remove("objClienteCertNegativa" & HID.Value.ToString)
        HID.Value = Guid.NewGuid().ToString
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteCertNegativa" & HID.Value.ToString) Is Nothing Then
            objCliente = CType(Session("objClienteCertNegativa" & HID.Value.ToString), [Lib].Negocio.Cliente)
            If (hdnControlePopup.Value.Equals("Manutencao")) Then
                txtCliente.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoCliente.Value = objCliente.Codigo & ";" & objCliente.CodigoEndereco
            Else
                txtClienteConsulta.Text = Funcoes.AlinharEsquerda(objCliente.Nome, 28, ".") & " - " & Funcoes.AlinharEsquerda(objCliente.Cidade, 20, ".") & " " & objCliente.CodigoEstado & " " & Funcoes.AlinharEsquerda(objCliente.CodigoFormatado, 18, ".") & "-" & CStr(objCliente.CodigoEndereco) & "-" & objCliente.Reduzido
                txtCodigoClienteConsulta.Value = objCliente.Codigo & ";" & objCliente.CodigoEndereco
            End If
        End If
    End Sub

    Private Function verificaDatas() As Boolean
        If String.IsNullOrWhiteSpace(txtDataEmissao.Text) OrElse String.IsNullOrWhiteSpace(txtDataValidade.Text) _
           OrElse Not IsDate(txtDataEmissao.Text) OrElse Not IsDate(txtDataValidade.Text) Then
            MsgBox(Me.Page, "Informe a data de emissão e a data de validade.")
            Return False
        ElseIf CDate(txtDataEmissao.Text) > CDate(txtDataValidade.Text) Then
            MsgBox(Me.Page, "Data de emissão não pode ser maior que a data de validade.")
            Return False
        End If
        Return True
    End Function

    Protected Sub lnkNovo_Click(sender As Object, e As EventArgs) Handles lnkNovo.Click
        Try
            If Funcoes.VerificaPermissao("CertidaoNegativa", "GRAVAR") Then
                If verificaDatas() Then
                    Dim cn As New [Lib].Negocio.CertidaoNegativa
                    cn.IUD = "I"
                    cn.CodigoCliente = txtCodigoCliente.Value.Split(";")(0)
                    cn.Numero = txtNumero.Text
                    cn.CodigoModeloCertidao = ddlModeloCN.SelectedValue
                    cn.CodigoAutenticidade = txtAutenticidade.Text
                    cn.DataEmissao = CDate(txtDataEmissao.Text)
                    cn.DataValidade = CDate(txtDataValidade.Text)

                    If String.IsNullOrWhiteSpace(cn.CodigoCliente) Then
                        MsgBox(Me.Page, "Recarregue o cliente.")
                        txtCodigoCliente.Value = ""
                        txtCliente.Text = ""
                        Exit Sub
                    End If

                    If cn.Salvar() Then
                        Dim lst As New [Lib].Negocio.ListCertidaoNegativa
                        CarregarHistorico()
                        gridHistorico.DataSource = lst.ToArray
                        lnkLimpar_Click(Nothing, Nothing)
                        MsgBox(Me.Page, "Registro salvo com Sucesso.", eTitulo.Sucess)
                    Else
                        MsgBox(Me.Page, "Erro durante o processo.")
                    End If
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
            If Funcoes.VerificaPermissao("CertidaoNegativa", "ALTERAR") Then
                If verificaDatas() Then
                    Dim cn As New [Lib].Negocio.CertidaoNegativa
                    cn.IUD = "U"
                    cn.CodigoCliente = txtCodigoCliente.Value.Split(";")(0)
                    cn.Numero = txtNumero.Text
                    cn.CodigoModeloCertidao = ddlModeloCN.SelectedValue
                    cn.CodigoAutenticidade = txtAutenticidade.Text
                    cn.DataEmissao = CDate(txtDataEmissao.Text)
                    cn.DataValidade = CDate(txtDataValidade.Text)

                    If String.IsNullOrWhiteSpace(cn.CodigoCliente) Then
                        MsgBox(Me.Page, "Recarregue o cliente.")
                        txtCodigoCliente.Value = ""
                        txtCliente.Text = ""
                        Exit Sub
                    End If

                    If cn.Salvar() Then
                        Dim lst As New [Lib].Negocio.ListCertidaoNegativa
                        CarregarHistorico()
                        gridHistorico.DataSource = lst.ToArray
                        lnkLimpar_Click(Nothing, Nothing)
                        MsgBox(Me.Page, "Registro alterado com Sucesso.", eTitulo.Sucess)
                    End If
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão alterar registro.")
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message)
        End Try
    End Sub

    Protected Sub lnkExcluir_Click(sender As Object, e As EventArgs) Handles lnkExcluir.Click
        Try
            If Funcoes.VerificaPermissao("CertidaoNegativa", "EXCLUIR") Then
                Dim cn As New [Lib].Negocio.CertidaoNegativa
                cn.IUD = "D"
                cn.CodigoCliente = txtCodigoCliente.Value.Split(";")(0)
                cn.Numero = txtNumero.Text

                If cn.Salvar() Then
                    Dim lst As New [Lib].Negocio.ListCertidaoNegativa
                    CarregarHistorico()
                    gridHistorico.DataSource = lst.ToArray
                    lnkLimpar_Click(Nothing, Nothing)
                    MsgBox(Me.Page, "Registro deletado com Sucesso.", eTitulo.Sucess)
                Else
                    MsgBox(Me.Page, "Erro durante o processo.")
                End If
            Else
                MsgBox(Me.Page, "Usuário sem permissão para excluir registro.", eTitulo.Info)
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        Try
            txtCliente.Text = ""
            txtCodigoCliente.Value = ""
            txtNumero.Text = ""
            txtAutenticidade.Text = ""
            txtDataEmissao.Text = ""
            txtDataValidade.Text = ""
            ddlModeloCN.SelectedIndex = 0
            lnkNovo.Enabled = True
            lnkAtualizar.Enabled = False
            lnkExcluir.Enabled = False
            cmdClientes.Enabled = True
            txtNumero.Enabled = True
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "CertidaoNegativa")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class