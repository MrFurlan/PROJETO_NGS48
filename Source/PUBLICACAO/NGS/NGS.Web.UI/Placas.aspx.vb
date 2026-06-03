Imports System.Data
Imports System.Collections
Imports System.IO
Imports System.Text
Imports CrystalDecisions.CrystalReports.Engine
Imports CrystalDecisions.Shared
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class Placas
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Gestao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("Placas", "ACESSAR") Then
                    ddl.Carregar(ddlTipoVeiculo, CarregarDDL.Tabela.TipoDeVeiculo, "", True)
                    ddl.Carregar(ddlViaTransporte, CarregarDDL.Tabela.ViaDeTransportes, "", True)
                    Limpar()
                Else
                    MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Gestao.aspx")
                    Exit Sub
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub GridPlacas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridPlacas.SelectedIndexChanged
        Try
            Limpar()
            txtPlaca.Text = GridPlacas.SelectedRow.Cells(1).Text()
            txtProprietario.Text = GridPlacas.SelectedRow.Cells(2).Text()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub Limpar()
        txtPlaca.Text = ""
        hdfProprietario.Value = ""
        txtProprietario.Text = ""
        ddlTipoVeiculo.SelectedIndex = 0
        ddlViaTransporte.SelectedIndex = 0
        HID.Value = Guid.NewGuid().ToString()
        ucConsultaClientes.SetarHID(HID.Value)
    End Sub

    Private Function getParam() As String
        Dim param As String = "Parametros:" & vbCrLf
        If Not String.IsNullOrWhiteSpace(txtPlaca.Text) Then
            param &= "Placa: " & txtPlaca.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtProprietario.Text) Then
            param &= "Proprietário: " & txtProprietario.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlViaTransporte.SelectedValue) Then
            param &= "Via de Transporte: " & ddlViaTransporte.SelectedValue & " - "
        End If
        If Not String.IsNullOrWhiteSpace(ddlTipoVeiculo.SelectedValue) Then
            param &= "Tipo de Veículo: " & ddlTipoVeiculo.SelectedValue & "."
        End If

        Return param
    End Function

    Function ValidarCampos() As Boolean
        If String.IsNullOrWhiteSpace(txtPlaca.Text) Then
            MsgBox(Me.Page, "O campo placa é obrigatório.")
            Return False
        End If
        Return True
    End Function

    Public Overrides Sub Carregar(ByVal obj As [Lib].Negocio.IBaseEntity)
        If Session("objCliente" & HID.Value) IsNot Nothing Then
            Dim pCliente As [Lib].Negocio.Cliente = CType(Session("objCliente" & HID.Value), [Lib].Negocio.Cliente)
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(pCliente)
            hdfProprietario.Value = itemCliente.Value
            txtProprietario.Text = itemCliente.Text
            Session.Remove("objCliente" & HID.Value)
        End If
    End Sub

    Protected Sub cmdProprietario_Click(sender As Object, e As EventArgs) Handles cmdProprietario.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me.Page, "objCliente" & HID.Value, "txtNome")
            Popup.CenterDialog(Me.Page, "divConsultaCliente")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        Try
            If ValidarCampos() Then
                Dim ds As New DataSet
                Dim Sql As String = ""

                Sql = "SELECT LTRIM(RTRIM(p.Placa_Id)) as Placa_Id,                                         " & vbCrLf & _
                      "LTRIM(RTRIM(p.Placa_Id)) as Placa,                                                   " & vbCrLf & _
                      "p.ViaTransporte_Id as ViaTransporte,                                                 " & vbCrLf & _
                      "p.TipoVeiculo_Id as TipoVeiculo,                                                     " & vbCrLf & _
                      "p.CidadePlaca as CidadePlaca,                                                        " & vbCrLf & _
                      "p.EstadoPlaca as EstadoPlaca,                                                        " & vbCrLf & _
                      "p.RNTRCPlaca as RNTRCPlaca,                                                          " & vbCrLf & _
                      "p.NomeMotorista as NomeMotorista,                                                    " & vbCrLf & _
                      "p.CidadeMotorista as CidadeMotorista,                                                " & vbCrLf & _
                      "p.EstadoMotorista as EstadoMotorista,                                                " & vbCrLf & _
                      "p.Habilitacao as Habilitacao,                                                        " & vbCrLf & _
                      "p.CpfMotorista as CpfMotorista,                                                      " & vbCrLf & _
                      "p.EndCpfMotorista as EndCpfMotorista,                                                " & vbCrLf & _
                      "p.Proprietario,                                                                      " & vbCrLf & _
                      "p.EndProprietario,                                                                   " & vbCrLf & _
                      "p.NomeProprietario,                                                                  " & vbCrLf & _
                      "'S' as Principal                                                                     " & vbCrLf & _
                      "FROM Placas p                                                                        " & vbCrLf & _
                      "        WHERE 1=1                                                                    " & vbCrLf & _
                      "AND p.Placa_Id = ' ADG-1942'                                                                                     " & vbCrLf & _
                      "UNION ALL                                                                            " & vbCrLf & _
                      "SELECT LTRIM(RTRIM(p1.Placa_Id)) as Placa_Id,                                        " & vbCrLf & _
                      "LTRIM(RTRIM(p1.Placa01)) as Placa,                                                   " & vbCrLf & _
                      "p1.ViaTransporte_Id as ViaTransporte,                                                " & vbCrLf & _
                      "p1.TipoVeiculo_Id as TipoVeiculo,                                                    " & vbCrLf & _
                      "p1.CidadePlaca01 as CidadePlaca,                                                     " & vbCrLf & _
                      "p1.EstadoPlaca01 as EstadoPlaca,                                                     " & vbCrLf & _
                      "p1.RNTRCPlaca01 as RNTRCPlaca,                                                       " & vbCrLf & _
                      "p1.NomeMotorista as NomeMotorista,                                                   " & vbCrLf & _
                      "p1.CidadeMotorista as CidadeMotorista,                                               " & vbCrLf & _
                      "p1.EstadoMotorista as EstadoMotorista,                                               " & vbCrLf & _
                      "p1.Habilitacao as Habilitacao,                                                       " & vbCrLf & _
                      "p1.CpfMotorista as CpfMotorista,                                                     " & vbCrLf & _
                      "p1.EndCpfMotorista as EndCpfMotorista,                                               " & vbCrLf & _
                      "p1.Proprietario01 as Proprietario,                                                   " & vbCrLf & _
                      "p1.EndProprietario01 as EndProprietario,                                             " & vbCrLf & _
                      "p1.NomeProprietario01 as NomeProprietario,                                           " & vbCrLf & _
                      "'N' as Principal                                                                     " & vbCrLf & _
                      "FROM Placas p1                                                                       " & vbCrLf & _
                      "        WHERE 1=1                                                                    " & vbCrLf & _
                      "AND p1.Placa_Id = ' ADG-1942'                                                                                     " & vbCrLf & _
                      "AND p1.Placa01 IS NOT NULL                                                           " & vbCrLf & _
                      "AND p1.Placa01 <> ''                                                                 " & vbCrLf & _
                      "UNION ALL                                                                            " & vbCrLf & _
                      "SELECT                                                                               " & vbCrLf & _
                      "	LTRIM(RTRIM(p2.Placa_Id)) as Placa_Id,                                              " & vbCrLf & _
                      "	LTRIM(RTRIM(p2.Placa02)) as Placa,                                                  " & vbCrLf & _
                      "	p2.ViaTransporte_Id as ViaTransporte,                                               " & vbCrLf & _
                      "	p2.TipoVeiculo_Id as TipoVeiculo,                                                   " & vbCrLf & _
                      "	p2.CidadePlaca02 as CidadePlaca,                                                    " & vbCrLf & _
                      "	p2.EstadoPlaca02 as EstadoPlaca,                                                    " & vbCrLf & _
                      "	p2.RNTRCPlaca02 as RNTRCPlaca,                                                      " & vbCrLf & _
                      "	p2.NomeMotorista as NomeMotorista,                                                  " & vbCrLf & _
                      "	p2.CidadeMotorista as CidadeMotorista,                                              " & vbCrLf & _
                      "	p2.EstadoMotorista as EstadoMotorista,                                              " & vbCrLf & _
                      "	p2.Habilitacao as Habilitacao,                                                      " & vbCrLf & _
                      "	p2.CpfMotorista as CpfMotorista,                                                    " & vbCrLf & _
                      "	p2.EndCpfMotorista as EndCpfMotorista,                                              " & vbCrLf & _
                      "	p2.Proprietario02 as Proprietario,                                                  " & vbCrLf & _
                      "	p2.EndProprietario02 as EndProprietario,                                            " & vbCrLf & _
                      "	p2.NomeProprietario02 as NomeProprietario,                                          " & vbCrLf & _
                      "'N' as Principal                                                                     " & vbCrLf & _
                      "FROM Placas p2                                                                       " & vbCrLf & _
                      "WHERE 1=1                                                                            " & vbCrLf & _
                      "AND p2.Placa_Id = ' ADG-1942'                                                                                     " & vbCrLf & _
                      "AND p2.Placa02 IS NOT NULL                                                           " & vbCrLf & _
                      "AND p2.Placa02 <> ''                                                                 " & vbCrLf & _
                      "UNION ALL                                                                            " & vbCrLf & _
                      "SELECT                                                                               " & vbCrLf & _
                      "	LTRIM(RTRIM(p3.Placa_Id)) as Placa_Id,                                              " & vbCrLf & _
                      "	LTRIM(RTRIM(p3.Placa03)) as Placa,                                                  " & vbCrLf & _
                      "	p3.ViaTransporte_Id as ViaTransporte,                                               " & vbCrLf & _
                      "	p3.TipoVeiculo_Id as TipoVeiculo,                                                   " & vbCrLf & _
                      "	p3.CidadePlaca03 as CidadePlaca,                                                    " & vbCrLf & _
                      "	p3.EstadoPlaca03 as EstadoPlaca,                                                    " & vbCrLf & _
                      "	p3.RNTRCPlaca03 as RNTRCPlaca,                                                      " & vbCrLf & _
                      "	p3.NomeMotorista as NomeMotorista,                                                  " & vbCrLf & _
                      "	p3.CidadeMotorista as CidadeMotorista,                                              " & vbCrLf & _
                      "	p3.EstadoMotorista as EstadoMotorista,                                              " & vbCrLf & _
                      "	p3.Habilitacao as Habilitacao,                                                      " & vbCrLf & _
                      "	p3.CpfMotorista as CpfMotorista,                                                    " & vbCrLf & _
                      "	p3.EndCpfMotorista as EndCpfMotorista,                                              " & vbCrLf & _
                      "	p3.Proprietario03 as Proprietario,                                                  " & vbCrLf & _
                      "	p3.EndProprietario03 as EndProprietario,                                            " & vbCrLf & _
                      "	p3.NomeProprietario03 as NomeProprietario,                                          " & vbCrLf & _
                      "'N' as Principal                                                                     " & vbCrLf & _
                      "FROM Placas p3                                                                       " & vbCrLf & _
                      "WHERE 1=1                                                                            " & vbCrLf & _
                     "AND p3.Placa_Id = ' ADG-1942'                                                                                     " & vbCrLf & _
                      "AND p3.Placa03 IS NOT NULL                                                           " & vbCrLf & _
                      "AND p3.Placa03 <> ''                                                                 " & vbCrLf



                ds = Banco.ConsultaDataSet(Sql, "Consulta")

                If ds IsNot Nothing AndAlso ds.Tables IsNot Nothing AndAlso ds.Tables.Count > 0 AndAlso ds.Tables(0).Rows.Count > 0 Then
                    GridPlacas.DataSource = ds
                    GridPlacas.DataBind()
                Else
                    MsgBox(Me.Page, "Nenhum resultado encontrado.")
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

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("Placas", "RELATORIO") Then
                If ValidarCampos() Then
                    Dim ds As New DataSet
                    Dim Sql As String = ""

                    Sql = "SELECT LTRIM(RTRIM(p1.Placa_Id)) as Placa_Id,                                                                      " & vbCrLf & _
                           "	vt.Descricao as ViaTransporte,                                                                                " & vbCrLf & _
                           "	tv.Descricao as TipoVeiculo,                                                                                  " & vbCrLf & _
                           "	p1.CidadePlaca as CidadePlaca,                                                                                " & vbCrLf & _
                           "	p1.EstadoPlaca as EstadoPlaca,                                                                                " & vbCrLf & _
                           "	isnull(p1.RNTRCPlaca,'') as RNTRCPlaca,                                                                       " & vbCrLf & _
                           "	p1.NomeMotorista as NomeMotorista,                                                                            " & vbCrLf & _
                           "	p1.CidadeMotorista as CidadeMotorista,                                                                        " & vbCrLf & _
                           "	p1.EstadoMotorista as EstadoMotorista,                                                                        " & vbCrLf & _
                           "	p1.Habilitacao as Habilitacao,                                                                                " & vbCrLf & _
                           "	p1.CpfMotorista as CpfMotorista,                                                                              " & vbCrLf & _
                           "	p1.EndCpfMotorista as EndCpfMotorista,                                                                        " & vbCrLf & _
                           "    isnull(p1.Proprietario,'') as Proprietario,                                                                   " & vbCrLf & _
                           "    isnull(p1.EndProprietario,'') as EndProprietario,                                                             " & vbCrLf & _
                           "    isnull(p1.NomeProprietario,'') as NomeProprietario,                                                           " & vbCrLf & _
                           "	LTRIM(RTRIM(p1.Placa01)) as Placa,	                                                                          " & vbCrLf & _
                           "	p1.CidadePlaca01 as CidadePlaca,                                                                              " & vbCrLf & _
                           "	p1.EstadoPlaca01 as EstadoPlaca,                                                                              " & vbCrLf & _
                           "	p1.RNTRCPlaca01 as RNTRCPlaca,                                                                                " & vbCrLf & _
                           "	p1.Proprietario01 as Proprietario,                                                                            " & vbCrLf & _
                           "	p1.EndProprietario01 as EndProprietario,                                                                      " & vbCrLf & _
                           "	p1.NomeProprietario01 as NomeProprietario                                                                     " & vbCrLf & _
                           "FROM Placas p1                                                                                                    " & vbCrLf & _
                           "INNER JOIN TiposDeVeiculos tv                                                                                     " & vbCrLf & _
                           "ON tv.Codigo_Id = p1.TipoVeiculo_Id                                                                               " & vbCrLf & _
                           "INNER JOIN ViaDeTransportes vt                                                                                    " & vbCrLf & _
                           "ON vt.Codigo_Id = p1.ViaTransporte_Id                                                                             " & vbCrLf & _
                           "WHERE 1=1                                                                                                         " & vbCrLf

                    If Not String.IsNullOrWhiteSpace(txtPlaca.Text) Then
                        Sql &= "AND p1.Placa_Id = '" & txtPlaca.Text.Trim() & "' " & vbCrLf
                    End If

                    If Not String.IsNullOrWhiteSpace(ddlViaTransporte.SelectedValue) Then
                        Sql &= "AND p1.ViaTransporte_Id = '" & ddlViaTransporte.SelectedValue & "' " & vbCrLf
                    End If

                    If Not String.IsNullOrWhiteSpace(ddlTipoVeiculo.SelectedValue) Then
                        Sql &= "AND p1.TipoVeiculo_Id = '" & ddlTipoVeiculo.SelectedValue & "' " & vbCrLf
                    End If


                    Sql &= "AND p1.Placa01 IS NOT NULL                                                                                                    " & vbCrLf & _
                            "AND p1.Placa01 <> ''                                                                                                         " & vbCrLf & _
                            "UNION ALL                                                                                                                    " & vbCrLf & _
                            "Select                                                                                                                       " & vbCrLf & _
                            "	LTRIM(RTRIM(p2.Placa_Id)) as Placa_Id,                                                                                    " & vbCrLf & _
                            "	vt.Descricao as ViaTransporte,                                                                                            " & vbCrLf & _
                            "	tv.Descricao as TipoVeiculo,                                                                                              " & vbCrLf & _
                            "	p2.CidadePlaca as CidadePlaca,                                                                                            " & vbCrLf & _
                            "	p2.EstadoPlaca as EstadoPlaca,                                                                                            " & vbCrLf & _
                            "	isnull(p2.RNTRCPlaca,'') as RNTRCPlaca,                                                                                   " & vbCrLf & _
                            "	p2.NomeMotorista as NomeMotorista,                                                                                        " & vbCrLf & _
                            "	p2.CidadeMotorista as CidadeMotorista,                                                                                    " & vbCrLf & _
                            "	p2.EstadoMotorista as EstadoMotorista,                                                                                    " & vbCrLf & _
                            "	p2.Habilitacao as Habilitacao,                                                                                            " & vbCrLf & _
                            "	p2.CpfMotorista as CpfMotorista,                                                                                          " & vbCrLf & _
                            "	p2.EndCpfMotorista as EndCpfMotorista,                                                                                    " & vbCrLf & _
                            "	isnull(p2.Proprietario,'') as Proprietario,                                                                               " & vbCrLf & _
                            "	isnull(p2.EndProprietario,'') as EndProprietario,                                                                         " & vbCrLf & _
                            "	isnull(p2.NomeProprietario,'') as NomeProprietario,                                                                       " & vbCrLf & _
                            "	LTRIM(RTRIM(p2.Placa02)) as Placa,                                                                                        " & vbCrLf & _
                            "	p2.CidadePlaca02 as CidadePlaca,                                                                                          " & vbCrLf & _
                            "	p2.EstadoPlaca02 as EstadoPlaca,                                                                                          " & vbCrLf & _
                            "	p2.RNTRCPlaca02 as RNTRCPlaca,                                                                                            " & vbCrLf & _
                            "	p2.Proprietario02 as Proprietario,                                                                                        " & vbCrLf & _
                            "	p2.EndProprietario02 as EndProprietario,                                                                                  " & vbCrLf & _
                            "	p2.NomeProprietario02 as NomeProprietario                                                                                 " & vbCrLf & _
                            "FROM Placas p2                                                                                                               " & vbCrLf & _
                            "INNER JOIN TiposDeVeiculos tv                                                                                                " & vbCrLf & _
                            "ON tv.Codigo_Id = p2.TipoVeiculo_Id                                                                                          " & vbCrLf & _
                            "INNER JOIN ViaDeTransportes vt                                                                                               " & vbCrLf & _
                            "ON vt.Codigo_Id = p2.ViaTransporte_Id                                                                                        " & vbCrLf & _
                            "WHERE 1=1                                                                                                                    " & vbCrLf

                    If Not String.IsNullOrWhiteSpace(txtPlaca.Text) Then
                        Sql &= "AND p2.Placa_Id = '" & txtPlaca.Text.Trim() & "' " & vbCrLf
                    End If

                    If Not String.IsNullOrWhiteSpace(ddlViaTransporte.SelectedValue) Then
                        Sql &= "AND p2.ViaTransporte_Id = '" & ddlViaTransporte.SelectedValue & "' " & vbCrLf
                    End If

                    If Not String.IsNullOrWhiteSpace(ddlTipoVeiculo.SelectedValue) Then
                        Sql &= "AND p2.TipoVeiculo_Id = '" & ddlTipoVeiculo.SelectedValue & "' " & vbCrLf
                    End If

                    Sql &= "AND p2.Placa02 IS NOT NULL                                                                                                    " & vbCrLf & _
                                       "AND p2.Placa02 <> ''                                                                                              " & vbCrLf & _
                                       "UNION ALL                                                                                                         " & vbCrLf & _
                                       "SELECT                                                                                                            " & vbCrLf & _
                                       "	LTRIM(RTRIM(p3.Placa_Id)) as Placa_Id,                                                                        " & vbCrLf & _
                                       "	vt.Descricao as ViaTransporte,                                                                                " & vbCrLf & _
                                       "	tv.Descricao as TipoVeiculo,                                                                                  " & vbCrLf & _
                                       "	p3.CidadePlaca as CidadePlaca,                                                                                " & vbCrLf & _
                                       "	p3.EstadoPlaca as EstadoPlaca,                                                                                " & vbCrLf & _
                                       "	isnull(p3.RNTRCPlaca,'') as RNTRCPlaca,                                                                       " & vbCrLf & _
                                       "	p3.NomeMotorista as NomeMotorista,                                                                            " & vbCrLf & _
                                       "	p3.CidadeMotorista as CidadeMotorista,                                                                        " & vbCrLf & _
                                       "	p3.EstadoMotorista as EstadoMotorista,                                                                        " & vbCrLf & _
                                       "	p3.Habilitacao as Habilitacao,                                                                                " & vbCrLf & _
                                       "	p3.CpfMotorista as CpfMotorista,                                                                              " & vbCrLf & _
                                       "	p3.EndCpfMotorista as EndCpfMotorista,                                                                        " & vbCrLf & _
                                       "	isnull(p3.Proprietario,'') as Proprietario,                                                                   " & vbCrLf & _
                                       "	isnull(p3.EndProprietario,'') as EndProprietario,                                                             " & vbCrLf & _
                                       "	isnull(p3.NomeProprietario,'') as NomeProprietario,                                                           " & vbCrLf & _
                                       "	LTRIM(RTRIM(p3.Placa03)) as Placa,                                                                            " & vbCrLf & _
                                       "	p3.CidadePlaca03 as CidadePlaca,                                                                              " & vbCrLf & _
                                       "	p3.EstadoPlaca03 as EstadoPlaca,                                                                              " & vbCrLf & _
                                       "	p3.RNTRCPlaca03 as RNTRCPlaca,                                                                                " & vbCrLf & _
                                       "	p3.Proprietario03 as Proprietario,                                                                            " & vbCrLf & _
                                       "	p3.EndProprietario03 as EndProprietario,                                                                      " & vbCrLf & _
                                       "	p3.NomeProprietario03 as NomeProprietario                                                                     " & vbCrLf & _
                                       "FROM Placas p3                                                                                                    " & vbCrLf & _
                                       "INNER JOIN TiposDeVeiculos tv                                                                                     " & vbCrLf & _
                                       "ON tv.Codigo_Id = p3.TipoVeiculo_Id                                                                               " & vbCrLf & _
                                       "INNER JOIN ViaDeTransportes vt                                                                                    " & vbCrLf & _
                                       "ON vt.Codigo_Id = p3.ViaTransporte_Id                                                                             " & vbCrLf & _
                                       "WHERE 1=1                                                                                                         " & vbCrLf

                    If Not String.IsNullOrWhiteSpace(txtPlaca.Text) Then
                        Sql &= "AND p3.Placa_Id = '" & txtPlaca.Text.Trim() & "' " & vbCrLf
                    End If

                    If Not String.IsNullOrWhiteSpace(ddlViaTransporte.SelectedValue) Then
                        Sql &= "AND p3.ViaTransporte_Id = '" & ddlViaTransporte.SelectedValue & "' " & vbCrLf
                    End If

                    If Not String.IsNullOrWhiteSpace(ddlTipoVeiculo.SelectedValue) Then
                        Sql &= "AND p3.TipoVeiculo_Id = '" & ddlTipoVeiculo.SelectedValue & "' " & vbCrLf
                    End If

                    Sql &= "AND p3.Placa03 IS NOT NULL                                                                                        " & vbCrLf & _
                           "AND p3.Placa03 <> ''                                                                                              " & vbCrLf

                    ds = Banco.ConsultaDataSet(Sql, "Placas")

                    Dim parameters As New Dictionary(Of String, Object)
                    parameters.Add("ConsultaParametros", getParam())

                    Funcoes.BindReport(Me.Page, ds, "Cr_Placas", eExportType.PDF, parameters)
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
            Funcoes.Ajuda(Me.Page, "Placas")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub

End Class