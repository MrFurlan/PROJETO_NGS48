Imports NGS.Lib.Negocio
Imports System.IO

Public Class RelatorioOrigemDaNotaFiscal
    Inherits BasePage

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Try
            Me.setMenu(eModulo.Expedicao)
            If Not IsPostBack And IsConnect Then
                If Funcoes.VerificaPermissao("RelatorioOrigemDaNotaFiscal", "ACESSAR") Then
                    BuncarUnidadeDeNegocio()
                    HID.Value = Guid.NewGuid().ToString()
                    ucConsultaClientes.SetarHID(HID.Value)
                    txtData1.Text = Format(Today, "dd/MM/yyyy")
                    txtData2.Text = Format(Today, "dd/MM/yyyy")
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

    Protected Sub ddlUnidadeNegocio_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ddlUnidadeNegocio.SelectedIndexChanged
        Try
            BuscarEmpresas()
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Protected Sub btnCliente_Click(sender As Object, e As EventArgs) Handles btnCliente.Click
        Try
            ucConsultaClientes.Limpar()
            Popup.ConsultaDeClientes(Me, "objClienteNota" & HID.Value.ToString, "txtNome")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Sub BuncarUnidadeDeNegocio()
        ddl.Carregar(ddlUnidadeNegocio, CarregarDDL.Tabela.UnidadeDeNegocio, "", True)
        Funcoes.VerificaUnidadeDeNegocio(ddlUnidadeNegocio, ddlEmpresa, True)

    End Sub

    Private Sub BuscarEmpresas()
        ddl.Carregar(ddlEmpresa, CarregarDDL.Tabela.Empresas, ddlUnidadeNegocio.SelectedValue, True)
    End Sub

    Private Sub Limpar()
        ddlUnidadeNegocio.SelectedIndex = 0
        ddlEmpresa.Items.Clear()
        txtCliente.Text = String.Empty
        txtCodigoCliente.Value = String.Empty
        txtNota.Text = String.Empty
        txtSerie.Text = String.Empty
        'txtData1.Text = String.Empty
        'txtData2.Text = String.Empty
        txtData1.Text = Format(Today, "dd/MM/yyyy")
        txtData2.Text = Format(Today, "dd/MM/yyyy")
    End Sub

    Private Sub LiberaEmpresa()
        If Not UsuarioServidor.LiberaEmpresa Then
            ddlUnidadeNegocio.Enabled = False
            ddlEmpresa.Enabled = False
        End If
    End Sub

    Public Overrides Sub Carregar(obj As IBaseEntity)
        If Not Session("objClienteNota" & HID.Value) Is Nothing Then
            Dim itemCliente As ListItem = Funcoes.FormatarListItemCliente(CType(Session("objClienteNota" & HID.Value), [Lib].Negocio.Cliente))
            txtCliente.Text = itemCliente.Text
            txtCodigoCliente.Value = itemCliente.Value
            Session.Remove("objClienteNota" & HID.Value)
        End If
    End Sub

    Protected Sub lnkRelatorio_Click(sender As Object, e As EventArgs) Handles lnkRelatorio.Click
        Try
            If Funcoes.VerificaPermissao("RelatorioOrigemDaNotaFiscal", "RELATORIO") Then
                If String.IsNullOrWhiteSpace(txtNota.Text) Then
                    MsgBox(Me.Page, "Informe a nota.")
                    Exit Sub
                End If

                Dim parametros As String = ""
                Dim ds As DataSet = getDataSet(parametros)
                Dim param As New Dictionary(Of String, Object)
                param.Add("parametros", parametros)

                Funcoes.BindReport(Me.Page, ds, "Cr_OrigemDaNotaFiscal", eExportType.PDF, param)
            Else
                MsgBox(Me.Page, "Usuário sem permissão para emitir relatório.", eTitulo.Info)
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

    Private Function getDataSet(ByRef parametros As String) As DataSet
        Dim ds As New DataSet()
        Dim sql As String = "   SELECT                                                                  " & vbCrLf & _
                            "   	CAST(NF.Nota_Id as varchar) + '-' + CAST(NF.Serie_Id as varchar) as Nota,               " & vbCrLf & _
                            "   	NF.Empresa_Id + '-' + CAST(NF.EndEmpresa_Id as varchar) as Empresa,                     " & vbCrLf & _
                            "   	NF.Cliente_Id + '-' + CAST(NF.EndCliente_Id as varchar) + '-' + CLI.Nome as Cliente,    " & vbCrLf & _
                            "   	NF.EntradaSaida_Id as EntradaSaida,                                                     " & vbCrLf & _
                            "   	CAST(NF.Operacao as varchar) + '-' + CAST(NF.SubOperacao as varchar) as Operacao,       " & vbCrLf & _
                            "   	NF.Movimento,                                                                           " & vbCrLf & _
                            "   	CASE WHEN NF.Operacao = 80 THEN 'CTRC'                                                  " & vbCrLf & _
                            "   		 WHEN NF.Operacao <> 80 AND NF.TipoDeDocumento = 2 THEN 'NFG'                       " & vbCrLf & _
                            "   		 ELSE 'NF' END as Tipo,                                                             " & vbCrLf & _
                            "                   NXI.Quantidade,                                                             " & vbCrLf & _
                            "                   NXI.Valor                                                                   " & vbCrLf & _
                            "   FROM NotasFiscais NF                                                                        " & vbCrLf & _
                            "   INNER JOIN (                                                                                " & vbCrLf & _
                            "   	SELECT                                                                                  " & vbCrLf & _
                            "   		NXI.Empresa_Id,                                                                     " & vbCrLf & _
                            "   		NXI.EndEmpresa_Id,                                                                  " & vbCrLf & _
                            "   		NXI.Cliente_Id,                                                                     " & vbCrLf & _
                            "   		NXI.EndCliente_Id,                                                                  " & vbCrLf & _
                            "   		NXI.EntradaSaida_Id,                                                                " & vbCrLf & _
                            "   		NXI.Nota_Id,                                                                        " & vbCrLf & _
                            "   		NXI.Serie_Id,                                                                       " & vbCrLf & _
                            "   		SUM(NXI.QuantidadeFiscal) AS Quantidade,                                            " & vbCrLf & _
                            "   		SUM(NXI.Valor) AS Valor                                                             " & vbCrLf & _
                            "   	FROM NotasFiscaisXItens NXI                                                             " & vbCrLf & _
                            "   	GROUP BY NXI.Empresa_Id,                                                                " & vbCrLf & _
                            "   			 NXI.EndEmpresa_Id,                                                             " & vbCrLf & _
                            "   			 NXI.Cliente_Id,                                                                " & vbCrLf & _
                            "   			 NXI.EndCliente_Id,                                                             " & vbCrLf & _
                            "   			 NXI.EntradaSaida_Id,                                                           " & vbCrLf & _
                            "   			 NXI.Serie_Id,                                                                  " & vbCrLf & _
                            "   			 NXI.Nota_Id                                                                    " & vbCrLf & _
                            "   ) AS NXI ON NXI.Empresa_Id    = NF.Empresa_Id                                               " & vbCrLf & _
                            "   	AND NXI.EndEmpresa_Id   = NF.EndEmpresa_Id                                              " & vbCrLf & _
                            "   	AND NXI.Cliente_Id      = NF.Cliente_Id                                                 " & vbCrLf & _
                            "   	AND NXI.EndCliente_Id   = NF.EndCliente_Id                                              " & vbCrLf & _
                            "   	AND NXI.EntradaSaida_Id = NF.EntradaSaida_Id                                            " & vbCrLf & _
                            "   	AND NXI.Nota_Id         = NF.Nota_Id                                                    " & vbCrLf & _
                            "   	AND NXI.Serie_Id        = NF.Serie_Id                                                   " & vbCrLf & _
                            "   INNER JOIN Clientes CLI                                                                     " & vbCrLf & _
                            "   	ON	CLI.Cliente_Id	= NF.Cliente_Id                                                     " & vbCrLf & _
                            "   	AND CLI.Endereco_Id = NF.EndCliente_Id                                                  " & vbCrLf & _
                            "   WHERE NF.Nota_Id = " & txtNota.Text & vbCrLf

        If Not String.IsNullOrWhiteSpace(txtNota.Text) Then
            parametros = "Nota: " & txtNota.Text & " - "
        End If
        If Not String.IsNullOrWhiteSpace(txtSerie.Text) Then
            sql &= "   AND NF.Serie_Id = " & txtSerie.Text & vbCrLf
            parametros &= "Série: " & txtSerie.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(ddlEmpresa.SelectedValue) Then
            sql &= "   AND NF.Empresa_Id = '" & ddlEmpresa.SelectedValue.Split("-")(0) & "'" & vbCrLf & _
        "   AND NF.EndEmpresa_Id = " & ddlEmpresa.SelectedValue.Split("-")(1) & vbCrLf
            Dim objEmpresa = New Cliente(ddlEmpresa.SelectedValue.Split("-")(0), ddlEmpresa.SelectedValue.Split("-")(1))
            parametros &= "Empresa: " & objEmpresa.Reduzido & " - " & objEmpresa.Nome & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtCodigoCliente.Value) Then
            sql &= "   AND NF.Cliente_Id = '" & txtCodigoCliente.Value.Split("-")(0) & "'" & vbCrLf & _
           "   AND NF.EndCliente_Id = " & txtCodigoCliente.Value.Split("-")(1) & vbCrLf
            parametros &= "Cliente: " & txtCliente.Text & vbCrLf
        End If
        If Not String.IsNullOrWhiteSpace(txtData1.Text) AndAlso Not String.IsNullOrWhiteSpace(txtData2.Text) Then
            sql &= "   AND NF.Movimento BETWEEN '" & CDate(txtData1.Text).ToString("yyyy-MM-dd") & "' AND '" & CDate(txtData2.Text).ToString("yyyy-MM-dd") & "'" & vbCrLf
            parametros &= "Movimento de: " & txtData1.Text & " á " & txtData2.Text & vbCrLf
        End If
        

        sql &= "   ORDER BY NF.Movimento                                                                       " & vbCrLf
        ds = Banco.ConsultaDataSet(sql, "OrigemDaNotaFiscal")

        For Each row As DataRow In ds.Tables(0).Rows
            row("Empresa") = Funcoes.FormatarCpfCnpj(row("Empresa").ToString().Split("-")(0)) & " - " & row("Empresa").ToString().Split("-")(1)

            Dim cnpj As String = Funcoes.FormatarCpfCnpj(row("Cliente").ToString().Split("-")(0))
            Dim nome As String = row("Cliente").ToString().Substring(row("Cliente").ToString().IndexOf("-"), row("Cliente").ToString().Length - row("Cliente").ToString().IndexOf("-"))
            row("Cliente") = String.Format("{0} {1}", cnpj, nome)
        Next

        Return ds
    End Function

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "RelatorioOrigemDaNotaFiscal")
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub
End Class