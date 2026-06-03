Imports System.Data
Imports NGS.Lib.Uteis
Imports NGS.Lib.Negocio

Public Class ConsultaNotaXNota
    Inherits BasePage

    Private Sql As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Me.setMenu(eModulo.Expedicao)
        If Not IsPostBack And IsConnect Then
            If Not Funcoes.VerificaPermissao("ConsultaNotaXNota", "ACESSAR") Then
                MsgBox(Me.Page, "Usuário sem permissão para acessar essa página.", "~/Expedicao.aspx", eTitulo.Info)
                Exit Sub
            End If
        End If
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs) Handles lnkConsultar.Click
        If String.IsNullOrWhiteSpace(txtNotaFiscal.Text) Then
            MsgBox(Me.Page, "Informe o número da Nota Fiscal.")
        Else
            Sql = "SELECT (NxN.Empresa_Id + '-' + convert(varchar,NxN.EndEmpresa_Id)) AS Empresa, (NxN.Cliente_Id + '-' + " & vbCrLf & _
                    "		convert(varchar,NxN.EndCliente_Id) + '  ' + CNF.Nome) AS Cliente, " & vbCrLf & _
                    "		(NxN.EntradaSaida_Id + '-' + convert(varchar,NxN.Nota_Id) + '-' + NxN.Serie_Id) AS Nota, " & vbCrLf & _
                    "		NF.Movimento, TdNF.Descricao, " & vbCrLf & _
                    "		(NxN.OrigemEmpresa_Id + '-' + convert(varchar,NxN.OrigemEndEmpresa_Id)) AS OrigemEmpresa, " & vbCrLf & _
                    "		(NxN.OrigemCliente_Id + '-' + convert(varchar,NxN.OrigemEndCliente_Id) + '  ' + CNFO.Nome) AS OrigemCliente, " & vbCrLf & _
                    "		(NxN.OrigemEntradaSaida_Id + '-' + convert(varchar,NxN.OrigemNota_Id) + '-' + NxN.OrigemSerie_Id) AS OrigemNota, " & vbCrLf & _
                    "		NFO.Movimento AS OrigemMovimento, TdNFO.Descricao AS OrigemDescricao " & vbCrLf & _
                    "FROM  NotasXNotas AS NxN " & vbCrLf & _
                    "	INNER JOIN NotasFiscais NF " & vbCrLf & _
                    "			 ON NF.Empresa_Id      = NxN.Empresa_Id  " & vbCrLf & _
                    "			AND NF.EndEmpresa_Id   = NxN.EndEmpresa_Id  " & vbCrLf & _
                    "			AND NF.Cliente_Id      = NxN.Cliente_Id  " & vbCrLf & _
                    "			AND NF.EndCliente_Id   = NxN.EndCliente_Id  " & vbCrLf & _
                    "			AND NF.EntradaSaida_Id = NxN.EntradaSaida_Id  " & vbCrLf & _
                    "			AND NF.Serie_Id        = NxN.Serie_Id  " & vbCrLf & _
                    "			AND NF.Nota_Id         = NxN.Nota_Id  " & vbCrLf & _
                    "	INNER JOIN TipoDeDocumento TdNF " & vbCrLf & _
                    "			 ON TdNF.Codigo_Id     = NF.TipoDeDocumento  " & vbCrLf & _
                    "	INNER JOIN Clientes CNF " & vbCrLf & _
                    "			 ON CNF.Cliente_Id      = NxN.Cliente_Id  " & vbCrLf & _
                    "			AND CNF.Endereco_Id     = NxN.EndCliente_Id  " & vbCrLf & _
                    "	INNER JOIN NotasFiscais NFO " & vbCrLf & _
                    "			 ON NFO.Empresa_Id      = NxN.OrigemEmpresa_Id  " & vbCrLf & _
                    "			AND NFO.EndEmpresa_Id   = NxN.OrigemEndEmpresa_Id  " & vbCrLf & _
                    "			AND NFO.Cliente_Id      = NxN.OrigemCliente_Id  " & vbCrLf & _
                    "			AND NFO.EndCliente_Id   = NxN.OrigemEndCliente_Id  " & vbCrLf & _
                    "			AND NFO.EntradaSaida_Id = NxN.OrigemEntradaSaida_Id  " & vbCrLf & _
                    "			AND NFO.Serie_Id        = NxN.OrigemSerie_Id  " & vbCrLf & _
                    "			AND NFO.Nota_Id         = NxN.OrigemNota_Id  " & vbCrLf & _
                    "	INNER JOIN TipoDeDocumento TdNFO " & vbCrLf & _
                    "			 ON TdNFO.Codigo_Id     = NFO.TipoDeDocumento  " & vbCrLf & _
                    "	INNER JOIN Clientes CNFO " & vbCrLf & _
                    "			 ON CNFO.Cliente_Id     = NxN.OrigemCliente_Id  " & vbCrLf & _
                    "			AND CNFO.Endereco_Id    = NxN.OrigemEndCliente_Id  " & vbCrLf & _
                    " WHERE (NxN.Nota_Id = " & txtNotaFiscal.Text & ") OR (NxN.OrigemNota_Id = " & txtNotaFiscal.Text & ") " & vbCrLf & _
                    " ORDER BY NF.Movimento DESC " & vbCrLf

            Dim ds As New DataSet
            ds = Banco.ConsultaDataSet(Sql, "NotasXNotas")
            If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
                MsgBox(Me.Page, "Notas X Notas não encontrado!")
            Else
                gridNxN.DataSource = ds
                gridNxN.DataBind()
            End If
        End If
    End Sub

    Protected Sub lnkLimpar_Click(sender As Object, e As EventArgs) Handles lnkLimpar.Click
        txtNotaFiscal.Text = ""
        gridNxN.DataSource = Nothing
        gridNxN.DataBind()
    End Sub

    Protected Sub lnkAjuda_Click(sender As Object, e As EventArgs) Handles lnkAjuda.Click
        Try
            Funcoes.Ajuda(Me.Page, "ConsultaNotaXNota")
        Catch ex As Exception
            MsgBox(Me.Page, "Não foi possível exibir a ajuda.", eTitulo.Info)
        End Try
    End Sub
End Class