Imports System.Data
Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class ucConsultaPedidosXNotas
    Inherits BaseUserControl

    Dim objNotaFiscal As NotaFiscal

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        If IsPostBack Then
            Return
        End If
    End Sub

    Public Overrides Sub SetarHID(ByVal guid As String)
        HID.Value = guid.ToString
    End Sub

    Public Function BindGridView() As Integer
        If Session("ssCampo" & HID.Value) = "NotaXItens" Then
            Return CargaNotasXItens()
        ElseIf Session("ssCampo" & HID.Value) = "AlterarLaudo" Then
            Return CargaNotasXItens()
        ElseIf Session("ssCampo" & HID.Value) = "AlterarEncargos" Then
            Return CargaNotasXItens()
        ElseIf Session("ssCampo" & HID.Value) = "AlterarEncargosCTE" Then
            Return CargaNotasXItens()
        ElseIf Session("ssCampo" & HID.Value) = "RateioDePesagem" Then
            Return CargaNotasXItens()
        ElseIf Session("ssCampo" & HID.Value) = "NotaFiscalGeral" Then
            Return CargaNotasGerais()
        ElseIf Session("ssCampo" & HID.Value) = "NotasDeTerceiro" Then
            Return CargaNotasDeTerceiro()
        ElseIf Session("ssCampo" & HID.Value) = "AlterarNotaGeral" Then
            Return CargaNotasGeraisAlteracao()
        Else
            CargaNotas()
        End If
    End Function

    Private Sub SessaoRecuperaNotaFiscal()
        objNotaFiscal = CType(Session("objNotaFiscal" & HID.Value.ToString), [Lib].Negocio.NotaFiscal)
    End Sub

    Private Function CargaNotasXItens() As Integer
        SessaoRecuperaNotaFiscal()
        Dim Sql As String
        Sql = " SELECT DISTINCT NF.Movimento," & vbCrLf &
              "	       NF.Empresa_id," & vbCrLf &
              "	       NF.EndEmpresa_id," & vbCrLf &
              "	       NF.Cliente_Id AS Cliente," & vbCrLf &
              "	       NF.EndCliente_Id AS EndCliente," & vbCrLf &
              "        C.Nome as NomeCliente," & vbCrLf &
              "	       NF.EntradaSaida_Id AS ES," & vbCrLf &
              "	       NF.Serie_Id AS Serie," & vbCrLf &
              "	       NF.Nota_Id AS Nota," & vbCrLf &
              "	       NF.Operacao," & vbCrLf &
              "	       NF.SubOperacao," & vbCrLf &
              "        case " & vbCrLf &
              "         when isnull(NFR.Nota_Id,0) > 0 " & vbCrLf &
              "             then NF.NossaEmissao + '-' + NFR.Retorno + '-' + NFR.MsgRetorno " & vbCrLf &
              "             else " & vbCrLf &
              "                 case " & vbCrLf &
              "                     when isnull(NFP.Nota_Id,0) > 0 " & vbCrLf &
              "                         then NF.NossaEmissao + '-' + NFP.Retorno + '-' + NFP.MsgRetorno " & vbCrLf &
              "                         else NF.NossaEmissao + '-' " & vbCrLf &
              "                     end " & vbCrLf &
              "         end AS NossaEmissao, " & vbCrLf &
              "	       NFxI.QuantidadeFiscal AS Quantidade," & vbCrLf &
              "	       NFxI.Unitario," & vbCrLf &
              "	       NFxI.Valor," & vbCrLf &
              "	       isnull(NFxR.Romaneio_Id,0) AS Romaneio," & vbCrLf &
              "        case " & vbCrLf &
              "         when isnull(NFR.Retorno,0) > 0 and isnull(NFR.Retorno,0) <> 100 and isnull(NFR.Retorno,0) <> 101 and isnull(NFR.Retorno,0) <> 110 and isnull(NFR.Retorno,0) <> 302" & vbCrLf &
              "             then Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao + '(' + NFR.Retorno + '-' + NFR.MsgRetorno + ') Verificar'" & vbCrLf &
              "             else " & vbCrLf &
              "                 case " & vbCrLf &
              "                     when isnull(NFP.Retorno,0) > 0 " & vbCrLf &
              "                         then Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao + '(' + NFP.Retorno + '-' + NFP.MsgRetorno + ') Verificar' " & vbCrLf &
              "                         else Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao " & vbCrLf &
              "                 end " & vbCrLf &
              "         end as DescSituacao" & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              " INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf &
              "	   ON NFxI.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf &
              "   AND NFxI.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf &
              "   AND NFxI.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf &
              "   AND NFxI.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf &
              "   AND NFxI.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf &
              "   AND NFxI.Serie_Id        = NFxR.Serie_Id" & vbCrLf &
              "   AND NFxI.Nota_Id         = NFxR.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfeRealizadas NFR" & vbCrLf &
              "	   ON NFR.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND NFR.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND NFR.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND NFR.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND NFR.Nota_Id         = NF.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfePendencias NFP" & vbCrLf &
              "	   ON NFP.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND NFP.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND NFP.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND NFP.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND NFP.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND NFP.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND NFP.Nota_Id         = NF.Nota_Id" & vbCrLf &
              " Inner Join Situacoes S" & vbCrLf &
              "    on S.Situacao_id = NF.Situacao" & vbCrLf &
              " INNER JOIN Clientes C" & vbCrLf &
              "    ON C.Cliente_Id  = NF.Cliente_Id" & vbCrLf &
              "   AND C.Endereco_id = NF.EndCliente_Id" & vbCrLf &
              " WHERE 1 = 1 " & vbCrLf

        If Session("ssCampo" & HID.Value) = "NotaXItens" AndAlso (Not TypeOf Me.Page Is AlterarMovimentoNotaFiscal And
                                                                  Not TypeOf Me.Page Is AlterarChaveNFE And
                                                                  Not TypeOf Me.Page Is CCeObservacao) Then Sql &= "   AND ISNULL(NF.NFG,0) = 0 " & vbCrLf

        If TypeOf Me.Page Is AlterarChaveNFE Then
            Sql &= "   AND NF.TipoDeDocumento IN(1,2,15,57,58) " & vbCrLf
        ElseIf TypeOf Me.Page Is CCeObservacao Then
            Sql &= "   AND NF.TipoDeDocumento IN(1,2,57) " & vbCrLf
        ElseIf TypeOf Me.Page Is AlterarLaudo Then
            Sql &= "   AND NF.TipoDeDocumento IN(1) " & vbCrLf
        ElseIf TypeOf Me.Page Is RateioDePesagem Then
            Sql &= "   AND NF.TipoDeDocumento IN(1) " & vbCrLf
        ElseIf TypeOf Me.Page Is CCeTransporte Then
            Sql &= "   AND NF.TipoDeDocumento IN(1,12) " & vbCrLf
        Else

            If TypeOf Me.Page Is AlterarEncargoCTE Then
                Sql &= "   AND NF.TipoDeDocumento IN(57) " & vbCrLf
            ElseIf Not TypeOf Me.Page Is AlterarEncargo AndAlso Not TypeOf Me.Page Is AlterarMovimentoNotaFiscal Then
                Sql &= "   AND NF.TipoDeDocumento IN(1,11) " & vbCrLf
            Else
                If objNotaFiscal.CodigoTipoDeDocumento > 0 Then Sql &= " AND NOT NF.TipoDeDocumento IN(57)  And NF.TipoDeDocumento = " & objNotaFiscal.CodigoTipoDeDocumento & vbCrLf
            End If

        End If

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa) Then
            Sql &= "   AND NF.Empresa_id       = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                   "   AND NF.EndEmpresa_id    = 0" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            Sql &= "   AND NF.Cliente_Id       ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                   "   AND NF.EndCliente_Id    = " & objNotaFiscal.EnderecoCliente & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.Serie) Then Sql &= "   AND NF.Serie_Id  = '" & objNotaFiscal.Serie & "'" & vbCrLf

        If objNotaFiscal.Codigo > 0 Then
            Sql &= "   AND NF.Nota_id = " & objNotaFiscal.Codigo & vbCrLf
        Else
            Sql &= "   AND (NF.Movimento between '" & objNotaFiscal.DataNota.ToSqlDate() & "' and '" & objNotaFiscal.Movimento.ToSqlDate() & "')" & vbCrLf
        End If

        If objNotaFiscal.CodigoPedido > 0 Then Sql &= "   And NF.Pedido = " & objNotaFiscal.CodigoPedido & vbCrLf

        If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.CodigoSubOperacao > 0 Then
            Sql &= "   AND NF.Operacao         = " & objNotaFiscal.CodigoOperacao & vbCrLf &
                   "   AND NF.SubOperacao      = " & objNotaFiscal.CodigoSubOperacao & vbCrLf
        End If

        If objNotaFiscal.CodigoSituacao > 0 Then Sql &= "   And NF.Situacao = " & objNotaFiscal.CodigoSituacao & vbCrLf


        If objNotaFiscal.NossaEmissao Then Sql &= "   And NF.NossaEmissao = 'S'" & vbCrLf

        Sql &= " ORDER BY NF.Movimento DESC, ES, Serie, Nota" & vbCrLf
        GridNotas.DataSource = Banco.ConsultaDataSet(Sql, "Notas").Tables(0)
        GridNotas.DataBind()
        Return GridNotas.Rows.Count
    End Function

    Public Function SQLAlterarNotasGerais() As String

        SessaoRecuperaNotaFiscal()

        Dim Sql As String

        Sql = " SELECT DISTINCT NF.Movimento," & vbCrLf &
              "	       NF.Empresa_id," & vbCrLf &
              "	       NF.EndEmpresa_id," & vbCrLf &
              "	       NF.Cliente_Id AS Cliente," & vbCrLf &
              "	       NF.EndCliente_Id AS EndCliente," & vbCrLf &
              "        C.Nome as NomeCliente," & vbCrLf &
              "	       NF.EntradaSaida_Id AS ES," & vbCrLf &
              "	       NF.Serie_Id AS Serie," & vbCrLf &
              "	       NF.Nota_Id AS Nota," & vbCrLf &
              "	       NF.Operacao," & vbCrLf &
              "	       NF.SubOperacao," & vbCrLf &
              "        case " & vbCrLf &
              "         when isnull(NFR.Nota_Id,0) > 0 " & vbCrLf &
              "             then NF.NossaEmissao + '-' + NFR.Retorno + '-' + NFR.MsgRetorno " & vbCrLf &
              "             else " & vbCrLf &
              "                 case " & vbCrLf &
              "                     when isnull(NFP.Nota_Id,0) > 0 " & vbCrLf &
              "                         then NF.NossaEmissao + '-' + NFP.Retorno + '-' + NFP.MsgRetorno " & vbCrLf &
              "                         else NF.NossaEmissao + '-' " & vbCrLf &
              "                     end " & vbCrLf &
              "         end AS NossaEmissao, " & vbCrLf &
              "	       NFxI.QuantidadeFiscal AS Quantidade," & vbCrLf &
              "	       NFxI.Unitario," & vbCrLf &
              "	       NFxI.Valor," & vbCrLf &
              "	       isnull(NFxR.Romaneio_Id,0) AS Romaneio," & vbCrLf &
              "        case " & vbCrLf &
              "         when isnull(NFR.Retorno,0) > 0 and isnull(NFR.Retorno,0) <> 100 and isnull(NFR.Retorno,0) <> 101 and isnull(NFR.Retorno,0) <> 110 and isnull(NFR.Retorno,0) <> 302" & vbCrLf &
              "             then Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao + '(' + NFR.Retorno + '-' + NFR.MsgRetorno + ') Verificar'" & vbCrLf &
              "             else " & vbCrLf &
              "                 case " & vbCrLf &
              "                     when isnull(NFP.Retorno,0) > 0 " & vbCrLf &
              "                         then Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao + '(' + NFP.Retorno + '-' + NFP.MsgRetorno + ') Verificar' " & vbCrLf &
              "                         else Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao " & vbCrLf &
              "                 end " & vbCrLf &
              "         end as DescSituacao" & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              " INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
              "    ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
              "   AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
              "   AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
              "   AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
              "   AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
              "   AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
              "   AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf &
              "	   ON NFxI.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf &
              "   AND NFxI.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf &
              "   AND NFxI.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf &
              "   AND NFxI.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf &
              "   AND NFxI.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf &
              "   AND NFxI.Serie_Id        = NFxR.Serie_Id" & vbCrLf &
              "   AND NFxI.Nota_Id         = NFxR.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfeRealizadas NFR" & vbCrLf &
              "	   ON NFR.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND NFR.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND NFR.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND NFR.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND NFR.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND NFR.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND NFR.Nota_Id         = NF.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfePendencias NFP" & vbCrLf &
              "	   ON NFP.Empresa_Id      = NF.Empresa_Id" & vbCrLf &
              "   AND NFP.EndEmpresa_Id   = NF.EndEmpresa_Id" & vbCrLf &
              "   AND NFP.Cliente_Id      = NF.Cliente_Id" & vbCrLf &
              "   AND NFP.EndCliente_Id   = NF.EndCliente_Id" & vbCrLf &
              "   AND NFP.EntradaSaida_Id = NF.EntradaSaida_Id" & vbCrLf &
              "   AND NFP.Serie_Id        = NF.Serie_Id" & vbCrLf &
              "   AND NFP.Nota_Id         = NF.Nota_Id" & vbCrLf &
              " Inner Join Situacoes S" & vbCrLf &
              "     ON S.Situacao_id = NF.Situacao" & vbCrLf &
              " INNER JOIN Clientes C" & vbCrLf &
              "     ON C.Cliente_Id  = NF.Cliente_Id" & vbCrLf &
              "     AND C.Endereco_id = NF.EndCliente_Id" & vbCrLf &
              " WHERE 1 = 1 " & vbCrLf &
              "      " & vbCrLf

        Sql &= "    AND NF.TipoDeDocumento IN(1,11) " & vbCrLf

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoEmpresa) Then
            Sql &= "   AND NF.Empresa_id       = '" & objNotaFiscal.CodigoEmpresa & "'" & vbCrLf &
                   "   AND NF.EndEmpresa_id    = 0" & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.CodigoCliente) Then
            Sql &= "   AND NF.Cliente_Id       ='" & objNotaFiscal.CodigoCliente & "'" & vbCrLf &
                   "   AND NF.EndCliente_Id    = " & objNotaFiscal.EnderecoCliente & vbCrLf
        End If

        If Not String.IsNullOrWhiteSpace(objNotaFiscal.Serie) Then Sql &= "   AND NF.Serie_Id  = '" & objNotaFiscal.Serie & "'" & vbCrLf

        If objNotaFiscal.Codigo > 0 Then
            Sql &= "   AND NF.Nota_id = " & objNotaFiscal.Codigo & vbCrLf
        Else
            Sql &= "   AND (NF.Movimento between '" & objNotaFiscal.DataNota.ToSqlDate() & "' and '" & objNotaFiscal.Movimento.ToSqlDate() & "')" & vbCrLf
        End If

        If objNotaFiscal.CodigoPedido > 0 Then Sql &= "   And NF.Pedido = " & objNotaFiscal.CodigoPedido & vbCrLf

        If objNotaFiscal.CodigoOperacao > 0 AndAlso objNotaFiscal.CodigoSubOperacao > 0 Then
            Sql &= "   AND NF.Operacao         = " & objNotaFiscal.CodigoOperacao & vbCrLf &
                   "   AND NF.SubOperacao      = " & objNotaFiscal.CodigoSubOperacao & vbCrLf
        End If

        If objNotaFiscal.CodigoSituacao > 0 Then Sql &= "   And NF.Situacao = " & objNotaFiscal.CodigoSituacao & vbCrLf

        Sql &= "   AND NF.NossaEmissao = 'N'" & vbCrLf

        Sql &= " ORDER BY NF.Movimento DESC, ES, Serie, Nota" & vbCrLf

        Return Sql

    End Function

    Private Function CargaNotasGeraisAlteracao() As Integer

        Dim Sql As String

        Sql = SQLAlterarNotasGerais()
        GridNotas.DataSource = Banco.ConsultaDataSet(Sql, "Notas").Tables(0)
        GridNotas.DataBind()
        Return GridNotas.Rows.Count

    End Function

    Private Function CargaNotasGerais() As Integer
        Dim strSQL As String
        strSQL = " SELECT DISTINCT TOP 50" & vbCrLf &
                 "        NF.Movimento," & vbCrLf &
                 "		  NF.Empresa_Id," & vbCrLf &
                 "		  NF.EndEmpresa_Id," & vbCrLf &
                 "		  NF.Cliente_Id AS Cliente," & vbCrLf &
                 "		  NF.EndCliente_Id AS EndCliente," & vbCrLf &
                 "        C.Nome as NomeCliente," & vbCrLf &
                 "		  NF.EntradaSaida_Id AS ES," & vbCrLf &
                 "		  NF.Serie_Id AS Serie," & vbCrLf &
                 "		  NF.Nota_Id AS Nota," & vbCrLf &
                 "		  NFxI.Operacao," & vbCrLf &
                 "		  NFxI.SubOperacao," & vbCrLf &
                 "		  NF.NossaEmissao," & vbCrLf &
                 "		  NFxI.QuantidadeFiscal AS Quantidade," & vbCrLf &
                 "		  NFxI.Unitario," & vbCrLf &
                 "		  NFxI.Valor," & vbCrLf &
                 "		  NFxR.Romaneio_Id AS Romaneio," & vbCrLf &
                 "        Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao as DescSituacao" & vbCrLf &
                 "   FROM NotasFiscais NF" & vbCrLf &
                 "  INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                 "	   ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                 "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                 "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                 "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                 "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                 "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                 "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                 "   LEFT OUTER JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf &
                 "	   ON NFxI.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf &
                 "    AND NFxI.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf &
                 "    AND NFxI.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf &
                 "    AND NFxI.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf &
                 "    AND NFxI.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf &
                 "    AND NFxI.Serie_Id        = NFxR.Serie_Id" & vbCrLf &
                 "    AND NFxI.Nota_Id         = NFxR.Nota_Id" & vbCrLf &
                 " Inner Join Situacoes S" & vbCrLf &
                 "    on S.Situacao_id = NF.Situacao" & vbCrLf &
                 "  INNER JOIN Clientes E " & vbCrLf &
                 "     ON E.Cliente_Id = NF.Empresa_Id " & vbCrLf &
                 "    AND E.Endereco_Id = NF.EndEmpresa_Id " & vbCrLf &
                 "  INNER JOIN Clientes C " & vbCrLf &
                 "     ON C.Cliente_Id = NF.Cliente_Id " & vbCrLf &
                 "    AND C.Endereco_Id = NF.EndCliente_Id " & vbCrLf &
                 "  INNER JOIN SubOperacoes SO" & vbCrLf &
                 "     ON NF.Operacao = SO.Operacao_Id" & vbCrLf &
                 "    AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                 "  WHERE (NF.NFG = 1 or NF.NFG is null) " & vbCrLf &
                 "    AND (isnull(NF.NossaEmissao,'N') = 'N' or SO.classe in('" & eClassesOperacoes.SERVICOS.ToString & "','" & eClassesOperacoes.DOACAO.ToString & "')) " & vbCrLf &
                 "    AND (EXISTS (SELECT 1 " & vbCrLf &
                 "                 FROM NotasFiscaisXItens NFI " & vbCrLf &
                 "                INNER JOIN Produtos P " & vbCrLf &
                 "                   ON P.Produto_Id = NFI.Produto_Id " & vbCrLf &
                 "                  AND (P.Agrupar = 'S' OR P.Grupo = '10101' OR P.Grupo = '10102' OR P.Grupo = '10103' OR P.Grupo = '10105' OR P.Grupo = '10106' OR P.Grupo = '10201' OR P.Grupo = '10401' OR P.Grupo = '20101' OR P.Grupo = '30101' OR P.Grupo = '10112' OR P.Grupo = '10111' OR P.Grupo = '10110' OR P.Grupo = '19101' OR P.Grupo = '40201' OR P.Grupo = '40101' OR P.Grupo = '90401')" & vbCrLf &
                 "                WHERE NFI.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
                 "                  AND NFI.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
                 "                  AND NFI.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
                 "                  AND NFI.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
                 "                  AND NFI.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                 "                  AND NFI.Serie_Id        = NF.Serie_Id " & vbCrLf &
                 "                  AND NFI.Nota_Id         = NF.Nota_Id) " & vbCrLf &
                 "        OR " & vbCrLf &
                 "         EXISTS (SELECT 1 " & vbCrLf &
                 "                 FROM NotasFiscais NFP " & vbCrLf &
                 "                WHERE NFP.Empresa_Id      = NF.Empresa_Id " & vbCrLf &
                 "                  AND NFP.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf &
                 "                  AND NFP.Cliente_Id      = NF.Cliente_Id " & vbCrLf &
                 "                  AND NFP.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf &
                 "                  AND NFP.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf &
                 "                  AND NFP.Serie_Id        = NF.Serie_Id " & vbCrLf &
                 "                  AND NFP.Nota_Id         = NF.Nota_Id " & vbCrLf &
                 "                  AND NFP.TipoDeDocumento = 15)) " & vbCrLf

        'NUBA NÃO PODE CONSULTAR FRETE AQUI EM NOTAS GERAIS - FURLAN - 07/08/2025
        'ORIX NÃO PODE CONSULTAR FRETE AQUI EM NOTAS GERAIS - FURLAN - 18/09/2025
        If Left(Session("ssEmpresa"), 8) = "53267147" OrElse Left(Session("ssEmpresa"), 8) = "62747840" Then
            strSQL &= " And (NF.TipoDeDocumento not in (2,3,4,8,10,11,14,57,58,67))" & vbCrLf
        End If

        If Not Session("ssChaveNfG" & HID.Value) Is Nothing Then
            Dim ssChaveNfe As String() = Session("ssChaveNfG" & HID.Value)
            strSQL &= " And NF.Empresa_Id      ='" & ssChaveNfe(0) & "'" & vbCrLf &
                      " And NF.EndEmpresa_Id   = " & ssChaveNfe(1) & vbCrLf &
                      " And NF.Cliente_Id      ='" & ssChaveNfe(2) & "'" & vbCrLf &
                      " And NF.EndCliente_Id   = " & ssChaveNfe(3) & vbCrLf &
                      " And NF.Nota_Id         = " & ssChaveNfe(6) & vbCrLf &
                      " And NF.Serie_Id        ='" & ssChaveNfe(5) & "'" & vbCrLf &
                      " And NF.EntradaSaida_Id ='" & ssChaveNfe(4) & "'"
        Else
            If Session("ssCnpjDaEmpresa" & HID.Value) IsNot Nothing AndAlso Session("ssCnpjDaEmpresa" & HID.Value).ToString.Length > 0 Then
                strSQL &= " AND (NF.Empresa_Id = '" & Session("ssCnpjDaEmpresa" & HID.Value) & "')"
                strSQL &= " AND (NF.EndEmpresa_Id = " & Session("ssEndDaEmpresa" & HID.Value) & ")"
            End If

            If Session("txtCnpjDoCliente" & HID.Value) IsNot Nothing AndAlso Session("txtCnpjDoCliente" & HID.Value).ToString.Length > 0 Then
                strSQL &= " AND (NF.Cliente_Id = '" & Session("txtCnpjDoCliente" & HID.Value) & "')"
                strSQL &= " AND (NF.EndCliente_Id = '" & Session("txtEndDoCliente" & HID.Value) & "')"
            End If

            If Session("txtNumero" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtNumero" & HID.Value)) Then
                strSQL &= "AND NF.Nota_Id = " & Session("txtNumero" & HID.Value) & " " & vbCrLf
            End If

            If Session("txtSerie" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtSerie" & HID.Value)) Then
                strSQL &= "AND NF.Serie_Id = '" & Session("txtSerie" & HID.Value) & "' " & vbCrLf
            End If

            If Session("txtDataDeEmissao" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtDataDeEmissao" & HID.Value)) AndAlso
                Session("txtDataDeEntrada" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtDataDeEntrada" & HID.Value)) Then
                'strSQL &= "AND NF.Movimento >= '" & Session("txtDataDeEmissao" & HID.Value).tosqldate() & "' " & vbCrLf

                strSQL &= " AND NF.Movimento between '" & CDate(Session("txtDataDeEmissao" & HID.Value)).ToString("yyyy-MM-dd") & "' and '" & CDate(Session("txtDataDeEntrada" & HID.Value)).ToString("yyyy-MM-dd") & "'" & vbCrLf
            End If

            'If Session("txtDataDeEntrada" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtDataDeEntrada" & HID.Value)) Then
            '    strSQL &= "AND NF.Movimento <= '" & Session("txtDataDeEntrada" & HID.Value).tosqldate() & "' " & vbCrLf
            'End If
        End If

        strSQL &= "ORDER BY NF.Movimento DESC, NF.Serie_Id, NF.Nota_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "NotasFiscaisGerais")
        GridNotas.DataSource = ds
        GridNotas.DataBind()

        If GridNotas.Rows.Count = 1 Then
            GridNotas.SelectedIndex = 0
            SelecionarNota()
        End If
        Return GridNotas.Rows.Count
    End Function

    Private Function CargaNotasDeTerceiro() As Integer
        Dim strSQL As String
        strSQL = " SELECT DISTINCT TOP 50" & vbCrLf &
                 "        NF.Movimento," & vbCrLf &
                 "		  NF.Empresa_Id," & vbCrLf &
                 "		  NF.EndEmpresa_Id," & vbCrLf &
                 "		  NF.Cliente_Id AS Cliente," & vbCrLf &
                 "		  NF.EndCliente_Id AS EndCliente," & vbCrLf &
                 "        C.Nome as NomeCliente," & vbCrLf &
                 "		  NF.EntradaSaida_Id AS ES," & vbCrLf &
                 "		  NF.Serie_Id AS Serie," & vbCrLf &
                 "		  NF.Nota_Id AS Nota," & vbCrLf &
                 "		  NFxI.Operacao," & vbCrLf &
                 "		  NFxI.SubOperacao," & vbCrLf &
                 "		  NF.NossaEmissao," & vbCrLf &
                 "		  NFxI.QuantidadeFiscal AS Quantidade," & vbCrLf &
                 "		  NFxI.Unitario," & vbCrLf &
                 "		  NFxI.Valor," & vbCrLf &
                 "		  NFxR.Romaneio_Id AS Romaneio," & vbCrLf &
                 "        Convert(nvarchar(2),NF.Situacao) + '-' + S.Descricao as DescSituacao" & vbCrLf &
                 "   FROM NotasFiscais NF" & vbCrLf &
                 "  INNER JOIN NotasFiscaisxItens NFxI" & vbCrLf &
                 "	   ON NF.Empresa_Id      = NFxI.Empresa_Id" & vbCrLf &
                 "    AND NF.EndEmpresa_Id   = NFxI.EndEmpresa_Id" & vbCrLf &
                 "    AND NF.Cliente_Id      = NFxI.Cliente_Id" & vbCrLf &
                 "    AND NF.EndCliente_Id   = NFxI.EndCliente_Id" & vbCrLf &
                 "    AND NF.EntradaSaida_Id = NFxI.EntradaSaida_Id" & vbCrLf &
                 "    AND NF.Serie_Id        = NFxI.Serie_Id" & vbCrLf &
                 "    AND NF.Nota_Id         = NFxI.Nota_Id" & vbCrLf &
                 "   LEFT OUTER JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf &
                 "	   ON NFxI.Empresa_Id      = NFxR.Empresa_Id" & vbCrLf &
                 "    AND NFxI.EndEmpresa_Id   = NFxR.EndEmpresa_Id" & vbCrLf &
                 "    AND NFxI.Cliente_Id      = NFxR.Cliente_Id" & vbCrLf &
                 "    AND NFxI.EndCliente_Id   = NFxR.EndCliente_Id" & vbCrLf &
                 "    AND NFxI.EntradaSaida_Id = NFxR.EntradaSaida_Id" & vbCrLf &
                 "    AND NFxI.Serie_Id        = NFxR.Serie_Id" & vbCrLf &
                 "    AND NFxI.Nota_Id         = NFxR.Nota_Id" & vbCrLf &
                 " Inner Join Situacoes S" & vbCrLf &
                 "    on S.Situacao_id = NF.Situacao" & vbCrLf &
                 "  INNER JOIN Clientes E " & vbCrLf &
                 "     ON E.Cliente_Id = NF.Empresa_Id " & vbCrLf &
                 "    AND E.Endereco_Id = NF.EndEmpresa_Id " & vbCrLf &
                 "  INNER JOIN Clientes C " & vbCrLf &
                 "     ON C.Cliente_Id = NF.Cliente_Id " & vbCrLf &
                 "    AND C.Endereco_Id = NF.EndCliente_Id " & vbCrLf &
                 "  INNER JOIN SubOperacoes SO" & vbCrLf &
                 "     ON NF.Operacao = SO.Operacao_Id" & vbCrLf &
                 "    AND NF.SubOperacao = SO.SubOperacoes_Id" & vbCrLf &
                 "  WHERE NF.Empresa_Id NOT IN (SELECT Empresa_Id FROM ClientesXEmpresas) " & vbCrLf

        If Not Session("ssChaveNfG" & HID.Value) Is Nothing Then
            Dim ssChaveNfe As String() = Session("ssChaveNfG" & HID.Value)
            strSQL &= " And NF.Empresa_Id      ='" & ssChaveNfe(0) & "'" & vbCrLf &
                      " And NF.EndEmpresa_Id   = " & ssChaveNfe(1) & vbCrLf &
                      " And NF.Cliente_Id      ='" & ssChaveNfe(2) & "'" & vbCrLf &
                      " And NF.EndCliente_Id   = " & ssChaveNfe(3) & vbCrLf &
                      " And NF.Nota_Id         = " & ssChaveNfe(6) & vbCrLf &
                      " And NF.Serie_Id        ='" & ssChaveNfe(5) & "'" & vbCrLf &
                      " And NF.EntradaSaida_Id ='" & ssChaveNfe(4) & "'"
        Else
            If Session("ssCnpjDaEmpresa" & HID.Value) IsNot Nothing AndAlso Session("ssCnpjDaEmpresa" & HID.Value).ToString.Length > 0 Then
                strSQL &= " AND (NF.Empresa_Id = '" & Session("ssCnpjDaEmpresa" & HID.Value) & "')"
                strSQL &= " AND (NF.EndEmpresa_Id = " & Session("ssEndDaEmpresa" & HID.Value) & ")"
            End If

            If Session("txtCnpjDoCliente" & HID.Value) IsNot Nothing AndAlso Session("txtCnpjDoCliente" & HID.Value).ToString.Length > 0 Then
                strSQL &= " AND (NF.Cliente_Id = '" & Session("txtCnpjDoCliente" & HID.Value) & "')"
                strSQL &= " AND (NF.EndCliente_Id = '" & Session("txtEndDoCliente" & HID.Value) & "')"
            End If

            If Session("txtNumero" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtNumero" & HID.Value)) Then
                strSQL &= "AND NF.Nota_Id = " & Session("txtNumero" & HID.Value) & " " & vbCrLf
            End If

            If Session("txtSerie" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtSerie" & HID.Value)) Then
                strSQL &= "AND NF.Serie_Id = '" & Session("txtSerie" & HID.Value) & "' " & vbCrLf
            End If

            If Session("txtDataDeEmissao" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtDataDeEmissao" & HID.Value)) AndAlso
                Session("txtDataDeEntrada" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtDataDeEntrada" & HID.Value)) Then
                'strSQL &= "AND NF.Movimento >= '" & Session("txtDataDeEmissao" & HID.Value).tosqldate() & "' " & vbCrLf

                strSQL &= " AND NF.Movimento between '" & CDate(Session("txtDataDeEmissao" & HID.Value)).ToString("yyyy-MM-dd") & "' and '" & CDate(Session("txtDataDeEntrada" & HID.Value)).ToString("yyyy-MM-dd") & "'" & vbCrLf
            End If

            'If Session("txtDataDeEntrada" & HID.Value) IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Session("txtDataDeEntrada" & HID.Value)) Then
            '    strSQL &= "AND NF.Movimento <= '" & Session("txtDataDeEntrada" & HID.Value).tosqldate() & "' " & vbCrLf
            'End If
        End If

        strSQL &= "ORDER BY NF.Movimento DESC, NF.Serie_Id, NF.Nota_Id" & vbCrLf

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Notas")
        GridNotas.DataSource = ds
        GridNotas.DataBind()

        If GridNotas.Rows.Count = 1 Then
            GridNotas.SelectedIndex = 0
            SelecionarNota()
        End If
        Return GridNotas.Rows.Count
    End Function

    Private Function CargaNotas() As Integer
        Dim ds As New DataSet
        Dim sql As String
        sql = " SELECT NotasFiscais.Empresa_id, NotasFiscais.EndEmpresa_id, NotasFiscais.Movimento, NotasFiscais.Cliente_Id AS Cliente, NotasFiscais.EndCliente_Id AS EndCliente, C.Nome as NomeCliente," & vbCrLf &
              "        NotasFiscais.EntradaSaida_Id AS ES, NotasFiscais.Serie_Id AS Serie, NotasFiscais.Nota_Id AS Nota, NotasFiscais.Operacao, NotasFiscais.SubOperacao, NF.NossaEmissao, " & vbCrLf &
              "        NotasFiscaisXItens.QuantidadeFiscal AS Quantidade, NotasFiscaisXItens.Unitario, NotasFiscaisXItens.Valor," & vbCrLf &
              "        nfxr.Romaneio_Id as Romaneio," & vbCrLf &
              "        Convert(nvarchar(2),NotasFiscais.Situacao) + '-' + S.Descricao as DescSituacao" & vbCrLf &
              "   FROM NotasFiscais" & vbCrLf &
              "  INNER JOIN NotasFiscaisXItens" & vbCrLf &
              "     ON NotasFiscais.Empresa_Id      = NotasFiscaisXItens.Empresa_Id" & vbCrLf &
              "    AND NotasFiscais.EndEmpresa_Id   = NotasFiscaisXItens.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscais.Cliente_Id      = NotasFiscaisXItens.Cliente_Id" & vbCrLf &
              "    AND NotasFiscais.EndCliente_Id   = NotasFiscaisXItens.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscais.EntradaSaida_Id = NotasFiscaisXItens.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscais.Serie_Id        = NotasFiscaisXItens.Serie_Id" & vbCrLf &
              "    AND NotasFiscais.Nota_Id         = NotasFiscaisXItens.Nota_Id" & vbCrLf &
              "   LEFT JOIN NotasFiscaisXRomaneios nfxr" & vbCrLf &
              "     ON NotasFiscaisXItens.Empresa_Id      = nfxr.Empresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndEmpresa_Id   = nfxr.EndEmpresa_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Cliente_Id      = nfxr.Cliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EndCliente_Id   = nfxr.EndCliente_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.EntradaSaida_Id = nfxr.EntradaSaida_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Serie_Id        = nfxr.Serie_Id" & vbCrLf &
              "    AND NotasFiscaisXItens.Nota_Id         = nfxr.Nota_Id" & vbCrLf &
              " Inner Join Situacoes S" & vbCrLf &
              "    on S.Situacao_id = NotasFiscais.Situacao" & vbCrLf &
              "  INNER JOIN Clientes C" & vbCrLf &
              "     ON C.Cliente_Id  = NotasFiscais.Cliente_Id" & vbCrLf &
              "    AND C.Endereco_id = NotasFiscais.EndCliente_Id" & vbCrLf &
              "  WHERE NotasFiscais.Situacao = 1 "

        If HttpContext.Current.Session("ssCnpjDaEmpresa" & HID.Value) IsNot Nothing AndAlso HttpContext.Current.Session("ssCnpjDaEmpresa" & HID.Value).ToString.Length > 0 Then
            sql &= " AND (NotasFiscais.Empresa_Id = '" & HttpContext.Current.Session("ssCnpjDaEmpresa" & HID.Value) & "')"
            sql &= " AND (NotasFiscais.EndEmpresa_Id = " & HttpContext.Current.Session("ssEndDaEmpresa" & HID.Value) & ")"
        End If

        If HttpContext.Current.Session("txtCnpjDoCliente" & HID.Value) IsNot Nothing AndAlso HttpContext.Current.Session("txtCnpjDoCliente" & HID.Value).ToString.Length > 0 Then
            sql &= " AND (NotasFiscais.Cliente_Id = '" & HttpContext.Current.Session("txtCnpjDoCliente" & HID.Value) & "')"
            sql &= " AND (NotasFiscais.EndCliente_Id = '" & HttpContext.Current.Session("txtEndDoCliente" & HID.Value) & "')"
        End If

        If HttpContext.Current.Session("txtPedido" & HID.Value) IsNot Nothing AndAlso HttpContext.Current.Session("txtPedido" & HID.Value).ToString.Length > 0 Then
            sql &= " AND (NotasFiscais.Pedido = " & HttpContext.Current.Session("txtPedido" & HID.Value) & ") "
        End If

        sql &= " Order  By NotasFiscais.Movimento DESC, NotasFiscais.EntradaSaida_Id, NotasFiscais.Serie_Id, NotasFiscais.Nota_Id"

        ds = Banco.ConsultaDataSet(sql, "Consulta")

        GridNotas.DataSource = ds
        GridNotas.DataBind()
        Return GridNotas.Rows.Count
    End Function

    Private Function getTextoConsultar(ByVal nfe As [Lib].Negocio.NotaFiscal) As String
        Dim sb As New StringBuilder()
        sb.Append("CHAVENFE=" & nfe.ChaveNFE & ControlChars.CrLf)
        sb.Append("NRECIBO =" & nfe.ReciboNota & ControlChars.CrLf)
        Return sb.ToString()
    End Function

    Private Sub ConsultarNFe(ByVal rowIndex As Integer)
        Dim Sqls As New ArrayList
        Dim Sql As String = String.Empty
        Dim aux As Boolean = True
        Dim msgNFE As String = String.Empty
        Try
            Dim objNFe As New [Lib].Negocio.NotaFiscal()
            objNFe.CodigoEmpresa = GridNotas.Rows(rowIndex).Cells(2).Text.Trim()
            objNFe.EnderecoEmpresa = GridNotas.Rows(rowIndex).Cells(3).Text.Trim()
            objNFe.CodigoCliente = GridNotas.Rows(rowIndex).Cells(4).Text.Trim()
            objNFe.EnderecoCliente = GridNotas.Rows(rowIndex).Cells(5).Text.Trim()
            objNFe.EntradaSaida = IIf(GridNotas.Rows(rowIndex).Cells(7).Text.Trim() = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            objNFe.Serie = GridNotas.Rows(rowIndex).Cells(8).Text.Trim()
            objNFe.Codigo = GridNotas.Rows(rowIndex).Cells(9).Text.Trim()
            objNFe = New [Lib].Negocio.NotaFiscal(objNFe)

            Sql = "Select * from  NFEPendencias "
            Sql &= "WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND "
            Sql &= "  Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND "
            Sql &= "  EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
            Sql &= "  Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; "
            Sqls.Add(Sql)

            Dim dsPendencia As New DataSet
            dsPendencia = Banco.ConsultaDataSet(Sql, "Consulta")

            Sqls.Clear()
            Sql = String.Empty

            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                objNFe.ChaveNFE = dsPendencia.Tables(0).Rows(0).Item("ChaveNfe")
                objNFe.ProtocoloNota = dsPendencia.Tables(0).Rows(0).Item("Protocolo")
            End If

            If Not String.IsNullOrWhiteSpace(objNFe.ChaveNFE) AndAlso Not String.IsNullOrWhiteSpace(objNFe.ProtocoloNota) Then
                Dim obj As New [Lib].Negocio.Fil()
                obj.IUD = "I"
                obj.NomeArquivo = String.Format("consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)
                obj.Texto = getTextoConsultar(objNFe)
                obj.SalvarSql(Sqls)

                If Not Banco.GravaBanco(Sqls) Then
                    MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                    Exit Sub
                End If

                'AGUARDAR RESPOSTA SEFAZ
                Dim resp As [Lib].Negocio.Resp = Nothing
                Dim fileName As String = String.Format("resp-consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa)

                While resp Is Nothing
                    resp = GetResp(fileName)
                End While

                If resp IsNot Nothing Then
                    Dim lstMsg As String() = resp.Texto.Split(New Char() {Environment.NewLine}, StringSplitOptions.RemoveEmptyEntries)
                    Dim resultado As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RESULTADO")).FirstOrDefault()
                    Dim mensagem As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("MENSAGEM")).FirstOrDefault()
                    Dim chave As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("CHAVE")).FirstOrDefault()
                    Dim recibo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("RECIBO")).FirstOrDefault()
                    Dim protocolo As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NPROTOCOLO")).FirstOrDefault()
                    Dim lote As String = lstMsg.Where(Function(s) s.ToUpper().Trim().Contains("NUMEROLOTE")).FirstOrDefault()

                    Dim strCodigo As String = String.Empty
                    If resultado IsNot Nothing AndAlso resultado.Length > 0 Then
                        strCodigo = resultado.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strMsg As String = String.Empty
                    If mensagem IsNot Nothing AndAlso mensagem.Length > 0 Then
                        strMsg = mensagem.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strChave As String = String.Empty
                    If chave IsNot Nothing AndAlso chave.Length > 0 Then
                        strChave = chave.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strRecibo As String = String.Empty
                    If recibo IsNot Nothing AndAlso recibo.Length > 0 Then
                        strRecibo = recibo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strProtocolo As String = String.Empty
                    If protocolo IsNot Nothing AndAlso protocolo.Length > 0 Then
                        strProtocolo = protocolo.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    Dim strLote As String = String.Empty
                    If lote IsNot Nothing AndAlso lote.Length > 0 Then
                        strLote = lote.Split(New Char() {"="c}, StringSplitOptions.RemoveEmptyEntries)(1).Trim()
                    End If

                    If Not String.IsNullOrWhiteSpace(strCodigo) AndAlso (strCodigo = "100" OrElse strCodigo = "101" OrElse strCodigo = "151" OrElse strCodigo = "110" OrElse strCodigo = "302") Then
                        Dim strOperacao = String.Empty
                        Dim strSituacao = String.Empty
                        Sqls.Clear()
                        Sql = String.Empty

                        If strCodigo = "100" OrElse strCodigo = "110" OrElse strCodigo = "302" Then
                            strOperacao = "Incluir"
                            If strCodigo = "110" Or strCodigo = "302" Then
                                strSituacao = CInt(eSituacao.Denegado)
                            Else
                                strSituacao = CInt(eSituacao.Normal)
                            End If

                            If Not dsPendencia Is Nothing AndAlso dsPendencia.Tables(0).Rows.Count > 0 Then
                                Sql = "Delete from  NFEPendencias "
                                Sql &= "WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND "
                                Sql &= "  Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND "
                                Sql &= "  EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
                                Sql &= "  Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; "
                                Sqls.Add(Sql)

                                If Not dsPendencia.Tables(0).Rows(0).Item("DadosAdicionais").ToString.Contains("CANCELAR") Then
                                    Sql = "INSERT INTO NFERealizadas "
                                    Sql &= "(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, "
                                    Sql &= "NfExpress, Operacao, Retorno, MsgRetorno, Recibo, ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, Protocolo, RegDPEC) "
                                    Sql &= "VALUES ('" & objNFe.CodigoEmpresa & "', " & objNFe.EnderecoEmpresa & ", '" & objNFe.CodigoCliente & "', '"
                                    Sql &= objNFe.EnderecoCliente & "', '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "', '" & objNFe.Serie & "', '" & objNFe.Codigo & "', '"
                                    Sql &= objNFe.DataInclusao.ToSqlDate() & "', '" & Format(objNFe.DataInclusao, "HH:mm:ss") & "', '" & UsuarioServidor.NomeUsuario & "', '"
                                    Sql &= String.Format("nota{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "', 'INCLUIR', '"
                                    Sql &= strCodigo & "', '" & strMsg & "', '" & strRecibo & "', '" & strChave & "', '" & strLote & "', '1', '" & objNFe.Observacoes & "', '', '" & strProtocolo & "', ''); "
                                    Sqls.Add(Sql)
                                End If

                                If strCodigo = "100" Then
                                    If Not String.IsNullOrWhiteSpace(objNFe.Empresa.EmailNFE) AndAlso Not String.IsNullOrWhiteSpace(objNFe.Cliente.EmailNFE) Then
                                        [Lib].Negocio.DocumentoEletronico.SendMailNFe(objNFe, "", False)
                                    End If

                                    If [Lib].Negocio.DocumentoEletronico.ImprimirNFe(objNFe, objNFe.Empresa.Empresa.ViasNFE, msgNFE, True) Then
                                        Try
                                            Dim bytes As Byte() = New FilesManager().getFile(String.Format("{0}.pdf", objNFe.ChaveNFE), eTipoDeDocumento.Nota)
                                            If bytes IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objNFe.ChaveNFE) Then
                                                Dim caminhoArquivo As String = Server.MapPath(String.Format("~/Files/{0}.pdf", objNFe.ChaveNFE))
                                                System.IO.File.WriteAllBytes(caminhoArquivo, bytes)
                                                Funcoes.AbrirArquivo(Me.Page, String.Format("Files/{0}.pdf", objNFe.ChaveNFE))
                                            End If
                                        Catch ex As Exception
                                            msgNFE = "Não foi possível encontrar o arquivo DANFE da nota fiscal " & objNFe.Codigo & "-" & objNFe.Serie & "!"
                                        End Try
                                    End If
                                End If
                            Else
                                Sql = "UPDATE NFERealizadas SET Operacao = '" & strOperacao & "', Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' "
                                Sql &= "WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND "
                                Sql &= "  Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND "
                                Sql &= "  EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
                                Sql &= "  Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; "
                                Sqls.Add(Sql)
                            End If

                            Sql = "UPDATE NotasFiscais Set Situacao = " & strSituacao & " "
                            Sql &= "WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = '" & objNFe.EnderecoEmpresa & "' AND "
                            Sql &= "      Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = '" & objNFe.EnderecoCliente & "' AND "
                            Sql &= "      EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND Serie_Id = '" & objNFe.Serie & "' AND "
                            Sql &= "      Nota_Id = '" & objNFe.Codigo & "'; "
                            Sqls.Add(Sql)
                        ElseIf strCodigo = "101" OrElse strCodigo = "151" Then
                            CType(Me.Page, NotaFiscalXItens).CancelarNotaFiscal(Sqls, objNFe)
                        Else
                            Sql = "UPDATE NFEPendencias SET Retorno = '" & strCodigo & "', MsgRetorno = '" & strMsg & "' "
                            Sql &= "WHERE Empresa_Id = '" & objNFe.CodigoEmpresa & "' AND EndEmpresa_Id = " & objNFe.EnderecoEmpresa & " AND "
                            Sql &= "  Cliente_Id = '" & objNFe.CodigoCliente & "' AND EndCliente_Id = " & objNFe.EnderecoCliente & " AND "
                            Sql &= "  EntradaSaida_Id = '" & IIf(objNFe.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' AND "
                            Sql &= "  Serie_Id = '" & objNFe.Serie & "' AND Nota_Id = " & objNFe.Codigo & "; "
                            Sqls.Add(Sql)
                        End If

                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-consulta{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-email{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)
                        Sql = "DELETE FROM nfe.resp WHERE nomearquivoresp = '" & String.Format("resp-danfe{0:000000000}#{1}.txt", objNFe.Codigo, objNFe.CodigoEmpresa) & "';"
                        Sqls.Add(Sql)

                        If Not Banco.GravaBanco(Sqls) Then
                            MsgBox(Me.Page, HttpContext.Current.Session("ssMessage"))
                            Exit Sub
                        ElseIf Not String.IsNullOrEmpty(msgNFE) Then
                            MsgBox(Me.Page, msgNFE)
                            Exit Sub
                        End If
                    End If

                    MsgBox(Me.Page, String.Format("{0} - {1}", strCodigo, strMsg))
                Else
                    MsgBox(Me.Page, "Sefaz não retornou nenhuma resposta, consulte novamente.")
                End If
            End If
        Catch ex As Exception
            MsgBox(Me.Page, ex.Message, eTitulo.Erro)
        End Try
    End Sub

    Private Function GetResp(ByVal nomeArquivo As String) As [Lib].Negocio.Resp
        Dim obj As New [Lib].Negocio.Resp(nomeArquivo)
        If Not obj.Codigo > 0 Then
            Return Nothing
        End If
        Return obj
    End Function

    Protected Sub GridNotas_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles GridNotas.SelectedIndexChanged
        SelecionarNota()
    End Sub

    Protected Sub GridNotas_RowDataBound(sender As Object, e As System.Web.UI.WebControls.GridViewRowEventArgs) Handles GridNotas.RowDataBound
        If e.Row.RowType = DataControlRowType.DataRow Then
            Dim x As Integer = e.Row.Cells(16).Text.Split("-")(0)

            Dim lnkConsultar As LinkButton = CType(e.Row.FindControl("lnkConsultar"), LinkButton)
            Dim hdfNossaEmissao As HiddenField = CType(e.Row.FindControl("hdfNossaEmissao"), HiddenField)
            lnkConsultar.Visible = False

            e.Row.Cells(16).Font.Bold = True
            Select Case x
                Case 1
                    e.Row.Cells(16).ForeColor = Drawing.Color.Green
                    If hdfNossaEmissao.Value.Split("-")(0) = "S" AndAlso
                        hdfNossaEmissao.Value.Split("-").Length > 1 AndAlso
                        hdfNossaEmissao.Value.Split("-")(1) <> "100" Then
                        e.Row.Cells(16).ForeColor = Drawing.Color.Red
                        e.Row.Cells(16).BackColor = Drawing.Color.Yellow
                        lnkConsultar.Visible = True
                    End If
                Case 2
                    e.Row.Cells(16).ForeColor = Drawing.Color.White
                    e.Row.Cells(16).BackColor = Drawing.Color.Red

                    If hdfNossaEmissao.Value.Split("-")(0) = "S" AndAlso
                        hdfNossaEmissao.Value.Split("-").Length > 1 AndAlso
                        hdfNossaEmissao.Value.Split("-")(1) <> "101" Then
                        e.Row.Cells(16).ForeColor = Drawing.Color.Red
                        e.Row.Cells(16).BackColor = Drawing.Color.Yellow
                        lnkConsultar.Visible = True
                    End If
                Case 10
                    e.Row.Cells(16).ForeColor = Drawing.Color.White
                    e.Row.Cells(16).BackColor = Drawing.Color.Red
                Case Else
                    e.Row.Cells(16).ForeColor = Drawing.Color.Black
                    e.Row.Cells(16).BackColor = Drawing.Color.Yellow
                    If hdfNossaEmissao.Value.Split("-")(0) = "S" AndAlso x = 7 Then lnkConsultar.Visible = True
            End Select
        End If
    End Sub

    Protected Sub btnFechar_Click(ByVal sender As Object, ByVal e As EventArgs) Handles btnFechar.Click
        Popup.CloseDialog(Me.Page, "divConsultaPedidosXNotas")
    End Sub

    Protected Sub lnkConsultar_Click(sender As Object, e As EventArgs)
        Dim rowIndex As Integer = CType(CType(sender, LinkButton).NamingContainer, GridViewRow).RowIndex
        ConsultarNFe(rowIndex)
        CargaNotasXItens()
    End Sub

    Protected Sub SelecionarNota()
        Dim sql As String = "SELECT Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, NfExpress, Operacao, Retorno, MsgRetorno, Recibo, " & vbCrLf &
              "       ChaveNfe, Lote, TpEmis, ObservacoesFiscais, DadosAdicionais, ISNULL(Protocolo, '') AS Protocolo, RegDPEC " & vbCrLf &
              "  FROM NFERealizadas " & vbCrLf &
              " WHERE Empresa_Id      ='" & GridNotas.SelectedRow.Cells(2).Text() & "'" & vbCrLf &
              "   AND EndEmpresa_Id   = " & GridNotas.SelectedRow.Cells(3).Text() & vbCrLf &
              "   AND Cliente_Id      ='" & GridNotas.SelectedRow.Cells(4).Text() & "'" & vbCrLf &
              "   AND EndCliente_Id   = " & GridNotas.SelectedRow.Cells(5).Text() & vbCrLf &
              "   AND EntradaSaida_Id ='" & GridNotas.SelectedRow.Cells(7).Text() & "'" & vbCrLf

        If Session("ssNomeUsuario") = "FURLAN" AndAlso GridNotas.SelectedRow.Cells(8).Text() = "&nbsp;" Then
            sql &= "   AND Serie_Id        =''" & vbCrLf
        Else
            sql &= "   AND Serie_Id        ='" & GridNotas.SelectedRow.Cells(8).Text() & "'" & vbCrLf
        End If

        sql &= "   AND Nota_Id         = " & GridNotas.SelectedRow.Cells(9).Text() & vbCrLf

        Dim ds As New DataSet
        ds = Banco.ConsultaDataSet(sql, "NFERealizadas")

        If ds Is Nothing Then
            HttpContext.Current.Session("ssNFE") = "N"
        ElseIf ds.Tables(0).Rows.Count = 0 Then
            HttpContext.Current.Session("ssNFE") = "N"
        Else
            HttpContext.Current.Session("ssNFE") = "S"
            HttpContext.Current.Session("ssChaveNFE") = ds.Tables(0).Rows(0).Item("ChaveNfe")
            HttpContext.Current.Session("ssProtocolo") = ds.Tables(0).Rows(0).Item("Protocolo")
        End If

        Dim objNFConsulta As New [Lib].Negocio.NotaFiscal()
        objNFConsulta.CodigoEmpresa = Server.HtmlDecode(GridNotas.SelectedRow.Cells(2).Text().Trim())
        objNFConsulta.EnderecoEmpresa = Server.HtmlDecode(GridNotas.SelectedRow.Cells(3).Text().Trim())
        objNFConsulta.CodigoCliente = Server.HtmlDecode(GridNotas.SelectedRow.Cells(4).Text().Trim())
        objNFConsulta.EnderecoCliente = Server.HtmlDecode(GridNotas.SelectedRow.Cells(5).Text().Trim())
        objNFConsulta.EntradaSaida = IIf(Server.HtmlDecode(GridNotas.SelectedRow.Cells(7).Text().Trim()) = "E", 0, 1)
        objNFConsulta.Serie = Server.HtmlDecode(GridNotas.SelectedRow.Cells(8).Text())
        objNFConsulta.Codigo = Server.HtmlDecode(GridNotas.SelectedRow.Cells(9).Text().Trim())
        objNFConsulta.ChaveNFE = ""
        objNFConsulta.Eletronica = False

        If HttpContext.Current.Session("ssNFE") = "S" Then
            If Not IsDBNull(ds.Tables(0).Rows(0).Item("ChaveNfe")) Then
                objNFConsulta.ChaveNFE = ds.Tables(0).Rows(0).Item("ChaveNfe")
                objNFConsulta.Eletronica = True
            End If
        End If

        If TypeOf Me.Page Is NotaFiscalXItens Then
            Session("objNFConsultaNXI" & HID.Value) = objNFConsulta
            Session("objNotaFiscal" & HID.Value) = objNFConsulta
            CType(Me.Page, NotaFiscalXItens).CarregarConsulta()
        ElseIf TypeOf Me.Page Is AlterarEncargo Then
            Session("objNFConsultaNXI" & HID.Value) = objNFConsulta
            Session("objNotaFiscal" & HID.Value) = objNFConsulta
            CType(Me.Page, AlterarEncargo).CarregarItem()
        ElseIf TypeOf Me.Page Is AlterarEncargoCTE Then
            Session("objNFConsultaNXI" & HID.Value) = objNFConsulta
            Session("objNotaFiscal" & HID.Value) = objNFConsulta
            CType(Me.Page, AlterarEncargoCTE).CarregarItem()
        ElseIf TypeOf Me.Page Is AlterarDepositoNotaFiscal Then
            Session("objNFConsultaNXI" & HID.Value) = objNFConsulta
            Session("objNotaFiscal" & HID.Value) = objNFConsulta
            CType(Me.Page, AlterarDepositoNotaFiscal).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarNotaFiscal Then
            Session("objNFConsultaNXI" & HID.Value) = objNFConsulta
            Session("objNotaFiscal" & HID.Value) = objNFConsulta
            CType(Me.Page, AlterarNotaFiscal).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is NotasFiscaisGerais Then
            Session("objNFConsultaNFG" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, NotasFiscaisGerais).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is NotasDeTerceiro Then
            Session("objNFConsultaTerceiro" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, NotasDeTerceiro).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeObservacao Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeObservacao).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeVolumes Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeVolumes).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeProduto Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeProduto).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarProdutoNotaGerais Then
            Session("objNFConsultaNFG" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarProdutoNotaGerais).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AjustarNumeroDoLote Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AjustarNumeroDoLote).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeTransporte Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeTransporte).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeCFOP Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeCFOP).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeModalidadeDeFrete Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeModalidadeDeFrete).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarMovimentoNotaFiscal Then
            Session("objNFConsulta" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarMovimentoNotaFiscal).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarNumeroPecas Then
            Session("objNFConsulta" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarNumeroPecas).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeVencimento Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeVencimento).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is CCeLocalEmbarque Then
            Session("objNFConsultaCCe" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, CCeLocalEmbarque).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarFinalidadeNotaFiscal Then
            Session("objNFConsulta" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarFinalidadeNotaFiscal).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarChaveNFE Then
            Session("objNFConsulta" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarChaveNFE).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarVencimentosNF Then
            Session("objNFConsultaNFTitulo" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarVencimentosNF).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is AlterarLaudo Then
            Session("objNFConsultaALTL" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, AlterarLaudo).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is RateioDePesagem Then
            Session("objNFConsultaRAT" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, RateioDePesagem).CarregarNotaFiscal()
        ElseIf TypeOf Me.Page Is DesvincularRomaneioNotaFiscal Then
            Session("objNFConsulta" & HID.Value) = New [Lib].Negocio.NotaFiscal(objNFConsulta)
            CType(Me.Page, DesvincularRomaneioNotaFiscal).CarregarNotaFiscal()
        End If

        Popup.CloseDialog(Me.Page, "divConsultaPedidosXNotas")
    End Sub

End Class