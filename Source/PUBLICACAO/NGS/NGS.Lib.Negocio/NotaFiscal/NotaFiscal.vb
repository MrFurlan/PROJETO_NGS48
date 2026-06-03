Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports System.Configuration
Imports System.Web
Imports NGS.Lib.Uteis
Imports System.IO
Imports System.Xml
Imports System.Globalization
Imports System.Web.UI.WebControls
Imports System.Net
Imports NGS.Lib.Negocio

<Serializable()>
Public Class ListNotasFiscais
    Inherits List(Of NotaFiscal)
    Implements IBaseEntity

    Public Sub New()

    End Sub

    Public Sub New(ByVal Empresa As String, ByVal EndEmpresa As String, ByVal DataInicial As String, ByVal DataFinal As String, Optional ByVal Cliente As String = "", Optional ByVal EndCliente As String = "", Optional ByVal EntradaSaida As String = "", Optional ByVal Pedido As String = "", Optional ByVal TipoDeDocumento As Integer = 0, Optional ByVal NumeroDaNota As Integer = 0, Optional ByVal bAnularPeriodo As Boolean = False)

        Dim strSQL As String = "SELECT isnull(Ped.Safra,'') as Safra," & vbCrLf &
                               "       isnull(Ped.UnidadeDeNegocio,UNI.Empresa_Id) as UnidadeDeNegocio," & vbCrLf &
                               "       NF.Empresa_Id," & vbCrLf &
                               "       NF.EndEmpresa_Id," & vbCrLf &
                               "       NF.Cliente_Id," & vbCrLf &
                               "       NF.EndCliente_Id," & vbCrLf &
                               "       NF.EntradaSaida_Id," & vbCrLf &
                               "       NF.Serie_Id," & vbCrLf &
                               "       NF.Nota_Id," & vbCrLf &
                               "       isnull(NF.TipoDeDocumento,1) as TipoDeDocumento," & vbCrLf &
                               "       isnull(NF.Formulario,NF.Nota_Id) as Formulario," & vbCrLf &
                               "       NF.Pedido," & vbCrLf &
                               "       isnull(NF.Procuracao,0) as Procuracao," & vbCrLf &
                               "       NF.Operacao," & vbCrLf &
                               "       NF.SubOperacao," & vbCrLf &
                               "       isnull(NF.Finalidade,0) as Finalidade," & vbCrLf &
                               "       NF.Movimento," & vbCrLf &
                               "       NF.DataDaNota," & vbCrLf &
                               "       NF.NossaEmissao, " & vbCrLf &
                               "       isnull(NF.SerieNotadoProdutor,'') as Serie, " & vbCrLf &
                               "       isnull(NF.NumeroNotadoProdutor,0) as NumeroNotadoProdutor," & vbCrLf &
                               "       isnull(NF.Deposito,'') as Deposito," & vbCrLf &
                               "       isnull(NF.EndDeposito,0) as EndDeposito," & vbCrLf &
                               "       isnull(NF.Destino,'') as Destino," & vbCrLf &
                               "       isnull(NF.EndDestino,0) as EndDestino," & vbCrLf &
                               "       isnull(NF.Transbordo,'') as Transbordo," & vbCrLf &
                               "       isnull(NF.EndTransbordo,0) as EndTransbordo," & vbCrLf &
                               "       isnull(NF.Agenciador,'') as Agenciador," & vbCrLf &
                               "       isnull(NF.EndAgenciador,0) as EndAgenciador," & vbCrLf &
                               "       isnull(NF.Entrega,'') as Entrega," & vbCrLf &
                               "       isnull(NF.EndEntrega,0) as EndEntrega," & vbCrLf &
                               "       isnull(NF.CIFFOB,'NEN') as CIFFOB," & vbCrLf &
                               "       isnull(NF.Observacoes,'') as Observacoes," & vbCrLf &
                               "       isnull(NF.Eletronica,'N') as Eletronica," & vbCrLf &
                               "       isnull(NF.UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
                               "       isnull(NF.UsuarioInclusaoData,CURRENT_TIMESTAMP) as UsuarioInclusaoData," & vbCrLf &
                               "       NF.Situacao," & vbCrLf &
                               "       isnull(NxT.Proprietario,'')as Transportador," & vbCrLf &
                               "       isnull(NxT.EndProprietario,0)as EndTransportador," & vbCrLf &
                               "       isnull(NxT.Motorista,'')as Motorista," & vbCrLf &
                               "       isnull(NxT.EndMotorista,0)as EndMotorista," & vbCrLf &
                               "       isnull(NxT.Placa,'') as placa," & vbCrLf &
                               "       isnull(NF.ObservacoesDeEmbarque,'') as ObservacoesDeEmbarque," & vbCrLf &
                               "       isnull(NF.ObservacoesDoProduto,'') as ObservacoesDoProduto," & vbCrLf &
                               "       isnull(NF.ObservacoesControleInterno,'') as ObservacoesControleInterno," & vbCrLf &
                               "       isnull(NfeR.ChaveNfe,'') as ChaveNfe," & vbCrLf &
                               "       isnull(NfeR.SegCodBarra,'') as SegCodBarra," & vbCrLf &
                               "       isnull(NfeR.Protocolo,'') as Protocolo," & vbCrLf &
                               "       isnull(NfeR.Recibo,'') as Recibo," & vbCrLf &
                               "       isnull(NR.Romaneio_id,0) as Romaneio_id," & vbCrLf &
                               "       isnull(NF.Autorizacao,0) as Autorizacao," & vbCrLf &
                               "       isnull(NF.UsuarioAlteracao,'') as UsuarioAlteracao," & vbCrLf &
                               "       isnull(NF.UsuarioAlteracaoData,CURRENT_TIMESTAMP) as UsuarioAlteracaoData," & vbCrLf &
                               "       isnull(NF.LocalEmbarque,'') as LocalEmbarque, isnull(NF.EndLocalEmbarque,0) as EndLocalEmbarque, " & vbCrLf &
                               "       isnull(NF.ContratoANTT,'') as ContratoANTT, isnull(NF.ProtocoloANTT,'') as ProtocoloANTT, " & vbCrLf &
                               "       isnull(NF.CartaoPgtoFrete,'') as CartaoPgtoFrete, isnull(NF.DataTermino,CURRENT_TIMESTAMP) as DataTermino, isnull(idPamcard,'') AS idPamcard, " & vbCrLf &
                               "       NF.TipoDeDocumentoFrete, " & vbCrLf &
                               "       isnull(NF.Favorecido,'') AS Favorecido, " & vbCrLf &
                               "       isnull(NF.EndFavorecido,0) AS EndFavorecido, " & vbCrLf &
                               "       isnull(NF.NFG,0) AS NFG, ISNULL(NF.Conferencia, 0) AS Conferencia, UsuarioConferencia, UsuarioConferenciaData, " & vbCrLf &
                               "       isnull(NF.Troca,0) as Troca, " & vbCrLf &
                               "       isnull(NF.ProprietarioDaMercadoria,'') AS ProprietarioDaMercadoria, " & vbCrLf &
                               "       isnull(NF.EndProprietarioDaMercadoria,0) AS EndProprietarioDaMercadoria, " & vbCrLf &
                               "       isnull(NF.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento," & vbCrLf &
                               "       isnull(NF.NFT,0) AS NFT, " & vbCrLf &
                               "       isnull(NF.Tomador,'') AS Tomador, " & vbCrLf &
                               "       isnull(NF.EndTomador,0) AS EndTomador, " & vbCrLf &
                               "       isnull(NF.ContratoDeFrete,0) AS ContratoDeFrete " & vbCrLf &
                               "  FROM NotasFiscais NF" & vbCrLf &
                               " INNER Join Pedidos Ped" & vbCrLf &
                               "   on NF.Empresa_id    = Ped.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id = Ped.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Pedido        = Ped.Pedido_Id" & vbCrLf &
                               "  LEFT JOIN GruposXEmpresas UNI" & vbCrLf &
                               "   on UNI.Cliente_Id     = NF.Empresa_Id" & vbCrLf &
                               "   and UNI.EndCliente_Id = NF.EndEmpresa_Id" & vbCrLf &
                               "  LEFT JOIN NotasFiscaisXTransportadores NxT" & vbCrLf &
                               "    on NF.Empresa_Id      = NxT.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id   = NxT.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Cliente_Id      = NxT.Cliente_Id" & vbCrLf &
                               "   and NF.EndCliente_Id   = NxT.EndCliente_Id" & vbCrLf &
                               "   and NF.EntradaSaida_Id = NxT.EntradaSaida_Id" & vbCrLf &
                               "   and NF.Serie_Id        = NxT.Serie_Id" & vbCrLf &
                               "   and NF.Nota_Id         = NxT.Nota_Id" & vbCrLf &
                               "  LEFT JOIN NfeRealizadas NfeR" & vbCrLf &
                               "    on NF.Empresa_Id      = NfeR.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id   = NfeR.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Cliente_Id      = NfeR.Cliente_Id" & vbCrLf &
                               "   and NF.EndCliente_Id   = NfeR.EndCliente_Id" & vbCrLf &
                               "   and NF.EntradaSaida_Id = NfeR.EntradaSaida_Id" & vbCrLf &
                               "   and NF.Serie_Id        = NfeR.Serie_Id" & vbCrLf &
                               "   and NF.Nota_Id         = NfeR.Nota_Id" & vbCrLf &
                               "  LEFT JOIN NotasFiscaisXRomaneios NR" & vbCrLf &
                               "    on NF.Empresa_Id      = NR.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id   = NR.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Cliente_Id      = NR.Cliente_Id" & vbCrLf &
                               "   and NF.EndCliente_Id   = NR.EndCliente_Id" & vbCrLf &
                               "   and NF.EntradaSaida_Id = NR.EntradaSaida_Id" & vbCrLf &
                               "   and NF.Serie_Id        = NR.Serie_Id" & vbCrLf &
                               "   and NF.Nota_Id         = NR.Nota_Id" & vbCrLf &
                               " WHERE NF.Situacao      = 1 " & vbCrLf &
                               "   AND NF.Empresa_Id    = '" & Empresa & "' " & vbCrLf &
                               "   AND NF.EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf &
                              "   AND NF.Nota_id > 0 " & vbCrLf


        If Cliente.Length > 0 Then
            strSQL &= "  AND NF.Cliente_Id    = '" & Cliente & "' " & vbCrLf &
                      "  AND NF.EndCliente_Id = " & EndCliente & " " & vbCrLf
        End If
        If DataInicial.Length > 0 Then
            strSQL &= "  AND NF.Movimento   >= '" & CDate(DataInicial).ToString("yyyy/MM/dd") & "' " & vbCrLf
        End If

        If DataFinal.Length > 0 Then
            strSQL &= "  AND NF.Movimento   <= '" & CDate(DataFinal).ToString("yyyy/MM/dd") & "' " & vbCrLf
        End If

        If EntradaSaida.Length > 0 Then
            strSQL &= "  AND NF.EntradaSaida_Id    = '" & EntradaSaida & "' "
        End If

        If NumeroDaNota > 0 Then
            strSQL &= "  AND NF.Nota_Id    = " & NumeroDaNota
        End If

        If Pedido.Length > 0 Then
            strSQL &= "  AND NF.Pedido  in(" & Pedido & ")"
        End If

        If TipoDeDocumento > 0 Then
            strSQL &= "  AND NF.TipoDeDocumento    = " & TipoDeDocumento
        End If

        Dim objBanco As New AcessaBanco
        Dim dsNotasFiscais As DataSet = objBanco.ConsultaDataSet(strSQL, "NotasFiscais")
        For Each row As DataRow In dsNotasFiscais.Tables(0).Rows
            Dim Nf As New Negocio.NotaFiscal
            Nf.CarregandoNota = True
            Nf.CodigoSafra = row("Safra")
            Nf.CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
            Nf.CodigoEmpresa = row("Empresa_Id")
            Nf.EnderecoEmpresa = row("EndEmpresa_Id")
            Nf.CodigoCliente = row("Cliente_Id")
            Nf.EnderecoCliente = row("EndCliente_Id")
            Nf.EntradaSaida = IIf(row("EntradaSaida_Id") = "S", 1, 0)
            Nf.Serie = row("Serie_Id")
            Nf.Codigo = row("Nota_Id")
            Nf.CodigoTipoDeDocumento = row("TipoDeDocumento")
            Nf.Formulario = row("Formulario")
            Nf.CodigoPedido = row("Pedido")
            Nf.CodigoProcuracao = row("Procuracao")
            Nf.CodigoOperacao = row("Operacao")
            Nf.CodigoSubOperacao = row("SubOperacao")
            Nf.CodigoFinalidade = row("Finalidade")
            Nf.Movimento = CDate(row("Movimento"))
            Nf.DataNota = CDate(row("DataDaNota"))
            Nf.NossaEmissao = row("NossaEmissao").Equals("S")
            Nf.SerieNotaProdutor = row("Serie")
            Nf.NotaProdutor = row("NumeroNotadoProdutor")
            Nf.CodigoDeposito = row("Deposito")
            Nf.EnderecoDeposito = row("EndDeposito")
            Nf.CodigoDestino = row("Destino")
            Nf.EnderecoDestino = row("EndDestino")
            Nf.CodigoTransbordo = row("Transbordo")
            Nf.EnderecoTransbordo = row("EndTransbordo")
            Nf.CodigoAgenciador = row("Agenciador")
            Nf.EnderecoAgenciador = row("EndAgenciador")
            Nf.CodigoEntrega = row("Entrega")
            Nf.EnderecoEntrega = row("EndEntrega")
            Nf.CIFFOB = [Enum].Parse(GetType(eTiposFrete), row("CIFFOB"))
            Nf.Observacoes = row("Observacoes")
            Nf.ObservacoesDeEmbarque = row("ObservacoesDeEmbarque")
            Nf.ObservacoesDoProduto = row("ObservacoesDoProduto")
            Nf.ObservacoesControleInterno = row("ObservacoesControleInterno")
            Nf.Eletronica = row("Eletronica").Equals("S")
            Nf.ChaveNFE = row("ChaveNfe")

            If row("SegCodBarra").ToString.Length > 0 Then
                Nf.TemSegCodBarra = True
            Else
                Nf.TemSegCodBarra = False
            End If
            Nf.SegCodBarra = row("SegCodBarra")

            Nf.ProtocoloNota = row("Protocolo")
            Nf.ReciboNota = row("Recibo")
            Nf.CodigoSituacao = row("Situacao")
            Nf.CodigoTransportador = row("Transportador")
            Nf.EnderecoTransportador = row("EndTransportador")
            Nf.PlacaTransportador = row("Placa")
            Nf.CodigoMotorista = row("Motorista")
            Nf.EnderecoMotorista = row("EndMotorista")
            Nf.CodigoRomaneio = row("Romaneio_id")
            Nf.CodigoAutorizacao = row("Autorizacao")
            Nf.UsuarioInclusao = row("UsuarioInclusao")
            Nf.DataInclusao = row("UsuarioInclusaoData")
            Nf.DataHoraNFE = CDate(row("UsuarioInclusaoData"))
            Nf.UsuarioAlteracao = row("UsuarioAlteracao")
            Nf.DataAlteracao = row("UsuarioAlteracaoData")
            Nf.CodigoLocalEmbarque = row("LocalEmbarque")
            Nf.EndLocalEmbarque = row("EndLocalEmbarque")
            Nf.ContratoANTT = row("ContratoANTT")
            Nf.ProtocoloANTT = row("ProtocoloANTT")
            Nf.CartaoPgtoFrete = row("CartaoPgtoFrete")
            Nf.DataTermino = row("DataTermino")
            Nf.idPamcard = row("idPamcard")
            If IsDBNull(row("TipoDeDocumentoFrete")) Then Nf.TipoDeDocumentoFrete = New Nullable(Of eTipoDeDocumentoFrete) Else Nf.TipoDeDocumentoFrete = CType(CInt(row("TipoDeDocumentoFrete")), eTipoDeDocumentoFrete)
            Nf.CodigoFavorecido = row("Favorecido")
            Nf.EnderecoFavorecido = row("EndFavorecido")
            Nf.CodigoProprietarioDaMercadoria = row("ProprietarioDaMercadoria")
            Nf.EnderecoProprietarioDaMercadoria = row("EndProprietarioDaMercadoria")
            Nf.NFG = row("NFG")
            Nf.Troca = row("Troca")
            Nf.CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
            If CBool(row("Conferencia")) Then
                Nf.Conferencia = row("Conferencia")
                If Not IsDBNull(row("UsuarioConferencia")) Then
                    Nf.UsuarioConferencia = row("UsuarioConferencia")
                End If
                If Not IsDBNull(row("UsuarioConferenciaData")) Then
                    Nf.UsuarioConferenciaData = row("UsuarioConferenciaData")
                End If
            End If
            Nf.NFT = row("NFT")

            Nf.CodigoTomador = row("Tomador")
            Nf.EnderecoTomador = row("EndTomador")

            Nf.ContratoDeFreteCTe = row("ContratoDeFrete")

            Me.Add(Nf)
        Next
    End Sub

End Class

<Serializable()>
Public Class NotaFiscal
    Implements IBaseEntity

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pNota As NotaFiscal)
        CarregarNota(pNota)
    End Sub

    Public Sub New(ByVal pCodigoTitulo As Integer)
        Dim sql As String
        sql = "Select Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id" & vbCrLf &
              "  From NotaFiscalXTitulo" & vbCrLf &
              " Where Titulo_Id = " & pCodigoTitulo &
              "   and nota_id > 0"


        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "NF")
        If ds.Tables(0).Rows.Count = 0 Then Exit Sub

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Dim NFConsulta As New NotaFiscal
        NFConsulta.CodigoEmpresa = row("Empresa_Id")
        NFConsulta.EnderecoEmpresa = row("EndEmpresa_Id")
        NFConsulta.CodigoCliente = row("Cliente_Id")
        NFConsulta.EnderecoCliente = row("EndCliente_Id")
        NFConsulta.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
        NFConsulta.Codigo = row("Nota_Id")
        NFConsulta.Serie = row("Serie_Id")
        CarregarNota(NFConsulta)
    End Sub

    Public Sub New(ByVal Nota As Integer, ByVal Empresa As String, ByVal EndEmpresa As String, ByVal Cliente As String, ByVal EndCliente As Integer, ByVal EntradaSaida As String, ByVal TipoDeDocumento As String, Optional ByVal DataInicial As String = "", Optional ByVal DataFinal As String = "")
        Dim strSQL As String = "SELECT isnull(Ped.Safra,'') as Safra," & vbCrLf &
                               "       isnull(Ped.UnidadeDeNegocio,UNI.Empresa_Id) as UnidadeDeNegocio," & vbCrLf &
                               "       NF.Empresa_Id," & vbCrLf &
                               "       NF.EndEmpresa_Id," & vbCrLf &
                               "       NF.Cliente_Id," & vbCrLf &
                               "       NF.EndCliente_Id," & vbCrLf &
                               "       NF.EntradaSaida_Id," & vbCrLf &
                               "       NF.Serie_Id," & vbCrLf &
                               "       NF.Nota_Id," & vbCrLf &
                               "       isnull(NF.TipoDeDocumento,1) as TipoDeDocumento," & vbCrLf &
                               "       isnull(NF.Formulario,NF.Nota_Id) as Formulario," & vbCrLf &
                               "       NF.Pedido," & vbCrLf &
                               "       isnull(NF.Procuracao,0) as Procuracao," & vbCrLf &
                               "       NF.Operacao," & vbCrLf &
                               "       NF.SubOperacao," & vbCrLf &
                               "       isnull(NF.Finalidade,0) as Finalidade," & vbCrLf &
                               "       NF.Movimento," & vbCrLf &
                               "       NF.DataDaNota," & vbCrLf &
                               "       NF.NossaEmissao, " & vbCrLf &
                               "       isnull(NF.SerieNotadoProdutor,'') as Serie, " & vbCrLf &
                               "       isnull(NF.NumeroNotadoProdutor,0) as NumeroNotadoProdutor," & vbCrLf &
                               "       isnull(NF.Deposito,'') as Deposito," & vbCrLf &
                               "       isnull(NF.EndDeposito,0) as EndDeposito," & vbCrLf &
                               "       isnull(NF.Destino,'') as Destino," & vbCrLf &
                               "       isnull(NF.EndDestino,0) as EndDestino," & vbCrLf &
                               "       isnull(NF.Transbordo,'') as Transbordo," & vbCrLf &
                               "       isnull(NF.EndTransbordo,0) as EndTransbordo," & vbCrLf &
                               "       isnull(NF.Agenciador,'') as Agenciador," & vbCrLf &
                               "       isnull(NF.EndAgenciador,0) as EndAgenciador," & vbCrLf &
                               "       isnull(NF.Entrega,'') as Entrega," & vbCrLf &
                               "       isnull(NF.EndEntrega,0) as EndEntrega," & vbCrLf &
                               "       isnull(NF.CIFFOB,'NEN') as CIFFOB," & vbCrLf &
                               "       isnull(NF.Observacoes,'') as Observacoes," & vbCrLf &
                               "       isnull(NF.Eletronica,'N') as Eletronica," & vbCrLf &
                               "       isnull(NF.UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
                               "       isnull(NF.UsuarioInclusaoData,CURRENT_TIMESTAMP) as UsuarioInclusaoData," & vbCrLf &
                               "       NF.Situacao," & vbCrLf &
                               "       isnull(NxT.Proprietario,'')as Transportador," & vbCrLf &
                               "       isnull(NxT.EndProprietario,0)as EndTransportador," & vbCrLf &
                               "       isnull(NxT.Motorista,'')as Motorista," & vbCrLf &
                               "       isnull(NxT.EndMotorista,0)as EndMotorista," & vbCrLf &
                               "       isnull(NxT.Placa,'') as placa," & vbCrLf &
                               "       isnull(NF.ObservacoesDeEmbarque,'') as ObservacoesDeEmbarque," & vbCrLf &
                               "       isnull(NF.ObservacoesDoProduto,'') as ObservacoesDoProduto," & vbCrLf &
                               "       isnull(NF.ObservacoesControleInterno,'') as ObservacoesControleInterno," & vbCrLf &
                               "       isnull(NfeR.ChaveNfe,'') as ChaveNfe," & vbCrLf &
                               "       isnull(NfeR.SegCodBarra,'') as SegCodBarra," & vbCrLf &
                               "       isnull(NfeR.Protocolo,'') as Protocolo," & vbCrLf &
                               "       isnull(NfeR.Recibo,'') as Recibo," & vbCrLf &
                               "       isnull(NR.Romaneio_id,0) as Romaneio_id," & vbCrLf &
                               "       isnull(NF.Autorizacao,0) as Autorizacao," & vbCrLf &
                               "       isnull(NF.UsuarioAlteracao,'') as UsuarioAlteracao," & vbCrLf &
                               "       isnull(NF.UsuarioAlteracaoData,CURRENT_TIMESTAMP) as UsuarioAlteracaoData," & vbCrLf &
                               "       isnull(NF.LocalEmbarque,'') as LocalEmbarque, isnull(NF.EndLocalEmbarque,0) as EndLocalEmbarque, " & vbCrLf &
                               "       ISNULL(NF.Representante,'') AS Representante, isnull(NF.EndRepresentante,0) AS EndRepresentante, " & vbCrLf &
                               "       isnull(NF.ContratoANTT,'') as ContratoANTT, isnull(NF.ProtocoloANTT,'') as ProtocoloANTT, " & vbCrLf &
                               "       isnull(NF.CartaoPgtoFrete,'') as CartaoPgtoFrete, isnull(NF.DataTermino,CURRENT_TIMESTAMP) as DataTermino, isnull(idPamcard,'') AS idPamcard, " & vbCrLf &
                               "       NF.TipoDeDocumentoFrete, " & vbCrLf &
                               "       isnull(NF.Favorecido,'') AS Favorecido, " & vbCrLf &
                               "       isnull(NF.EndFavorecido,0) AS EndFavorecido, " & vbCrLf &
                               "       isnull(NF.NFG,0) AS NFG, ISNULL(NF.Conferencia, 0) AS Conferencia, UsuarioConferencia, UsuarioConferenciaData, " & vbCrLf &
                               "       isnull(NF.Troca,0) as Troca, " & vbCrLf &
                               "       isnull(NF.ProprietarioDaMercadoria,'') AS ProprietarioDaMercadoria, " & vbCrLf &
                               "       isnull(NF.EndProprietarioDaMercadoria,0) AS EndProprietarioDaMercadoria, " & vbCrLf &
                               "       isnull(NF.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento, " & vbCrLf &
                               "       isnull(NF.InvoiceNavio,0) AS InvoiceNavio," & vbCrLf &
                               "       isnull(NF.NFT,0) AS NFT, " & vbCrLf &
                               "       isnull(NF.Tomador,'') AS Tomador, " & vbCrLf &
                               "       isnull(NF.EndTomador,0) AS EndTomador, " & vbCrLf &
                               "       isnull(NF.ContratoDeFrete,0) AS ContratoDeFrete " & vbCrLf &
                               "  FROM NotasFiscais NF" & vbCrLf &
                               " INNER Join Pedidos Ped" & vbCrLf &
                               "   on NF.Empresa_id    = Ped.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id = Ped.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Pedido        = Ped.Pedido_Id" & vbCrLf &
                               "  LEFT JOIN GruposXEmpresas UNI" & vbCrLf &
                               "   on UNI.Cliente_Id     = NF.Empresa_Id" & vbCrLf &
                               "   and UNI.EndCliente_Id = NF.EndEmpresa_Id" & vbCrLf &
                               "  LEFT JOIN NotasFiscaisXTransportadores NxT" & vbCrLf &
                               "    on NF.Empresa_Id      = NxT.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id   = NxT.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Cliente_Id      = NxT.Cliente_Id" & vbCrLf &
                               "   and NF.EndCliente_Id   = NxT.EndCliente_Id" & vbCrLf &
                               "   and NF.EntradaSaida_Id = NxT.EntradaSaida_Id" & vbCrLf &
                               "   and NF.Serie_Id        = NxT.Serie_Id" & vbCrLf &
                               "   and NF.Nota_Id         = NxT.Nota_Id" & vbCrLf &
                               "  LEFT JOIN NfeRealizadas NfeR" & vbCrLf &
                               "    on NF.Empresa_Id      = NfeR.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id   = NfeR.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Cliente_Id      = NfeR.Cliente_Id" & vbCrLf &
                               "   and NF.EndCliente_Id   = NfeR.EndCliente_Id" & vbCrLf &
                               "   and NF.EntradaSaida_Id = NfeR.EntradaSaida_Id" & vbCrLf &
                               "   and NF.Serie_Id        = NfeR.Serie_Id" & vbCrLf &
                               "   and NF.Nota_Id         = NfeR.Nota_Id" & vbCrLf &
                               "  LEFT JOIN NotasFiscaisXRomaneios NR" & vbCrLf &
                               "    on NF.Empresa_Id      = NR.Empresa_Id" & vbCrLf &
                               "   and NF.EndEmpresa_Id   = NR.EndEmpresa_Id" & vbCrLf &
                               "   and NF.Cliente_Id      = NR.Cliente_Id" & vbCrLf &
                               "   and NF.EndCliente_Id   = NR.EndCliente_Id" & vbCrLf &
                               "   and NF.EntradaSaida_Id = NR.EntradaSaida_Id" & vbCrLf &
                               "   and NF.Serie_Id        = NR.Serie_Id" & vbCrLf &
                               "   and NF.Nota_Id         = NR.Nota_Id" & vbCrLf &
                               " WHERE NF.Situacao      = 1 " & vbCrLf &
                               "   AND NF.Empresa_Id    = '" & Empresa & "' " & vbCrLf &
                               "   AND NF.EndEmpresa_Id = " & EndEmpresa & " " & vbCrLf &
                               "   AND NF.Nota_id > 0" & vbCrLf
        If Cliente.Length > 0 Then
            strSQL &= "  AND NF.Cliente_Id    = '" & Cliente & "' " & vbCrLf &
                      "  AND NF.EndCliente_Id = " & EndCliente & " " & vbCrLf
        End If

        If DataInicial.Length > 0 AndAlso DataFinal.Length > 0 Then
            strSQL &= " AND NF.Movimento BETWEEN '" & CDate(DataInicial).ToString("yyyy/MM/dd") & "' AND '" & CDate(DataFinal).ToString("yyyy/MM/dd") & "' "
        End If

        If EntradaSaida.Length > 0 Then
            strSQL &= "  AND NF.EntradaSaida_Id    = '" & EntradaSaida & "' "
        End If

        If Nota > 0 Then
            strSQL &= "  AND NF.Nota_Id    = " & Nota
        End If

        If TipoDeDocumento > 0 Then
            strSQL &= "  AND NF.TipoDeDocumento    = " & TipoDeDocumento
        End If

        Dim objBanco As New AcessaBanco
        Dim dsNotasFiscais As DataSet = objBanco.ConsultaDataSet(strSQL, "NotasFiscais")

        If dsNotasFiscais.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row As DataRow = dsNotasFiscais.Tables(0).Rows(0)

        CarregandoNota = True
        CodigoSafra = row("Safra")
        CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
        CodigoEmpresa = row("Empresa_Id")
        EnderecoEmpresa = row("EndEmpresa_Id")
        CodigoCliente = row("Cliente_Id")
        EnderecoCliente = row("EndCliente_Id")
        EntradaSaida = IIf(row("EntradaSaida_Id") = "S", 1, 0)
        Serie = row("Serie_Id")
        Codigo = row("Nota_Id")
        CodigoTipoDeDocumento = row("TipoDeDocumento")
        Formulario = row("Formulario")
        CodigoPedido = row("Pedido")
        CodigoProcuracao = row("Procuracao")
        CodigoOperacao = row("Operacao")
        CodigoSubOperacao = row("SubOperacao")
        CodigoFinalidade = row("Finalidade")
        Movimento = CDate(row("Movimento"))
        DataNota = CDate(row("DataDaNota"))
        NossaEmissao = row("NossaEmissao").Equals("S")
        SerieNotaProdutor = row("Serie")
        NotaProdutor = row("NumeroNotadoProdutor")
        CodigoDeposito = row("Deposito")
        EnderecoDeposito = row("EndDeposito")
        CodigoDestino = row("Destino")
        EnderecoDestino = row("EndDestino")
        CodigoTransbordo = row("Transbordo")
        EnderecoTransbordo = row("EndTransbordo")
        CodigoAgenciador = row("Agenciador")
        EnderecoAgenciador = row("EndAgenciador")
        CodigoEntrega = row("Entrega")
        EnderecoEntrega = row("EndEntrega")
        CIFFOB = [Enum].Parse(GetType(eTiposFrete), row("CIFFOB"))
        Observacoes = row("Observacoes")
        ObservacoesDeEmbarque = row("ObservacoesDeEmbarque")
        ObservacoesDoProduto = row("ObservacoesDoProduto")
        ObservacoesControleInterno = row("ObservacoesControleInterno")
        Eletronica = row("Eletronica").Equals("S")
        ChaveNFE = row("ChaveNfe")

        If row("SegCodBarra").ToString.Length > 0 Then
            TemSegCodBarra = True
        Else
            TemSegCodBarra = False
        End If
        SegCodBarra = row("SegCodBarra")

        ProtocoloNota = row("Protocolo")
        ReciboNota = row("Recibo")
        CodigoSituacao = row("Situacao")
        CodigoTransportador = row("Transportador")
        EnderecoTransportador = row("EndTransportador")
        PlacaTransportador = row("Placa")
        CodigoMotorista = row("Motorista")
        EnderecoMotorista = row("EndMotorista")
        CodigoRomaneio = row("Romaneio_id")
        CodigoAutorizacao = row("Autorizacao")
        UsuarioInclusao = row("UsuarioInclusao")
        DataInclusao = row("UsuarioInclusaoData")
        DataHoraNFE = CDate(row("UsuarioInclusaoData"))
        UsuarioAlteracao = row("UsuarioAlteracao")
        DataAlteracao = row("UsuarioAlteracaoData")
        CodigoLocalEmbarque = row("LocalEmbarque")
        EndLocalEmbarque = row("EndLocalEmbarque")
        CodigoRepresentante = row("Representante")
        EndRepresentante = row("EndRepresentante")
        ContratoANTT = row("ContratoANTT")
        ProtocoloANTT = row("ProtocoloANTT")
        CartaoPgtoFrete = row("CartaoPgtoFrete")
        DataTermino = row("DataTermino")
        idPamcard = row("idPamcard")
        If IsDBNull(row("TipoDeDocumentoFrete")) Then TipoDeDocumentoFrete = New Nullable(Of eTipoDeDocumentoFrete) Else TipoDeDocumentoFrete = CType(CInt(row("TipoDeDocumentoFrete")), eTipoDeDocumentoFrete)
        CodigoFavorecido = row("Favorecido")
        EnderecoFavorecido = row("EndFavorecido")
        CodigoProprietarioDaMercadoria = row("ProprietarioDaMercadoria")
        EnderecoProprietarioDaMercadoria = row("EndProprietarioDaMercadoria")
        NFG = row("NFG")
        Troca = row("Troca")
        CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
        InvoiceNavio = row("InvoiceNavio")
        If CBool(row("Conferencia")) Then
            Conferencia = row("Conferencia")
            If Not IsDBNull(row("UsuarioConferencia")) Then
                UsuarioConferencia = row("UsuarioConferencia")
            End If
            If Not IsDBNull(row("UsuarioConferenciaData")) Then
                UsuarioConferenciaData = row("UsuarioConferenciaData")
            End If
        End If
        NFT = row("NFT")

        CodigoTomador = row("Tomador")
        EnderecoTomador = row("EndTomador")

        ContratoDeFreteCTe = row("ContratoDeFrete")

    End Sub


    Private Sub CarregarNota(ByVal pNota As NotaFiscal)

        Dim Sql As String = String.Empty
        Sql = "SELECT NF.Empresa_Id," & vbCrLf &
              "       NF.EndEmpresa_Id," & vbCrLf &
              "       NF.Cliente_Id," & vbCrLf &
              "       NF.EndCliente_Id," & vbCrLf &
              "       NF.EntradaSaida_Id," & vbCrLf &
              "       NF.Serie_Id," & vbCrLf &
              "       NF.Nota_Id," & vbCrLf &
              "       isnull(NF.TipoDeDocumento,1) as TipoDeDocumento," & vbCrLf &
              "       isnull(NF.Formulario,NF.Nota_Id) as Formulario," & vbCrLf &
              "       isnull(NF.Pedido,0) as Pedido," & vbCrLf &
              "       isnull(NF.Procuracao,0) as Procuracao," & vbCrLf &
              "       NF.Operacao," & vbCrLf &
              "       NF.SubOperacao," & vbCrLf &
              "       isnull(NF.Finalidade,0) as Finalidade," & vbCrLf &
              "       NF.Movimento," & vbCrLf &
              "       NF.DataDaNota," & vbCrLf &
              "       NF.NossaEmissao, " & vbCrLf &
              "       isnull(NF.SerieNotadoProdutor,'') as Serie, " & vbCrLf &
              "       isnull(NF.NumeroNotadoProdutor,0) as NumeroNotadoProdutor," & vbCrLf &
              "       isnull(NF.Deposito,'') as Deposito," & vbCrLf &
              "       isnull(NF.EndDeposito,0) as EndDeposito," & vbCrLf &
              "       isnull(NF.Destino,'') as Destino," & vbCrLf &
              "       isnull(NF.EndDestino,0) as EndDestino," & vbCrLf &
              "       isnull(NF.Transbordo,'') as Transbordo," & vbCrLf &
              "       isnull(NF.EndTransbordo,0) as EndTransbordo," & vbCrLf &
              "       isnull(NF.Agenciador,'') as Agenciador," & vbCrLf &
              "       isnull(NF.EndAgenciador,0) as EndAgenciador," & vbCrLf &
              "       isnull(NF.Entrega,'') as Entrega," & vbCrLf &
              "       isnull(NF.EndEntrega,0) as EndEntrega," & vbCrLf &
              "       isnull(NF.CIFFOB,'NEN') as CIFFOB," & vbCrLf &
              "       isnull(NF.Observacoes,'') as Observacoes," & vbCrLf &
              "       isnull(NF.Eletronica,'N') as Eletronica," & vbCrLf &
              "       isnull(NF.UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
              "       isnull(NF.UsuarioInclusaoData,CURRENT_TIMESTAMP) as UsuarioInclusaoData," & vbCrLf &
              "       NF.Situacao," & vbCrLf &
              "       isnull(NxT.Proprietario,'')as Transportador," & vbCrLf &
              "       isnull(NxT.EndProprietario,0)as EndTransportador," & vbCrLf &
              "       isnull(NxT.Motorista,'')as Motorista," & vbCrLf &
              "       isnull(NxT.EndMotorista,0)as EndMotorista," & vbCrLf &
              "       isnull(NxT.Placa,'') as placa," & vbCrLf &
              "       isnull(NF.ObservacoesDeEmbarque,'') as ObservacoesDeEmbarque," & vbCrLf &
              "       isnull(NF.ObservacoesDoProduto,'') as ObservacoesDoProduto," & vbCrLf &
              "       isnull(NF.ObservacoesControleInterno,'') as ObservacoesControleInterno," & vbCrLf &
              "       isnull(CASE WHEN NfeR.ChaveNfe IS NULL THEN NFeC.ChaveNfe ELSE NfeR.ChaveNfe END,'') as ChaveNfe," & vbCrLf &
              "       isnull(CASE WHEN NfeR.SegCodBarra IS NULL THEN Nfec.SegCodBarra ELSE NfeR.SegCodBarra END,'') as SegCodBarra," & vbCrLf &
              "       isnull(CASE WHEN NfeR.Protocolo IS NULL THEN NFeC.Protocolo ELSE NfeR.Protocolo END,'') as Protocolo," & vbCrLf &
              "       isnull(CASE WHEN NfeR.Retorno IS NULL THEN NFeC.Retorno ELSE NfeR.Retorno END,'') as Retorno," & vbCrLf &
              "       isnull(CASE WHEN NfeR.MsgRetorno IS NULL THEN NFeC.MsgRetorno ELSE NfeR.MsgRetorno END,'') as MsgRetorno," & vbCrLf &
              "       CASE WHEN NfeR.Data IS NULL THEN NFeC.Data ELSE NfeR.Data END as Data," & vbCrLf &
              "       CASE WHEN NfeR.Hora IS NULL THEN NFeC.Hora ELSE NfeR.Hora END as Hora," & vbCrLf &
              "       isnull(NR.Romaneio_Id,0) as Romaneio_id," & vbCrLf &
              "       isnull(NF.Autorizacao,0) as Autorizacao," & vbCrLf &
              "       isnull(NF.UsuarioAlteracao,'') as UsuarioAlteracao," & vbCrLf &
              "       isnull(NF.UsuarioAlteracaoData,CURRENT_TIMESTAMP) as UsuarioAlteracaoData," & vbCrLf &
              "       isnull(NF.LocalEmbarque,'') as LocalEmbarque, isnull(NF.EndLocalEmbarque,0) as EndLocalEmbarque, " & vbCrLf &
              "       ISNULL(NF.Representante,'') AS Representante, isnull(NF.EndRepresentante,0) AS EndRepresentante, " & vbCrLf &
              "       isnull(NF.ContratoANTT,'') as ContratoANTT, isnull(NF.ProtocoloANTT,'') as ProtocoloANTT, " & vbCrLf &
              "       isnull(NF.CartaoPgtoFrete,'') as CartaoPgtoFrete, isnull(NF.DataTermino,CURRENT_TIMESTAMP) as DataTermino, isnull(idPamcard,'') AS idPamcard, " & vbCrLf &
              "       NF.TipoDeDocumentoFrete, " & vbCrLf &
              "       isnull(NF.Favorecido,'') AS Favorecido, " & vbCrLf &
              "       isnull(NF.EndFavorecido,0) AS EndFavorecido, " & vbCrLf &
              "       isnull(P.UnidadeDeNegocio,UNI.Empresa_Id) AS UnidadeDeNegocio," & vbCrLf &
              "       isnull(P.Safra,'') AS Safra, " & vbCrLf &
              "       isnull(NF.NFG,0) AS NFG, isnull(NF.Conferencia,0) AS Conferencia, UsuarioConferencia, UsuarioConferenciaData," & vbCrLf &
              "       isnull(NF.Troca,0) AS Troca, " & vbCrLf &
              "       isnull(NF.ProprietarioDaMercadoria,'') AS ProprietarioDaMercadoria, " & vbCrLf &
              "       isnull(NF.EndProprietarioDaMercadoria,0) AS EndProprietarioDaMercadoria, " & vbCrLf &
              "       isnull(NF.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento, " & vbCrLf &
              "       isnull(NF.InvoiceNavio,0) AS InvoiceNavio," & vbCrLf &
              "       isnull(NF.NFT,0) AS NFT, " & vbCrLf &
              "       isnull(NF.Tomador,'') AS Tomador, " & vbCrLf &
              "       isnull(NF.EndTomador,0) AS EndTomador, " & vbCrLf &
              "       isnull(NF.ContratoDeFrete,0) AS ContratoDeFrete " & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXTransportadores NxT" & vbCrLf &
              "    on NF.Empresa_Id      = NxT.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NxT.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NxT.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NxT.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NxT.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NxT.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NxT.Nota_Id" & vbCrLf &
              "  LEFT JOIN NFEContingencia NFeC" & vbCrLf &
              "    on NF.Empresa_Id      = NFeC.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NFeC.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NFeC.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NFeC.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NFeC.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NFeC.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NFeC.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfeRealizadas NfeR" & vbCrLf &
              "    on NF.Empresa_Id      = NfeR.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NfeR.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NfeR.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NfeR.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NfeR.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NfeR.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NfeR.Nota_Id" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXRomaneios NR" & vbCrLf &
              "    on NF.Empresa_Id      = NR.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NR.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NR.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NR.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NR.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NR.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NR.Nota_Id" & vbCrLf &
              "  LEFT JOIN Pedidos P" & vbCrLf &
              "    On P.Empresa_id    = NF.Empresa_id" & vbCrLf &
              "   and P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf &
              "   and P.Pedido_id     = NF.Pedido" & vbCrLf &
              "  LEFT JOIN GruposXEmpresas UNI" & vbCrLf &
              "   on UNI.Cliente_Id     = NF.Empresa_Id" & vbCrLf &
              "   and UNI.EndCliente_Id = NF.EndEmpresa_Id" & vbCrLf &
              "  WHERE NF.Empresa_Id      ='" & pNota.CodigoEmpresa & "'" & vbCrLf &
              "   and NF.EndEmpresa_Id   = " & pNota.EnderecoEmpresa.ToString & vbCrLf &
              "   and NF.EntradaSaida_Id ='" & pNota.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
              "   and NF.Nota_Id         = " & pNota.Codigo & vbCrLf &
              "   and NF.Serie_Id        ='" & pNota.Serie & "'" & vbCrLf & vbCrLf &
              "   and NF.Nota_id         > 0" & vbCrLf

        If Not String.IsNullOrWhiteSpace(pNota.CodigoCliente) Then
            Sql &= "   and NF.Cliente_Id      ='" & pNota.CodigoCliente & "'" & vbCrLf &
                   "   and NF.EndCliente_Id   = " & pNota.EnderecoCliente & vbCrLf
        End If

        If pNota.CodigoPedido > 0 Then
            Sql &= "   and NF.Pedido      = " & pNota.CodigoPedido & vbCrLf
        End If

        Dim db As New AcessaBanco
        Dim ds As New DataSet
        ds = db.ConsultaDataSet(Sql, "Nota")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row As DataRow = ds.Tables(0).Rows(0)

        CarregandoNota = True
        IUD = "U"
        CodigoEmpresa = row("Empresa_Id")
        EnderecoEmpresa = row("EndEmpresa_Id")
        CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
        CodigoSafra = row("Safra")
        CodigoCliente = row("Cliente_Id")
        EnderecoCliente = row("EndCliente_Id")
        EntradaSaida = IIf(row("EntradaSaida_Id") = "S", 1, 0)
        Serie = row("Serie_Id")
        Codigo = row("Nota_Id")
        CodigoTipoDeDocumento = row("TipoDeDocumento")
        Formulario = row("Formulario")
        CodigoPedido = row("Pedido")
        CodigoProcuracao = row("Procuracao")
        CodigoOperacao = row("Operacao")
        CodigoSubOperacao = row("SubOperacao")
        CodigoFinalidade = row("Finalidade")
        Movimento = CDate(row("Movimento"))
        DataNota = CDate(row("DataDaNota"))
        NossaEmissao = row("NossaEmissao").Equals("S")
        SerieNotaProdutor = row("Serie")
        NotaProdutor = row("NumeroNotadoProdutor")
        CodigoDeposito = row("Deposito")
        EnderecoDeposito = row("EndDeposito")
        CodigoDestino = row("Destino")
        EnderecoDestino = row("EndDestino")
        CodigoTransbordo = row("Transbordo")
        EnderecoTransbordo = row("EndTransbordo")
        CodigoAgenciador = row("Agenciador")
        EnderecoAgenciador = row("EndAgenciador")
        CodigoEntrega = row("Entrega")
        EnderecoEntrega = row("EndEntrega")
        Me.CIFFOB = [Enum].Parse(GetType(eTiposFrete), row("CIFFOB"))
        Observacoes = row("Observacoes")
        ObservacoesDeEmbarque = row("ObservacoesDeEmbarque")
        ObservacoesDoProduto = row("ObservacoesDoProduto")
        ObservacoesControleInterno = row("ObservacoesControleInterno")
        Eletronica = row("Eletronica").Equals("S")
        ChaveNFE = row("ChaveNfe")

        If row("SegCodBarra").ToString.Length > 0 Then
            TemSegCodBarra = True
        Else
            TemSegCodBarra = False
        End If
        SegCodBarra = row("SegCodBarra")

        ProtocoloNota = row("Protocolo")
        Retorno = row("Retorno")
        MsgRetorno = row("MsgRetorno")

        CodigoSituacao = row("Situacao")
        CodigoTransportador = row("Transportador")
        EnderecoTransportador = row("EndTransportador")
        PlacaTransportador = row("Placa")
        CodigoMotorista = row("Motorista")
        EnderecoMotorista = row("EndMotorista")
        CodigoRomaneio = row("Romaneio_id")
        CodigoAutorizacao = row("Autorizacao")
        UsuarioInclusao = row("UsuarioInclusao")
        DataInclusao = row("UsuarioInclusaoData")
        DataHoraNFE = CDate(row("UsuarioInclusaoData"))

        If Not IsDBNull(row("Data")) AndAlso Not IsDBNull(row("Hora")) Then
            DataHoraNFE = CDate(CDate(row("Data")).ToString("yyyy-MM-dd") & " " & CDate(row("Hora")).ToString("HH:mm:ss"))
        End If

        UsuarioAlteracao = row("UsuarioAlteracao")
        DataAlteracao = row("UsuarioAlteracaoData")
        CodigoLocalEmbarque = row("LocalEmbarque")
        EndLocalEmbarque = row("EndLocalEmbarque")
        CodigoRepresentante = row("Representante")
        EndRepresentante = row("EndRepresentante")
        ContratoANTT = row("ContratoANTT")
        ProtocoloANTT = row("ProtocoloANTT")
        CartaoPgtoFrete = row("CartaoPgtoFrete")
        DataTermino = row("DataTermino")
        idPamcard = row("idPamcard")
        If IsDBNull(row("TipoDeDocumentoFrete")) Then TipoDeDocumentoFrete = New Nullable(Of eTipoDeDocumentoFrete) Else TipoDeDocumentoFrete = CType(CInt(row("TipoDeDocumentoFrete")), eTipoDeDocumentoFrete)
        CodigoFavorecido = row("Favorecido")
        EnderecoFavorecido = row("EndFavorecido")

        CodigoProprietarioDaMercadoria = row("ProprietarioDaMercadoria")
        EnderecoProprietarioDaMercadoria = row("EndProprietarioDaMercadoria")

        NFG = row("NFG")
        Troca = row("Troca")
        CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
        InvoiceNavio = row("InvoiceNavio")

        If CBool(row("Conferencia")) AndAlso Not IsDBNull(row("UsuarioConferencia")) AndAlso Not IsDBNull(row("UsuarioConferenciaData")) Then
            Me.Conferencia = row("Conferencia")
            Me.UsuarioConferencia = row("UsuarioConferencia")
            Me.UsuarioConferenciaData = row("UsuarioConferenciaData")
        End If

        NFT = row("NFT")

        CodigoTomador = row("Tomador")
        EnderecoTomador = row("EndTomador")

        ContratoDeFreteCTe = row("ContratoDeFrete")

        'ja tem o carregando = true itens no new da ListNotaFiscalXItem
        Itens = New ListNotaFiscalXItem(Me)

        CarregandoItens = True
        NotasXNotas = New NotasXNotas(Me)
        AtualizaTotais() 'Aqui não pode atualizar 
        CarregandoItens = False

        Dim Embalagem As New NotasFiscaisXEmbalagens(Me)
        If Embalagem IsNot Nothing AndAlso Embalagem.Cliente IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Embalagem.Cliente) Then
            Quantidade = Embalagem.Quantidade
            Especie = Embalagem.Especie
            Marca = Embalagem.Marca
            Numero = Embalagem.Numero
            PesoBruto = Embalagem.PesoBruto
            PesoLiquido = Embalagem.PesoLiquido
        End If
        CarregandoNota = False
    End Sub

    Public Sub CarregarNotaComChaveXML(ByVal pchaveXML As String)

        Dim Sql As String = String.Empty
        Sql = "SELECT NF.Empresa_Id," & vbCrLf &
              "       NF.EndEmpresa_Id," & vbCrLf &
              "       NF.Cliente_Id," & vbCrLf &
              "       NF.EndCliente_Id," & vbCrLf &
              "       NF.EntradaSaida_Id," & vbCrLf &
              "       NF.Serie_Id," & vbCrLf &
              "       NF.Nota_Id," & vbCrLf &
              "       isnull(NF.TipoDeDocumento,1) as TipoDeDocumento," & vbCrLf &
              "       isnull(NF.Formulario,NF.Nota_Id) as Formulario," & vbCrLf &
              "       isnull(NF.Pedido,0) as Pedido," & vbCrLf &
              "       isnull(NF.Procuracao,0) as Procuracao," & vbCrLf &
              "       NF.Operacao," & vbCrLf &
              "       NF.SubOperacao," & vbCrLf &
              "       isnull(NF.Finalidade,0) as Finalidade," & vbCrLf &
              "       NF.Movimento," & vbCrLf &
              "       NF.DataDaNota," & vbCrLf &
              "       NF.NossaEmissao, " & vbCrLf &
              "       isnull(NF.SerieNotadoProdutor,'') as Serie, " & vbCrLf &
              "       isnull(NF.NumeroNotadoProdutor,0) as NumeroNotadoProdutor," & vbCrLf &
              "       isnull(NF.Deposito,'') as Deposito," & vbCrLf &
              "       isnull(NF.EndDeposito,0) as EndDeposito," & vbCrLf &
              "       isnull(NF.Destino,'') as Destino," & vbCrLf &
              "       isnull(NF.EndDestino,0) as EndDestino," & vbCrLf &
              "       isnull(NF.Transbordo,'') as Transbordo," & vbCrLf &
              "       isnull(NF.EndTransbordo,0) as EndTransbordo," & vbCrLf &
              "       isnull(NF.Agenciador,'') as Agenciador," & vbCrLf &
              "       isnull(NF.EndAgenciador,0) as EndAgenciador," & vbCrLf &
              "       isnull(NF.Entrega,'') as Entrega," & vbCrLf &
              "       isnull(NF.EndEntrega,0) as EndEntrega," & vbCrLf &
              "       isnull(NF.CIFFOB,'NEN') as CIFFOB," & vbCrLf &
              "       isnull(NF.Observacoes,'') as Observacoes," & vbCrLf &
              "       isnull(NF.Eletronica,'N') as Eletronica," & vbCrLf &
              "       isnull(NF.UsuarioInclusao,'') as UsuarioInclusao," & vbCrLf &
              "       isnull(NF.UsuarioInclusaoData,CURRENT_TIMESTAMP) as UsuarioInclusaoData," & vbCrLf &
              "       NF.Situacao," & vbCrLf &
              "       isnull(NxT.Proprietario,'')as Transportador," & vbCrLf &
              "       isnull(NxT.EndProprietario,0)as EndTransportador," & vbCrLf &
              "       isnull(NxT.Motorista,'')as Motorista," & vbCrLf &
              "       isnull(NxT.EndMotorista,0)as EndMotorista," & vbCrLf &
              "       isnull(NxT.Placa,'') as placa," & vbCrLf &
              "       isnull(NF.ObservacoesDeEmbarque,'') as ObservacoesDeEmbarque," & vbCrLf &
              "       isnull(NF.ObservacoesDoProduto,'') as ObservacoesDoProduto," & vbCrLf &
              "       isnull(NF.ObservacoesControleInterno,'') as ObservacoesControleInterno," & vbCrLf &
              "       isnull(CASE WHEN NfeR.ChaveNfe IS NULL THEN NFeC.ChaveNfe ELSE NfeR.ChaveNfe END,'') as ChaveNfe," & vbCrLf &
              "       isnull(CASE WHEN NfeR.SegCodBarra IS NULL THEN Nfec.SegCodBarra ELSE NfeR.SegCodBarra END,'') as SegCodBarra," & vbCrLf &
              "       isnull(CASE WHEN NfeR.Protocolo IS NULL THEN NFeC.Protocolo ELSE NfeR.Protocolo END,'') as Protocolo," & vbCrLf &
              "       isnull(CASE WHEN NfeR.Retorno IS NULL THEN NFeC.Retorno ELSE NfeR.Retorno END,'') as Retorno," & vbCrLf &
              "       isnull(CASE WHEN NfeR.MsgRetorno IS NULL THEN NFeC.MsgRetorno ELSE NfeR.MsgRetorno END,'') as MsgRetorno," & vbCrLf &
              "       CASE WHEN NfeR.Data IS NULL THEN NFeC.Data ELSE NfeR.Data END as Data," & vbCrLf &
              "       CASE WHEN NfeR.Hora IS NULL THEN NFeC.Hora ELSE NfeR.Hora END as Hora," & vbCrLf &
              "       isnull(NR.Romaneio_Id,0) as Romaneio_id," & vbCrLf &
              "       isnull(NF.Autorizacao,0) as Autorizacao," & vbCrLf &
              "       isnull(NF.UsuarioAlteracao,'') as UsuarioAlteracao," & vbCrLf &
              "       isnull(NF.UsuarioAlteracaoData,CURRENT_TIMESTAMP) as UsuarioAlteracaoData," & vbCrLf &
              "       isnull(NF.LocalEmbarque,'') as LocalEmbarque, isnull(NF.EndLocalEmbarque,0) as EndLocalEmbarque, " & vbCrLf &
              "       ISNULL(NF.Representante,'') AS Representante, isnull(NF.EndRepresentante,0) AS EndRepresentante, " & vbCrLf &
              "       isnull(NF.ContratoANTT,'') as ContratoANTT, isnull(NF.ProtocoloANTT,'') as ProtocoloANTT, " & vbCrLf &
              "       isnull(NF.CartaoPgtoFrete,'') as CartaoPgtoFrete, isnull(NF.DataTermino,CURRENT_TIMESTAMP) as DataTermino, isnull(idPamcard,'') AS idPamcard, " & vbCrLf &
              "       NF.TipoDeDocumentoFrete, " & vbCrLf &
              "       isnull(NF.Favorecido,'') AS Favorecido, " & vbCrLf &
              "       isnull(NF.EndFavorecido,0) AS EndFavorecido, " & vbCrLf &
              "       isnull(P.UnidadeDeNegocio,UNI.Empresa_Id) AS UnidadeDeNegocio," & vbCrLf &
              "       isnull(P.Safra,'') AS Safra, " & vbCrLf &
              "       isnull(NF.NFG,0) AS NFG, isnull(NF.Conferencia,0) AS Conferencia, UsuarioConferencia, UsuarioConferenciaData," & vbCrLf &
              "       isnull(NF.Troca,0) AS Troca, " & vbCrLf &
              "       isnull(NF.ProprietarioDaMercadoria,'') AS ProprietarioDaMercadoria, " & vbCrLf &
              "       isnull(NF.EndProprietarioDaMercadoria,0) AS EndProprietarioDaMercadoria, " & vbCrLf &
              "       isnull(NF.CodigoNaturezaDeRendimento,0) AS CodigoNaturezaDeRendimento, " & vbCrLf &
              "       isnull(NF.InvoiceNavio,0) AS InvoiceNavio," & vbCrLf &
              "       isnull(NF.NFT,0) AS NFT, " & vbCrLf &
              "       isnull(NF.Tomador,'') AS Tomador, " & vbCrLf &
              "       isnull(NF.EndTomador,0) AS EndTomador, " & vbCrLf &
              "       isnull(NF.ContratoDeFrete,0) AS ContratoDeFrete " & vbCrLf &
              "  FROM NotasFiscais NF" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXTransportadores NxT" & vbCrLf &
              "    on NF.Empresa_Id      = NxT.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NxT.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NxT.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NxT.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NxT.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NxT.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NxT.Nota_Id" & vbCrLf &
              "  LEFT JOIN NFEContingencia NFeC" & vbCrLf &
              "    on NF.Empresa_Id      = NFeC.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NFeC.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NFeC.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NFeC.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NFeC.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NFeC.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NFeC.Nota_Id" & vbCrLf &
              "  LEFT JOIN NfeRealizadas NfeR" & vbCrLf &
              "    on NF.Empresa_Id      = NfeR.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NfeR.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NfeR.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NfeR.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NfeR.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NfeR.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NfeR.Nota_Id" & vbCrLf &
              "  LEFT JOIN NotasFiscaisXRomaneios NR" & vbCrLf &
              "    on NF.Empresa_Id      = NR.Empresa_Id" & vbCrLf &
              "   and NF.EndEmpresa_Id   = NR.EndEmpresa_Id" & vbCrLf &
              "   and NF.Cliente_Id      = NR.Cliente_Id" & vbCrLf &
              "   and NF.EndCliente_Id   = NR.EndCliente_Id" & vbCrLf &
              "   and NF.EntradaSaida_Id = NR.EntradaSaida_Id" & vbCrLf &
              "   and NF.Serie_Id        = NR.Serie_Id" & vbCrLf &
              "   and NF.Nota_Id         = NR.Nota_Id" & vbCrLf &
              "  LEFT JOIN Pedidos P" & vbCrLf &
              "    On P.Empresa_id    = NF.Empresa_id" & vbCrLf &
              "   and P.EndEmpresa_id = NF.EndEmpresa_id" & vbCrLf &
              "   and P.Pedido_id     = NF.Pedido" & vbCrLf &
              "  INNER JOIN NFERealizadas NFR " & vbCrLf &
              "  ON NF.Empresa_Id      = NFR.Empresa_Id" & vbCrLf &
              "     AND NF.Cliente_Id      = NFR.Cliente_Id" & vbCrLf &
              "     AND NF.Nota_Id         = NFR.Nota_Id" & vbCrLf &
              "  LEFT JOIN GruposXEmpresas UNI" & vbCrLf &
              "   on UNI.Cliente_Id     = NF.Empresa_Id" & vbCrLf &
              "   and UNI.EndCliente_Id = NF.EndEmpresa_Id" & vbCrLf &
              "  WHERE  NFR.ChaveNfe      ='" & pchaveXML & "'"

        Dim db As New AcessaBanco
        Dim ds As New DataSet
        ds = db.ConsultaDataSet(Sql, "Nota")

        If ds.Tables(0).Rows.Count = 0 Then Exit Sub
        Dim row As DataRow = ds.Tables(0).Rows(0)

        CarregandoNota = True
        IUD = "U"
        CodigoEmpresa = row("Empresa_Id")
        EnderecoEmpresa = row("EndEmpresa_Id")
        CodigoUnidadeDeNegocio = row("UnidadeDeNegocio")
        CodigoSafra = row("Safra")
        CodigoCliente = row("Cliente_Id")
        EnderecoCliente = row("EndCliente_Id")
        EntradaSaida = IIf(row("EntradaSaida_Id") = "S", 1, 0)
        Serie = row("Serie_Id")
        Codigo = row("Nota_Id")
        CodigoTipoDeDocumento = row("TipoDeDocumento")
        Formulario = row("Formulario")
        CodigoPedido = row("Pedido")
        CodigoProcuracao = row("Procuracao")
        CodigoOperacao = row("Operacao")
        CodigoSubOperacao = row("SubOperacao")
        CodigoFinalidade = row("Finalidade")
        Movimento = CDate(row("Movimento"))
        DataNota = CDate(row("DataDaNota"))
        NossaEmissao = row("NossaEmissao").Equals("S")
        SerieNotaProdutor = row("Serie")
        NotaProdutor = row("NumeroNotadoProdutor")
        CodigoDeposito = row("Deposito")
        EnderecoDeposito = row("EndDeposito")
        CodigoDestino = row("Destino")
        EnderecoDestino = row("EndDestino")
        CodigoTransbordo = row("Transbordo")
        EnderecoTransbordo = row("EndTransbordo")
        CodigoAgenciador = row("Agenciador")
        EnderecoAgenciador = row("EndAgenciador")
        CodigoEntrega = row("Entrega")
        EnderecoEntrega = row("EndEntrega")
        Me.CIFFOB = [Enum].Parse(GetType(eTiposFrete), row("CIFFOB"))
        Observacoes = row("Observacoes")
        ObservacoesDeEmbarque = row("ObservacoesDeEmbarque")
        ObservacoesDoProduto = row("ObservacoesDoProduto")
        ObservacoesControleInterno = row("ObservacoesControleInterno")
        Eletronica = row("Eletronica").Equals("S")
        ChaveNFE = row("ChaveNfe")

        If row("SegCodBarra").ToString.Length > 0 Then
            TemSegCodBarra = True
        Else
            TemSegCodBarra = False
        End If
        SegCodBarra = row("SegCodBarra")

        ProtocoloNota = row("Protocolo")
        Retorno = row("Retorno")
        MsgRetorno = row("MsgRetorno")

        CodigoSituacao = row("Situacao")
        CodigoTransportador = row("Transportador")
        EnderecoTransportador = row("EndTransportador")
        PlacaTransportador = row("Placa")
        CodigoMotorista = row("Motorista")
        EnderecoMotorista = row("EndMotorista")
        CodigoRomaneio = row("Romaneio_id")
        CodigoAutorizacao = row("Autorizacao")
        UsuarioInclusao = row("UsuarioInclusao")
        DataInclusao = row("UsuarioInclusaoData")
        DataHoraNFE = CDate(row("UsuarioInclusaoData"))

        If Not IsDBNull(row("Data")) AndAlso Not IsDBNull(row("Hora")) Then
            DataHoraNFE = CDate(CDate(row("Data")).ToString("yyyy-MM-dd") & " " & CDate(row("Hora")).ToString("HH:mm:ss"))
        End If

        UsuarioAlteracao = row("UsuarioAlteracao")
        DataAlteracao = row("UsuarioAlteracaoData")
        CodigoLocalEmbarque = row("LocalEmbarque")
        EndLocalEmbarque = row("EndLocalEmbarque")
        CodigoRepresentante = row("Representante")
        EndRepresentante = row("EndRepresentante")
        ContratoANTT = row("ContratoANTT")
        ProtocoloANTT = row("ProtocoloANTT")
        CartaoPgtoFrete = row("CartaoPgtoFrete")
        DataTermino = row("DataTermino")
        idPamcard = row("idPamcard")
        If IsDBNull(row("TipoDeDocumentoFrete")) Then TipoDeDocumentoFrete = New Nullable(Of eTipoDeDocumentoFrete) Else TipoDeDocumentoFrete = CType(CInt(row("TipoDeDocumentoFrete")), eTipoDeDocumentoFrete)
        CodigoFavorecido = row("Favorecido")
        EnderecoFavorecido = row("EndFavorecido")

        CodigoProprietarioDaMercadoria = row("ProprietarioDaMercadoria")
        EnderecoProprietarioDaMercadoria = row("EndProprietarioDaMercadoria")

        NFG = row("NFG")
        Troca = row("Troca")
        CodigoNaturezaDeRendimento = row("CodigoNaturezaDeRendimento")
        InvoiceNavio = row("InvoiceNavio")

        If CBool(row("Conferencia")) AndAlso Not IsDBNull(row("UsuarioConferencia")) AndAlso Not IsDBNull(row("UsuarioConferenciaData")) Then
            Me.Conferencia = row("Conferencia")
            Me.UsuarioConferencia = row("UsuarioConferencia")
            Me.UsuarioConferenciaData = row("UsuarioConferenciaData")
        End If

        NFT = row("NFT")

        CodigoTomador = row("Tomador")
        EnderecoTomador = row("EndTomador")

        ContratoDeFreteCTe = row("ContratoDeFrete")


        'ja tem o carregando = true itens no new da ListNotaFiscalXItem
        Itens = New ListNotaFiscalXItem(Me)

        CarregandoItens = True
        NotasXNotas = New NotasXNotas(Me)
        AtualizaTotais() 'Aqui não pode atualizar 
        CarregandoItens = False

        Dim Embalagem As New NotasFiscaisXEmbalagens(Me)
        If Embalagem IsNot Nothing AndAlso Embalagem.Cliente IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(Embalagem.Cliente) Then
            Quantidade = Embalagem.Quantidade
            Especie = Embalagem.Especie
            Marca = Embalagem.Marca
            Numero = Embalagem.Numero
            PesoBruto = Embalagem.PesoBruto
            PesoLiquido = Embalagem.PesoLiquido
        End If
        CarregandoNota = False
    End Sub

#End Region

#Region "Fields"

    Private _NFG As Boolean = False
    Private _Troca As Boolean = False
    Private _Retroativa As Boolean = False

    Private _IUD As String
    Private _Empresa As Negocio.Cliente
    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer

    Private _CodigoUnidadeDeNegocio As String = ""
    Private _EndUnidadeDeNegocio As Integer
    Private _UnidadeDeNegocio As Negocio.Cliente

    Private _CodigoSafra As String = ""
    Private _Safra As Negocio.Safra

    Private _CodigoCliente As String = ""
    Private _EnderecoCliente As Integer
    Private _Cliente As Negocio.Cliente
    Private _EntradaSaida As Negocio.eEntradaSaida
    Private _Serie As String = ""
    Private _Codigo As Integer
    Private _Formulario As Integer

    Private _CodigoTipoDeDocumento As Integer = eTipoDeDocumento.Nota
    Private _TipoDeDocumento As Negocio.TipoDeDocumento

    Private _Cfop As Negocio.CFOP
    Private _NaturezaDaOperacao As String

    Private _CriarPedido As Boolean = False
    Private _CodigoPedido As Integer
    Private _CodigoPedidoMIC As String
    Private _Pedido As Negocio.Pedido

    Private _CodigoProcuracao As Integer
    Private _Procuracao As Negocio.Procuracao
    Private _ProcuracaoSaldoPedido As Decimal

    Private _CodigoOperacao As Integer
    Private _Operacao As Negocio.Operacao
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As Negocio.SubOperacao
    Private _CodigoFinalidade As Integer
    Private _Finalidade As Negocio.Finalidade

    Private _Movimento As DateTime = Date.Today
    Private _MomentoIndice As String = ""

    Private _DataNota As DateTime = Date.Today
    Private _NossaEmissao As Boolean = False
    Private _SerieNotaProdutor As String = ""
    Private _NotaProdutor As Integer

    Private _CodigoDeposito As String = ""
    Private _EnderecoDeposito As Integer
    Private _Deposito As Negocio.Cliente
    Private _CodigoDestino As String = ""
    Private _EnderecoDestino As Integer
    Private _Destino As Negocio.Cliente
    Private _CodigoAgenciador As String = ""
    Private _EnderecoAgenciador As Integer
    Private _Agenciador As Negocio.Cliente
    Private _CodigoEntrega As String = ""
    Private _EnderecoEntrega As Integer
    Private _Entrega As Negocio.Cliente
    Private _CIFFOB As Negocio.eTiposFrete
    Private _UsuarioInclusao As String = ""
    Private _DataInclusao As DateTime
    Private _UsuarioAlteracao As String = ""
    Private _DataAlteracao As DateTime
    Private _CodigoSituacao As Integer
    Private _Situacao As Negocio.Situacao
    Private _Itens As Negocio.ListNotaFiscalXItem

    'NFE ELETRONICA
    Private _Eletronica As Boolean = False
    Private _ChaveNFE As String = ""
    Private _TemSegCodBarra As Boolean = False
    Private _SegCodBarra As String = ""
    Private _ProtocoloNota As String
    Private _ReciboNota As String
    Private _Retorno As String
    Private _MsgRetorno As String
    Private _ObservacaoCancelamento As String = ""
    Private _Usuario As String
    Private _DataHoraNFE As DateTime?

    'RESUMOS
    Private _ValorBaseIcms As Decimal
    Private _ValorIcms As Decimal
    Private _ValorIPI As Decimal
    Private _ValorBaseIcmsST As Decimal
    Private _ValorIcmsST As Decimal
    Private _ValorFrete As Decimal
    Private _ValorAduaneira As Decimal
    Private _ValorDesconto As Decimal
    Private _ValorSeguro As Decimal

    'Totais
    Private _TotalProduto As Decimal
    Private _TotalNota As Decimal
    Private _TotalProdutoMoeda As Decimal
    Private _TotalNotaMoeda As Decimal

    'TRANSPORTADOR
    Private _CodigoTransportador As String = ""
    Private _EnderecoTransportador As Integer
    Private _Transportador As Negocio.Cliente
    Private _PlacaTransportador As String = ""
    Private _PlacaDetalhes As Negocio.Placa

    'TRANSBORDO
    Private _CodigoTransbordo As String = ""
    Private _EnderecoTransbordo As Integer
    Private _Transbordo As Negocio.Cliente

    'MOTORISTA
    Private _CodigoMotorista As String = ""
    Private _EnderecoMotorista As Integer
    Private _Motorista As Negocio.Cliente

    'PROPRIETéRIO DA MERCADORIA
    Private _CodigoProprietarioDaMercadoria As String = ""
    Private _EnderecoProprietarioDaMercadoria As Integer
    Private _ProprietarioDaMercadoria As Negocio.Cliente

    'OBSERVACOES
    Private _Observacoes As String = ""
    Private _ObservacoesDoProduto As String = ""
    Private _ObservacoesDeEmbarque As String = ""
    Private _ObservacoesControleInterno As String = ""

    'PESOS - NUMERACAO - VOLUMES
    Private _Quantidade As Decimal
    Private _Numero As String
    Private _Especie As String
    Private _Marca As String
    Private _PesoBruto As Decimal
    Private _PesoLiquido As Decimal

    'TROCA DE NOTA
    Private _NotasTrocaOrigem As List(Of Negocio.NotaFiscal)
    Private _NotaTrocaOrigem As Negocio.NotaFiscal
    Private _NotaTrocaDestino As Negocio.NotaFiscal
    Private _NotaDeTroca As Negocio.NotaFiscal

    Private _NotaOrigemImportacaoXML As List(Of Negocio.NotaFiscal)

    'ROMANEIRO
    Private _CriarRomaneio As Boolean = False
    Private _CodigoRomaneio As Integer
    Private _Romaneio As Negocio.Romaneio

    'AUTORIZACAO DE RETIRADA
    Private _Autorizacao As Negocio.AutorizacaoDeRetirada
    Private _CodigoAutorizacao As Integer = 0

    'CONTROLE
    Public CarregandoItens As Boolean = False
    Public CarregandoNota As Boolean = False

    Private _Razao As Negocio.Razao
    Private _LancamentosContabeis As ListRazao

    Private _NotasXNotas As Negocio.NotasXNotas
    Private _ListNotasXNotas As Negocio.ListNotasXNotas
    Private _PesoBrutoRecuperado As Double

    'VENCIMENTOS
    Private _VencimentosPedido As Negocio.ListTitulo
    Private _VencimentosNota As Negocio.ListTitulo

    Private _TitulosPedido As Novo.ListTituloNovo
    Private _Titulos As Novo.ListTituloNovo

    'LOCAL DE EMBARQUE
    Private _CodigoLocalEmbarque As String = ""
    Private _EndLocalEmbarque As Integer
    Private _LocalEmbarque As Negocio.Cliente

    'PAMCARD
    Private _ContratoANTT As String = ""
    Private _ProtocoloANTT As String = ""
    Private _CartaoPgtoFrete As String = ""
    Private _idPamcard As String = ""
    Private _DataTermino As DateTime = Date.Today
    Public CartaoNovo As Boolean = False

    'EXPORTACAO
    Private _DadosDaExportacao As Negocio.NotaFiscalXExportacao
    Private _DadosDaExportacaoRE As ListNotaFiscalXRE
    Private _DadosDaExportacaoCE As ListNotaFiscalXCE

    'IMPORTACAO
    Private _DadosDaImportacao As Negocio.NotaFiscalXImportacao

    'FRETE
    Private _CodigoTipoDeDocumentoFrete As Nullable(Of Integer)
    Private _TipoDeDocumentoFrete As Nullable(Of eTipoDeDocumentoFrete)
    Private _NotasReferenciais As ListNotaFiscalReferencial

    'FAVORECIDO
    Private _CodigoFavorecido As String = ""
    Private _EnderecoFavorecido As Integer

    'COMISSéES
    Private _ComissoesXBaixas As ListComissoesXBaixas

    'PERCUSOS
    Private _NotaFiscalXPercursos As ListNotaFiscalXPercurso

    'NOTA ORIGINAL
    Private _NotaFiscalOriginal As NotaFiscal

    'Arquivo
    Private _Arquivos As ListArquivo

    'Conferéncia da Nota Fiscal
    Private _Conferencia As Boolean = False
    Private _UsuarioConferencia As String
    Private _UsuarioConferenciaData As DateTime

    'XML obsoleto
    'Private _XML As String

    'Peso de Chegada
    Private _PesoDeChegada As NotaFiscalXDestino

    'Saldo
    Private _SaldosItensNota As ListSaldoPedido2015

    'Obriga NF Produtor
    Private _ObrigaNFProdutor As Boolean = False

    Private _ModoComplemento As Boolean = False

    'Verifica se Nota não foi recusada
    Private _TemRecusa As Boolean = False

    'Cédigo da Natureza de Redimento, utilizado no SpedReinf
    Private _CodigoNaturezaDeRendimento As Integer

    Private _NCMXML As String
    Private _FecharTelaClassificacao As Boolean
    Private _ParteNomeProdNCMXML As String

    Private _XMLImportado As Boolean
    '******* INFORMAção PARA CUSTO DO NAVIO ********
    Private _InvoiceNavio As Integer

    Private _MicSistemas As Boolean
    Private _NotaDeTerceiro As Boolean

    Private _VincularNotas As Boolean

    'Verifica se Nota Tem confirmação da Operação
    Private _TemConfirmacaoDaOperacao As Boolean = False

    Private _DiferencaValorNFXProdutoXML As Decimal
    Private _TotalNotaValorModificado As Decimal
    Private _TotalBaseICMSXML As Decimal
    Private _TotalICMSXML As Decimal

    'Representante
    Private _CodigoRepresentante As String = ""
    Private _EndRepresentante As Integer
    Private _Representante As Negocio.Cliente

    'NOTA FISCAL DE TERCEIROS
    Private _NFT As Boolean = False

    'REFAZER OS ENCARGO COM BASE NA DATA DA NOTA DE ORIGEM A DEVOLUÇÃO 
    Private _DataDevolucao As DateTime = Date.Today

    Private _EmpresaEmissor As Boolean

    'TOMADOR
    Private _CodigoTomador As String = ""
    Private _EnderecoTomador As Integer
    Private _Tomador As Negocio.Cliente

    'USADO NO CT-e  
    'Para saber se teve Contrato de Frete 
    Private _ContratoDeFreteCTe As Boolean = False
    'Para saber se está Emitindo o CT-e
    Private _EmitindoCTe As Boolean = False
    'Para saber se tem pedágio no CTe 
    Private _TemPedagioCTe As Boolean = False

#End Region

#Region "Property"
    Public Property NFG As Boolean
        Get
            Return _NFG
        End Get
        Set(value As Boolean)
            _NFG = value
        End Set
    End Property

    Public Property Troca As Boolean
        Get
            Return _Troca
        End Get
        Set(value As Boolean)
            _Troca = value
        End Set
    End Property

    Public Property Retroativa As Boolean
        Get
            Return _Retroativa
        End Get
        Set(value As Boolean)
            _Retroativa = value
        End Set
    End Property

    Public ReadOnly Property FinanceiroNovo As Boolean
        Get
            Return CBool(ConfigurationManager.AppSettings("financeiroNovo"))
        End Get
    End Property

    Public Property NotasXNotas() As Negocio.NotasXNotas
        Get
            If _NotasXNotas Is Nothing Then
                _NotasXNotas = New Negocio.NotasXNotas(Me)
                If _NotasXNotas.Serie <> "UN" Then
                    _PesoBrutoRecuperado = 0
                Else
                    Dim consNota As New Negocio.NotaFiscal
                    consNota.CodigoEmpresa = _NotasXNotas.OrigemEmpresaCnpj
                    consNota.EnderecoEmpresa = _NotasXNotas.OrigemEndEmpresa
                    consNota.CodigoCliente = _NotasXNotas.OrigemClienteCnpj
                    consNota.EnderecoCliente = _NotasXNotas.OrigemEndCliente
                    consNota.EntradaSaida = IIf(_NotasXNotas.OrigemEntradaSaida = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                    consNota.Codigo = _NotasXNotas.OrigemNota
                    consNota.Serie = _NotasXNotas.OrigemSerie
                    _ListNotasXNotas = New Negocio.ListNotasXNotas(consNota, False)
                    For Each dr As Negocio.NotasXNotas In _ListNotasXNotas.ToArray
                        _PesoBrutoRecuperado += dr.PesoFiscal
                    Next
                End If
            End If
            Return _NotasXNotas
        End Get
        Set(ByVal value As Negocio.NotasXNotas)
            _NotasXNotas = value
        End Set
    End Property

    Public Property PesoBrutoRecuperado() As Double
        Get
            Return _PesoBrutoRecuperado
        End Get
        Set(ByVal value As Double)
            _PesoBrutoRecuperado = value
        End Set
    End Property

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Razao() As Negocio.Razao
        Get
            If _Razao Is Nothing Then _Razao = New Negocio.Razao(Me)
            Return _Razao
        End Get
    End Property

    Public Property LancamentosContabeis As ListRazao
        Get
            If _LancamentosContabeis Is Nothing Then _LancamentosContabeis = New Negocio.ListRazao(Me, NFG)
            Return _LancamentosContabeis
        End Get
        Set(value As ListRazao)
            _LancamentosContabeis = value
        End Set
    End Property

    Public Property CFOP() As Negocio.CFOP
        Get
            If _Cfop Is Nothing And Not Itens Is Nothing AndAlso Itens.Count > 0 Then _Cfop = New CFOP(Itens(0).CFOP)
            Return _Cfop
        End Get
        Set(ByVal value As Negocio.CFOP)
            _Cfop = value
        End Set
    End Property

    Public ReadOnly Property NaturezaDaOperacao() As String
        Get
            If CFOP Is Nothing Then Return ""
            Return CFOP.Codigo & " - " & CFOP.Descricao
        End Get
    End Property

    Public Property CodigoTipoDeDocumento() As Integer
        Get
            Return _CodigoTipoDeDocumento
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoDeDocumento = value
        End Set
    End Property

    Public Property TipoDeDocumento() As Negocio.TipoDeDocumento
        Get
            If _TipoDeDocumento Is Nothing Then Return New Negocio.TipoDeDocumento(_CodigoTipoDeDocumento)
            Return _TipoDeDocumento
        End Get
        Set(ByVal value As Negocio.TipoDeDocumento)
            _TipoDeDocumento = value
        End Set
    End Property

    '********* UNIDADE DE NEGOCIO *********
    Public Property CodigoUnidadeDeNegocio As String
        Get
            Return _CodigoUnidadeDeNegocio
        End Get
        Set(value As String)
            _CodigoUnidadeDeNegocio = value
            _UnidadeDeNegocio = Nothing
        End Set
    End Property

    Public Property EndUnidadeDeNegocio As Integer
        Get
            Return _EndUnidadeDeNegocio
        End Get
        Set(value As Integer)
            _EndUnidadeDeNegocio = value
            _UnidadeDeNegocio = Nothing
        End Set
    End Property

    Public ReadOnly Property UnidadeDeNegocio As Negocio.Cliente
        Get
            If _UnidadeDeNegocio Is Nothing And _CodigoUnidadeDeNegocio.Length > 0 Then _UnidadeDeNegocio = New Negocio.Cliente(CodigoUnidadeDeNegocio, EndUnidadeDeNegocio)
            Return _UnidadeDeNegocio
        End Get
    End Property

    '*************** SAFRA ****************
    Public Property CodigoSafra As String
        Get
            Return _CodigoSafra
        End Get
        Set(value As String)
            _CodigoSafra = value
            _Safra = Nothing
        End Set
    End Property

    Public ReadOnly Property Safra As Negocio.Safra
        Get
            Return _Safra
        End Get
    End Property

    '*********  EMPRESA ************
    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
        End Set
    End Property

    Public Property Empresa() As Negocio.Cliente
        Get
            If _Empresa Is Nothing And Me.CodigoEmpresa.Trim.Length > 0 Then _Empresa = New Negocio.Cliente(Me.CodigoEmpresa, Me.EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Empresa = value
        End Set
    End Property

    '*********  CLIENTE ************
    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property Cliente() As Negocio.Cliente
        Get
            If _Cliente Is Nothing And Me.CodigoCliente.Trim.Length > 0 Then _Cliente = New Negocio.Cliente(Me.CodigoCliente, Me.EnderecoCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Cliente = value
        End Set
    End Property

    Public ReadOnly Property DescClienteCompleto As String
        Get
            If String.IsNullOrWhiteSpace(Me.CodigoCliente) Then Return ""
            Return Me.CodigoCliente & "-" & Me.EnderecoCliente & " : " & Me.Cliente.Nome
        End Get
    End Property

    '*********  NOTA ***************
    Public Property EntradaSaida() As Negocio.eEntradaSaida
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As Negocio.eEntradaSaida)
            _EntradaSaida = value
        End Set
    End Property

    Public Property Serie() As String
        Get
            Return _Serie
        End Get
        Set(ByVal value As String)
            _Serie = value.ToUpper
        End Set
    End Property

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Formulario() As Integer
        Get
            Return _Formulario
        End Get
        Set(ByVal value As Integer)
            _Formulario = value
        End Set
    End Property

    Public Property CodigoProcuracao() As Integer
        Get
            Return _CodigoProcuracao
        End Get
        Set(ByVal value As Integer)
            _CodigoProcuracao = value
        End Set
    End Property

    Public Property Procuracao() As Procuracao
        Get
            If _Procuracao Is Nothing And _CodigoProcuracao > 0 Then
                _Procuracao = New Negocio.Procuracao(_CodigoEmpresa, _EnderecoEmpresa, _CodigoProcuracao)
            End If

            Return _Procuracao
        End Get
        Set(ByVal value As Procuracao)
            _Procuracao = value
        End Set
    End Property

    Public Property ProcuracaoSaldoPedido() As Decimal
        Get
            Return _ProcuracaoSaldoPedido
        End Get
        Set(ByVal value As Decimal)
            _ProcuracaoSaldoPedido = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
            'If Not CarregandoNota Then AtualizaItens(3)
        End Set
    End Property

    Public Property DataNota() As DateTime
        Get
            Return _DataNota
        End Get
        Set(ByVal value As DateTime)
            _DataNota = value
            VerificarIndice()
            'If Not CarregandoNota Then AtualizaItens(1)
        End Set
    End Property

    Private ReadOnly Property MomentoIndice() As String
        Get
            Return _MomentoIndice
        End Get
    End Property

    Public Property NossaEmissao() As Boolean
        Get
            Return _NossaEmissao
        End Get
        Set(ByVal value As Boolean)
            _NossaEmissao = value
            If Not CarregandoNota Then AtualizaNumeradorNota()
        End Set
    End Property

    Public Property SerieNotaProdutor() As String
        Get
            Return _SerieNotaProdutor
        End Get
        Set(ByVal value As String)
            _SerieNotaProdutor = value
        End Set
    End Property

    Public Property NotaProdutor() As Integer
        Get
            Return _NotaProdutor
        End Get
        Set(ByVal value As Integer)
            _NotaProdutor = value
        End Set
    End Property

    '*********  FINALIDADE *************
    Public Property CodigoFinalidade() As Integer
        Get
            Return _CodigoFinalidade
        End Get
        Set(ByVal value As Integer)
            _CodigoFinalidade = value
            _Finalidade = Nothing
        End Set
    End Property

    Public Property Finalidade() As Negocio.Finalidade
        Get
            If _Finalidade Is Nothing And _CodigoFinalidade > 0 Then _Finalidade = New Negocio.Finalidade(_CodigoFinalidade)
            Return _Finalidade
        End Get
        Set(ByVal value As Negocio.Finalidade)
            _Finalidade = value
        End Set
    End Property

    '*********  PEDIDO *************
    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
            _Pedido = Nothing
        End Set
    End Property

    Public Property Pedido() As Negocio.Pedido
        Get
            If _Pedido Is Nothing And Me.CodigoPedido > 0 And Me.CodigoEmpresa.Trim.Length > 0 Then
                _Pedido = New Negocio.Pedido(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.CodigoPedido)
            End If
            Return _Pedido
        End Get
        Set(ByVal value As Negocio.Pedido)
            _Pedido = value
        End Set
    End Property

    '*********  OPERACAO ************
    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
            _Operacao = Nothing
            _SubOperacao = Nothing
        End Set
    End Property

    Public Property Operacao() As Negocio.Operacao
        Get
            If _Operacao Is Nothing And Me.CodigoOperacao > 0 Then _Operacao = New Negocio.Operacao(Me.CodigoOperacao)
            Return _Operacao
        End Get
        Set(ByVal value As Negocio.Operacao)
            _Operacao = value
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
            _SubOperacao = Nothing
            If Not CarregandoNota Then
                If NossaEmissao Then AtualizaNumeradorNota()

                _EntradaSaida = SubOperacao.EntradaSaida

                If Not Itens Is Nothing And Not NFG Then ' AndAlso Itens.Where(Function(s) s.IUD <> "D").Count > 0 Then
                    'AtualizaItens(2)
                    Itens.ForEach(Function(s)
                                      If s.IUD <> "D" Then
                                          s.IUD = "U"
                                          s.CodigoSubOperacao = value
                                      End If
                                      Return True
                                  End Function)
                End If
            End If
        End Set
    End Property

    Public Property SubOperacao() As Negocio.SubOperacao
        Get
            If _SubOperacao Is Nothing And Me.CodigoSubOperacao > 0 And Me.CodigoOperacao > 0 Then _SubOperacao = New Negocio.SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)
            Return _SubOperacao
        End Get
        Set(ByVal value As Negocio.SubOperacao)
            _SubOperacao = value
        End Set
    End Property

    Public ReadOnly Property DescSubOperacaoCompleto As String
        Get
            If Not Me.CodigoSubOperacao > 0 Then Return ""
            Return Me.CodigoOperacao & "-" & Me.CodigoSubOperacao & " : " & Me.SubOperacao.Descricao
        End Get
    End Property

    '*********  DEPOSITO ORIGEM ************
    Public Property CodigoDeposito() As String
        Get
            Return _CodigoDeposito
        End Get
        Set(ByVal value As String)
            _CodigoDeposito = value
        End Set
    End Property

    Public Property EnderecoDeposito() As Integer
        Get
            Return _EnderecoDeposito
        End Get
        Set(ByVal value As Integer)
            _EnderecoDeposito = value
        End Set
    End Property

    Public Property Deposito() As Negocio.Cliente
        Get
            If _Deposito Is Nothing And Me.CodigoDeposito.Trim.Length > 0 Then _Deposito = New Negocio.Cliente(Me.CodigoDeposito, Me.EnderecoDeposito)
            Return _Deposito
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Deposito = value
        End Set
    End Property

    Public ReadOnly Property DescDepositoCompleto As String
        Get
            If String.IsNullOrWhiteSpace(Me.CodigoDeposito) Then Return ""
            Return Me.CodigoDeposito & "-" & Me.EnderecoDeposito & " : " & Me.Deposito.Nome & " / " & Me.Deposito.CodigoEstado
        End Get
    End Property

    '********* DEPOSITO DESTINO ************
    Public Property CodigoDestino() As String
        Get
            Return _CodigoDestino
        End Get
        Set(ByVal value As String)
            _CodigoDestino = value
        End Set
    End Property

    Public Property EnderecoDestino() As Integer
        Get
            Return _EnderecoDestino
        End Get
        Set(ByVal value As Integer)
            _EnderecoDestino = value
        End Set
    End Property

    Public Property Destino() As Negocio.Cliente
        Get
            If _Destino Is Nothing And Me.CodigoDestino.Trim.Length > 0 Then _Destino = New Negocio.Cliente(Me.CodigoDestino, Me.EnderecoDestino)
            Return _Destino
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Destino = value
        End Set
    End Property

    '********* Local da Entrega ************
    Public Property CodigoEntrega() As String
        Get
            Return _CodigoEntrega
        End Get
        Set(ByVal value As String)
            _CodigoEntrega = value
        End Set
    End Property

    Public Property EnderecoEntrega() As Integer
        Get
            Return _EnderecoEntrega
        End Get
        Set(ByVal value As Integer)
            _EnderecoEntrega = value
        End Set
    End Property

    Public Property Entrega() As Negocio.Cliente
        Get
            If _Entrega Is Nothing And Me.CodigoEntrega.Trim.Length > 0 Then _Entrega = New Negocio.Cliente(Me.CodigoEntrega, Me.EnderecoEntrega)
            Return _Entrega
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Entrega = value
        End Set
    End Property

    Public ReadOnly Property DescDestinoCompleto As String
        Get
            If String.IsNullOrWhiteSpace(Me.CodigoDestino) Then Return ""
            Return Me.CodigoDestino & "-" & Me.EnderecoDestino & " : " & Me.Destino.Nome & " / " & Me.Destino.CodigoEstado
        End Get
    End Property

    '********* AGENCIADOR ************
    Public Property CodigoAgenciador() As String
        Get
            Return _CodigoAgenciador
        End Get
        Set(ByVal value As String)
            _CodigoAgenciador = value
        End Set
    End Property

    Public Property EnderecoAgenciador() As Integer
        Get
            Return _EnderecoAgenciador
        End Get
        Set(ByVal value As Integer)
            _EnderecoAgenciador = value
        End Set
    End Property

    Public Property Agenciador() As Negocio.Cliente
        Get
            If _Agenciador Is Nothing And Me.CodigoAgenciador.Trim.Length > 0 Then _Agenciador = New Negocio.Cliente(Me.CodigoAgenciador, Me.EnderecoAgenciador)
            Return _Agenciador
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Agenciador = value
        End Set
    End Property

    Public Property CIFFOB() As Negocio.eTiposFrete
        Get
            Return _CIFFOB
        End Get
        Set(ByVal value As Negocio.eTiposFrete)
            _CIFFOB = value
        End Set
    End Property

    Public ReadOnly Property TipoFreteSefaz As Integer
        Get
            If Me.CIFFOB = eTiposFrete.NEN Then
                Return 9
            ElseIf Me.CIFFOB = eTiposFrete.CIF Then
                Return 0
            ElseIf Me.CIFFOB = eTiposFrete.FOB Then
                Return 1
            ElseIf Me.CIFFOB = eTiposFrete.TER Then
                Return 2
            Else
                Return 9
            End If
        End Get
    End Property

    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
        End Set
    End Property

    Public Property DataInclusao() As DateTime
        Get
            Return _DataInclusao
        End Get
        Set(ByVal value As DateTime)
            _DataInclusao = value
        End Set
    End Property

    Public Property UsuarioAlteracao() As String
        Get
            Return _UsuarioAlteracao
        End Get
        Set(ByVal value As String)
            _UsuarioAlteracao = value
        End Set
    End Property

    Public Property DataAlteracao() As DateTime
        Get
            Return _DataAlteracao
        End Get
        Set(ByVal value As DateTime)
            _DataAlteracao = value
        End Set
    End Property

    Public Property CodigoSituacao() As Integer
        Get
            Return _CodigoSituacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSituacao = value
        End Set
    End Property

    Public Property Situacao() As Negocio.Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Negocio.Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Negocio.Situacao)
            _Situacao = value
        End Set
    End Property

    Public Property Itens() As ListNotaFiscalXItem
        Get
            If _Itens Is Nothing And CarregandoNota Then
                _Itens = New ListNotaFiscalXItem(Me)
            ElseIf _Itens Is Nothing And Not CarregandoNota Then
                _Itens = New ListNotaFiscalXItem(Me, False)
            End If
            Return _Itens
        End Get
        Set(ByVal value As ListNotaFiscalXItem)
            _Itens = value
        End Set
    End Property

    '******************  VENCIMENTOS *********************
    Public Property VencimentosPedido() As Negocio.ListTitulo
        Get
            If _VencimentosPedido Is Nothing AndAlso _CodigoTipoDeDocumento = 1 AndAlso _CodigoPedido > 0 Then _VencimentosPedido = New Negocio.ListTitulo(_CodigoPedido, Pedido.SubOperacao.EntradaSaida.ToString.Substring(0, 1), Pedido.CodigoEmpresa, Pedido.EnderecoEmpresa)
            Return _VencimentosPedido
        End Get
        Set(ByVal value As Negocio.ListTitulo)
            _VencimentosPedido = value
        End Set
    End Property

    Public Property VencimentosNota() As Negocio.ListTitulo
        Get
            If _VencimentosNota Is Nothing Then _VencimentosNota = New Negocio.ListTitulo(Me)
            Return _VencimentosNota
        End Get
        Set(ByVal value As Negocio.ListTitulo)
            _VencimentosNota = value
        End Set
    End Property

    Public Property TitulosPedido() As Novo.ListTituloNovo
        Get
            If _TitulosPedido Is Nothing Then
                _TitulosPedido = New Novo.ListTituloNovo(Me.Pedido.CodigoEmpresa, Me.Pedido.EnderecoEmpresa, Me.Pedido.Codigo)
            End If
            Return _TitulosPedido
        End Get
        Set(ByVal value As Novo.ListTituloNovo)
            _TitulosPedido = value
        End Set
    End Property

    Public Property Titulos() As Novo.ListTituloNovo
        Get
            If _Titulos Is Nothing Then
                _Titulos = New Novo.ListTituloNovo(Me)
            End If
            Return _Titulos
        End Get
        Set(ByVal value As Novo.ListTituloNovo)
            _Titulos = value
        End Set
    End Property

    '***************** RESUMOS *****************************
    'Bruto
    Public Property TotalProduto() As Decimal
        Get
            Return _TotalProduto
        End Get
        Set(ByVal value As Decimal)
            _TotalProduto = value
        End Set
    End Property

    Public Property TotalProdutoMoeda() As Decimal
        Get
            Return _TotalProdutoMoeda
        End Get
        Set(ByVal value As Decimal)
            _TotalProdutoMoeda = value
        End Set
    End Property

    'Liquido
    Public Property TotalNota() As Decimal
        Get
            Return _TotalNota
        End Get
        Set(ByVal value As Decimal)
            _TotalNota = value
        End Set
    End Property

    Public Property TotalNotaMoeda() As Decimal
        Get
            Return _TotalNotaMoeda
        End Get
        Set(ByVal value As Decimal)
            _TotalNotaMoeda = value
        End Set
    End Property

    Public ReadOnly Property IndiceNota
        Get
            If (Pedido Is Nothing OrElse Pedido.Moeda Is Nothing OrElse Pedido.Moeda.Classificacao = eTiposMoeda.Oficial) OrElse (Me.TotalQuantidadeFiscal <= 0) Then
                Return 0
            Else
                If Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira AndAlso (Pedido.CodigoIndexador = 99 OrElse Pedido.IndexadorFixo) Then
                    Return Pedido.IndiceFixado
                ElseIf Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                    Dim indiceDolar As Decimal = New Cotacao(Pedido.CodigoIndexador, DataNota).Indice 'DÓLAR DO INDEXADOR PEDIDO
                    Return indiceDolar
                Else
                    Return Me.Itens.ValorBruto_Oficial / Me.Itens.ValorBruto_Moeda
                End If

            End If
        End Get
    End Property

    'Demais
    Public Property ValorBaseIcms() As Decimal
        Get
            Return _ValorBaseIcms
        End Get
        Set(ByVal value As Decimal)
            _ValorBaseIcms = value
        End Set
    End Property

    Public Property ValorIcms() As Decimal
        Get
            Return _ValorIcms
        End Get
        Set(ByVal value As Decimal)
            _ValorIcms = value
        End Set
    End Property

    Public Property ValorBaseIcmsST() As Decimal
        Get
            Return _ValorBaseIcmsST
        End Get
        Set(ByVal value As Decimal)
            _ValorBaseIcmsST = value
        End Set
    End Property

    Public Property ValorIcmsST() As Decimal
        Get
            Return _ValorIcmsST
        End Get
        Set(ByVal value As Decimal)
            _ValorIcmsST = value
        End Set
    End Property

    Public Property ValorIPI() As Decimal
        Get
            Return _ValorIPI
        End Get
        Set(ByVal value As Decimal)
            _ValorIPI = value
        End Set
    End Property

    Public Property ValorFrete() As Decimal
        Get
            Return _ValorFrete
        End Get
        Set(ByVal value As Decimal)
            _ValorFrete = value
        End Set
    End Property

    Public Property ValorDesconto() As Decimal
        Get
            Return _ValorDesconto
        End Get
        Set(ByVal value As Decimal)
            _ValorDesconto = value
        End Set
    End Property

    Public Property ValorAduaneira() As Decimal
        Get
            Return _ValorAduaneira
        End Get
        Set(ByVal value As Decimal)
            _ValorAduaneira = value
        End Set
    End Property

    Public Property ValorSeguro() As Decimal
        Get
            Return _ValorSeguro
        End Get
        Set(ByVal value As Decimal)
            _ValorSeguro = value
        End Set
    End Property

    '********* Resumo Item **************
    Public ReadOnly Property TotalNotaBruto
        Get
            Return Me.Itens.Sum(Function(s) s.ValorTotal)
        End Get
    End Property

    Public ReadOnly Property TotalQuantidadeFiscal
        Get
            Return Me.Itens.Sum(Function(s) s.QuantidadeFiscal)
        End Get
    End Property

    Public ReadOnly Property TotalQuantidadeFisica
        Get
            Return Me.Itens.Sum(Function(s) s.QuantidadeFisica)
        End Get
    End Property

    '*********  TRANSPORTADOR ************
    Public Property CodigoTransportador() As String
        Get
            Return _CodigoTransportador
        End Get
        Set(ByVal value As String)
            _CodigoTransportador = value
        End Set
    End Property

    Public Property EnderecoTransportador() As Integer
        Get
            Return _EnderecoTransportador
        End Get
        Set(ByVal value As Integer)
            _EnderecoTransportador = value
        End Set
    End Property

    Public Property Transportador() As Negocio.Cliente
        Get
            If _Transportador Is Nothing And Me.CodigoTransportador.Trim.Length > 0 Then _Transportador = New Negocio.Cliente(Me.CodigoTransportador, Me.EnderecoTransportador)
            Return _Transportador
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Transportador = value
        End Set
    End Property

    Public Property PlacaTransportador() As String
        Get
            Return _PlacaTransportador
        End Get
        Set(ByVal value As String)
            _PlacaTransportador = value
        End Set
    End Property

    Public Property PlacaDetalhes() As Negocio.Placa
        Get
            If (_PlacaDetalhes Is Nothing OrElse String.IsNullOrWhiteSpace(_PlacaDetalhes.Placa01)) AndAlso Not String.IsNullOrWhiteSpace(_PlacaTransportador) Then
                _PlacaDetalhes = New Negocio.Placa(_PlacaTransportador)
            End If
            Return _PlacaDetalhes
        End Get
        Set(ByVal value As Negocio.Placa)
            _PlacaDetalhes = value
        End Set
    End Property

    '*********  TRANSBORDO ************
    Public Property CodigoTransbordo() As String
        Get
            Return _CodigoTransbordo
        End Get
        Set(ByVal value As String)
            _CodigoTransbordo = value
        End Set
    End Property

    Public Property EnderecoTransbordo() As Integer
        Get
            Return _EnderecoTransbordo
        End Get
        Set(ByVal value As Integer)
            _EnderecoTransbordo = value
        End Set
    End Property

    Public Property Transbordo() As Negocio.Cliente
        Get
            If _Transbordo Is Nothing And _CodigoTransbordo.Trim.Length > 0 Then _Transbordo = New Negocio.Cliente(_CodigoTransbordo, _EnderecoTransbordo)
            Return _Transbordo
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Transbordo = value
        End Set
    End Property

    '*********  MOTORISTA ************
    Public Property CodigoMotorista() As String
        Get
            Return _CodigoMotorista
        End Get
        Set(ByVal value As String)
            _CodigoMotorista = value
        End Set
    End Property

    Public Property EnderecoMotorista() As Integer
        Get
            Return _EnderecoMotorista
        End Get
        Set(ByVal value As Integer)
            _EnderecoMotorista = value
        End Set
    End Property

    Public Property Motorista() As Negocio.Cliente
        Get
            If _Motorista Is Nothing And _CodigoMotorista.Trim.Length > 0 Then _Motorista = New Negocio.Cliente(_CodigoMotorista, _EnderecoMotorista)
            Return _Motorista
        End Get
        Set(ByVal value As Negocio.Cliente)
            _Motorista = value
        End Set
    End Property


    '*********  PROPRIETéRIO DA MERCADORIA ************
    Public Property CodigoProprietarioDaMercadoria() As String
        Get
            Return _CodigoProprietarioDaMercadoria
        End Get
        Set(ByVal value As String)
            _CodigoProprietarioDaMercadoria = value
        End Set
    End Property

    Public Property EnderecoProprietarioDaMercadoria() As Integer
        Get
            Return _EnderecoProprietarioDaMercadoria
        End Get
        Set(ByVal value As Integer)
            _EnderecoProprietarioDaMercadoria = value
        End Set
    End Property

    Public Property ProprietarioDaMercadoria() As Negocio.Cliente
        Get
            If _ProprietarioDaMercadoria Is Nothing And _CodigoProprietarioDaMercadoria.Trim.Length > 0 Then _ProprietarioDaMercadoria = New Negocio.Cliente(_CodigoProprietarioDaMercadoria, _EnderecoProprietarioDaMercadoria)
            Return _ProprietarioDaMercadoria
        End Get
        Set(ByVal value As Negocio.Cliente)
            _ProprietarioDaMercadoria = value
        End Set
    End Property


    '******** OBSERVACOES  ************
    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property ObservacoesDoProduto() As String
        Get
            Return _ObservacoesDoProduto
        End Get
        Set(ByVal value As String)
            _ObservacoesDoProduto = value
        End Set
    End Property

    Public Property ObservacoesDeEmbarque() As String
        Get
            Return _ObservacoesDeEmbarque
        End Get
        Set(ByVal value As String)
            _ObservacoesDeEmbarque = value
        End Set
    End Property

    Public Property ObservacoesControleInterno() As String
        Get
            Return _ObservacoesControleInterno
        End Get
        Set(ByVal value As String)
            _ObservacoesControleInterno = value
        End Set
    End Property

    '******** PESOS - NUMERACAO - VOLUMES  ************
    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Numero() As String
        Get
            Return _Numero
        End Get
        Set(ByVal value As String)
            _Numero = value
        End Set
    End Property

    Public Property Especie() As String
        Get
            Return _Especie
        End Get
        Set(ByVal value As String)
            _Especie = value
        End Set
    End Property

    Public Property Marca() As String
        Get
            Return _Marca
        End Get
        Set(ByVal value As String)
            _Marca = value
        End Set
    End Property

    Public Property PesoBruto() As Decimal
        Get
            Return _PesoBruto
        End Get
        Set(ByVal value As Decimal)
            _PesoBruto = value
        End Set
    End Property

    Public Property PesoLiquido() As Decimal
        Get
            Return _PesoLiquido
        End Get
        Set(ByVal value As Decimal)
            _PesoLiquido = value
        End Set
    End Property

    '***************** NFE  **************************
    Public Property Eletronica() As Boolean
        Get
            Return _Eletronica
        End Get
        Set(ByVal value As Boolean)
            _Eletronica = value
        End Set
    End Property

    Public Property ChaveNFE() As String
        Get
            Return _ChaveNFE
        End Get
        Set(ByVal value As String)
            _ChaveNFE = value
        End Set
    End Property

    Public Property TemSegCodBarra() As Boolean
        Get
            Return _TemSegCodBarra
        End Get
        Set(ByVal value As Boolean)
            _TemSegCodBarra = value
        End Set
    End Property
    Public Property SegCodBarra() As String
        Get
            Return _SegCodBarra
        End Get
        Set(ByVal value As String)
            _SegCodBarra = value
        End Set
    End Property

    Public Property ProtocoloNota() As String
        Get
            Return _ProtocoloNota
        End Get
        Set(ByVal value As String)
            _ProtocoloNota = value
        End Set
    End Property

    Public Property ReciboNota() As String
        Get
            Return _ReciboNota
        End Get
        Set(ByVal value As String)
            _ReciboNota = value
        End Set
    End Property

    Public Property Retorno() As String
        Get
            Return _Retorno
        End Get
        Set(ByVal value As String)
            _Retorno = value
        End Set
    End Property

    Public Property MsgRetorno() As String
        Get
            Return _MsgRetorno
        End Get
        Set(ByVal value As String)
            _MsgRetorno = value
        End Set
    End Property

    Public Property ObservacaoCancelamento() As String
        Get
            Return _ObservacaoCancelamento
        End Get
        Set(ByVal value As String)
            _ObservacaoCancelamento = value
        End Set
    End Property

    Public Property Usuario() As String
        Get
            Return _Usuario
        End Get
        Set(ByVal value As String)
            _Usuario = value
        End Set
    End Property

    Public Property DataHoraNFE() As DateTime?
        Get
            Return _DataHoraNFE
        End Get
        Set(ByVal value As DateTime?)
            _DataHoraNFE = value
        End Set
    End Property

    '****************  TROCA DE NOTA  *****************
    Public ReadOnly Property TemNotaTroca() As Boolean
        Get
            If NotaTrocaOrigem Is Nothing And NotaTrocaDestino Is Nothing Then
                Return False
            Else
                Return True
            End If
        End Get
    End Property

    Public ReadOnly Property NotaDeTroca() As Negocio.NotaFiscal
        Get
            Return _NotaDeTroca
        End Get
    End Property

    Public Property NotasTrocaOrigem() As List(Of Negocio.NotaFiscal)
        Get
            If _NotasTrocaOrigem Is Nothing Then
                Dim Banco As New AcessaBanco
                Dim ds As DataSet
                Dim sql As String = ""
                sql = "	SELECT NxN.OrigemEmpresa_Id," & vbCrLf &
                      "        NxN.OrigemEndEmpresa_Id," & vbCrLf &
                      "        NxN.OrigemCliente_Id," & vbCrLf &
                      "        NxN.OrigemEndCliente_Id," & vbCrLf &
                      "        NxN.OrigemEntradaSaida_Id," & vbCrLf &
                      "        NxN.OrigemSerie_Id," & vbCrLf &
                      "        NxN.OrigemNota_Id" & vbCrLf &
                      "	  FROM NotasXNotas NxN " & vbCrLf &
                      "  INNER JOIN NotasFiscais NF" & vbCrLf &
                      "     ON NF.Empresa_Id      = NxN.OrigemEmpresa_Id" & vbCrLf &
                      "    AND NF.EndEmpresa_Id   = NxN.OrigemEndEmpresa_Id" & vbCrLf &
                      "    AND NF.Cliente_id      = NxN.OrigemCliente_Id" & vbCrLf &
                      "    AND NF.EndCliente_id   = NxN.OrigemEndCliente_Id" & vbCrLf &
                      "    AND NF.EntradaSaida_Id = NxN.OrigemEntradaSaida_Id" & vbCrLf &
                      "    AND NF.Serie_Id        = NxN.OrigemSerie_Id" & vbCrLf &
                      "    AND NF.Nota_Id         = NxN.OrigemNota_Id" & vbCrLf &
                      "	 WHERE NxN.Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "	   AND NxN.EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                      "	   AND NxN.Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                      "	   AND NxN.EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                      "	   AND NxN.EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "	   AND NxN.Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                      "	   AND NxN.Nota_Id         = " & Me.Codigo & vbCrLf
                ds = Banco.ConsultaDataSet(sql, "NxN")

                If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                    For Each row As DataRow In ds.Tables(0).Rows
                        Dim nf As New NotaFiscal
                        nf.CodigoEmpresa = row("OrigemEmpresa_Id")
                        nf.EnderecoEmpresa = row("OrigemEndEmpresa_Id")
                        nf.CodigoCliente = row("OrigemCliente_Id")
                        nf.EnderecoCliente = row("OrigemEndCliente_Id")
                        nf.EntradaSaida = IIf(row("OrigemEntradaSaida_Id") = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                        nf.Serie = row("OrigemSerie_Id")
                        nf.Codigo = row("OrigemNota_Id")
                        If _NotasTrocaOrigem Is Nothing Then
                            _NotasTrocaOrigem = New List(Of Negocio.NotaFiscal)
                        End If
                        _NotasTrocaOrigem.Add(New NotaFiscal(nf))
                        _NotaTrocaDestino = Me
                    Next
                End If
            End If
            Return _NotasTrocaOrigem
        End Get
        Set(ByVal value As List(Of Negocio.NotaFiscal))
            _NotasTrocaOrigem = value
            _NotaTrocaDestino = Me
        End Set
    End Property

    Public Property NotaTrocaOrigem() As Negocio.NotaFiscal
        Get
            If _NotaTrocaOrigem Is Nothing AndAlso Not Cliente Is Nothing Then
                Dim ds As DataSet
                Dim db As New AcessaBanco
                Dim sql As String = ""
                sql = "	SELECT TOP 1 NxN.OrigemEmpresa_Id," & vbCrLf &
                      "        NxN.OrigemEndEmpresa_Id," & vbCrLf &
                      "        NxN.OrigemCliente_Id," & vbCrLf &
                      "        NxN.OrigemEndCliente_Id," & vbCrLf &
                      "        NxN.OrigemEntradaSaida_Id," & vbCrLf &
                      "        NxN.OrigemSerie_Id," & vbCrLf &
                      "        NxN.OrigemNota_Id" & vbCrLf &
                      "	  FROM NotasXNotas NxN " & vbCrLf &
                      "  Inner Join NotasFiscais NF" & vbCrLf &
                      "     on NF.Empresa_Id      = NxN.OrigemEmpresa_Id" & vbCrLf &
                      "    and NF.EndEmpresa_Id   = NxN.OrigemEndEmpresa_Id" & vbCrLf &
                      "    and NF.Cliente_id      = NxN.OrigemCliente_Id" & vbCrLf &
                      "    and NF.EndCliente_id   = NxN.OrigemEndCliente_Id" & vbCrLf &
                      "    and NF.EntradaSaida_Id = NxN.OrigemEntradaSaida_Id" & vbCrLf &
                      "    and NF.Serie_Id        = NxN.OrigemSerie_Id" & vbCrLf &
                      "    and NF.Nota_Id         = NxN.OrigemNota_Id" & vbCrLf &
                      "	 Where NxN.Empresa_Id      ='" & Empresa.Codigo & "'" & vbCrLf &
                      "	   and NxN.EndEmpresa_Id   = " & Empresa.CodigoEndereco & vbCrLf &
                      "	   and NxN.Cliente_Id      ='" & Cliente.Codigo & "'" & vbCrLf &
                      "	   and NxN.EndCliente_Id   = " & Cliente.CodigoEndereco & vbCrLf &
                      "	   and NxN.EntradaSaida_Id ='" & EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "	   and NxN.Serie_Id        ='" & Serie & "'" & vbCrLf &
                      "	   and NxN.Nota_Id         = " & Codigo & vbCrLf &
                      "    and NF.TipoDeDocumento IN (1,2,8, 15, 57)"

                ds = db.ConsultaDataSet(sql, "NxN")

                If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                    Dim row As DataRow = ds.Tables(0).Rows(0)
                    Dim NotaCons As New NotaFiscal
                    NotaCons.CodigoEmpresa = row("OrigemEmpresa_Id")
                    NotaCons.EnderecoEmpresa = row("OrigemEndEmpresa_Id")
                    NotaCons.CodigoCliente = row("OrigemCliente_Id")
                    NotaCons.EnderecoCliente = row("OrigemEndCliente_Id")
                    NotaCons.EntradaSaida = IIf(row("OrigemEntradaSaida_Id") = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                    NotaCons.Serie = row("OrigemSerie_Id")
                    NotaCons.Codigo = row("OrigemNota_Id")
                    _NotaTrocaOrigem = New NotaFiscal(NotaCons)
                    _NotaTrocaDestino = Me
                    _NotaDeTroca = _NotaTrocaOrigem
                End If
            End If
            Return _NotaTrocaOrigem
        End Get
        Set(ByVal value As Negocio.NotaFiscal)
            _NotaTrocaOrigem = value
            _NotaTrocaDestino = Me
            _NotaDeTroca = _NotaTrocaOrigem
        End Set
    End Property

    Public Property NotaTrocaDestino() As Negocio.NotaFiscal
        Get
            If _NotaTrocaDestino Is Nothing Then
                Dim Banco As New AcessaBanco
                Dim ds As DataSet
                Dim sql As String = ""
                sql = "	SELECT NxN.Empresa_Id," & vbCrLf &
                      "        NxN.EndEmpresa_Id," & vbCrLf &
                      "        NxN.Cliente_Id," & vbCrLf &
                      "        NxN.EndCliente_Id," & vbCrLf &
                      "        NxN.EntradaSaida_Id," & vbCrLf &
                      "        NxN.Serie_Id," & vbCrLf &
                      "        NxN.Nota_Id" & vbCrLf &
                      "	  FROM NotasXNotas NxN " & vbCrLf &
                      "  Inner Join NotasFiscais NF" & vbCrLf &
                      "     on NF.Empresa_Id      = NxN.Empresa_Id" & vbCrLf &
                      "    and NF.EndEmpresa_Id   = NxN.EndEmpresa_Id" & vbCrLf &
                      "    and NF.Cliente_id      = NxN.Cliente_Id" & vbCrLf &
                      "    and NF.EndCliente_id   = NxN.EndCliente_Id" & vbCrLf &
                      "    and NF.EntradaSaida_Id = NxN.EntradaSaida_Id" & vbCrLf &
                      "    and NF.Serie_Id        = NxN.Serie_Id" & vbCrLf &
                      "    and NF.Nota_Id         = NxN.Nota_Id" & vbCrLf &
                      "	 Where NxN.OrigemEmpresa_Id      ='" & Empresa.Codigo & "'" & vbCrLf &
                      "	   and NxN.OrigemEndEmpresa_Id   = " & Empresa.CodigoEndereco & vbCrLf &
                      "	   and NxN.OrigemCliente_Id      ='" & Cliente.Codigo & "'" & vbCrLf &
                      "	   and NxN.OrigemEndCliente_Id   = " & Cliente.CodigoEndereco & vbCrLf &
                      "	   and NxN.OrigemEntradaSaida_Id ='" & EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "	   and NxN.OrigemSerie_Id        ='" & Serie & "'" & vbCrLf &
                      "	   and NxN.OrigemNota_Id         = " & Codigo & vbCrLf &
                      "    and NF.TipoDeDocumento        = 1"
                ds = Banco.ConsultaDataSet(sql, "NxN")
                If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                    Dim row As DataRow = ds.Tables(0).Rows(0)
                    Dim NotaCons As New NotaFiscal
                    NotaCons.CodigoEmpresa = row("Empresa_Id")
                    NotaCons.EnderecoEmpresa = row("EndEmpresa_Id")
                    NotaCons.CodigoCliente = row("Cliente_Id")
                    NotaCons.EnderecoCliente = row("EndCliente_Id")
                    NotaCons.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                    NotaCons.Serie = row("Serie_Id")
                    NotaCons.Codigo = row("Nota_Id")
                    _NotaTrocaDestino = New NotaFiscal(NotaCons)
                    _NotaTrocaOrigem = Me
                    _NotaDeTroca = _NotaTrocaOrigem
                End If
            End If
            Return _NotaTrocaDestino
        End Get
        Set(ByVal value As Negocio.NotaFiscal)
            _NotaTrocaDestino = value
            _NotaTrocaOrigem = Me
            _NotaDeTroca = _NotaTrocaDestino
        End Set
    End Property


    Public ReadOnly Property NotasOrigemDestino() As List(Of Negocio.NotaFiscal)
        Get
            Dim Notas As List(Of Negocio.NotaFiscal)
            Dim Banco As New AcessaBanco
            Dim ds As DataSet
            Dim sql As String = ""
            sql = "	SELECT NxN.OrigemEmpresa_Id," & vbCrLf &
                  "        NxN.OrigemEndEmpresa_Id," & vbCrLf &
                  "        NxN.OrigemCliente_Id," & vbCrLf &
                  "        NxN.OrigemEndCliente_Id," & vbCrLf &
                  "        NxN.OrigemEntradaSaida_Id," & vbCrLf &
                  "        NxN.OrigemSerie_Id," & vbCrLf &
                  "        NxN.OrigemNota_Id" & vbCrLf &
                  "   FROM (" & vbCrLf &
                  "		     --ORIGEM" & vbCrLf &
                  "		     SELECT NxN.OrigemEmpresa_Id," & vbCrLf &
                  "		            NxN.OrigemEndEmpresa_Id," & vbCrLf &
                  "		            NxN.OrigemCliente_Id," & vbCrLf &
                  "		            NxN.OrigemEndCliente_Id," & vbCrLf &
                  "		            NxN.OrigemEntradaSaida_Id," & vbCrLf &
                  "		            NxN.OrigemSerie_Id," & vbCrLf &
                  "		            NxN.OrigemNota_Id," & vbCrLf &
                  "		            NxN.Empresa_Id," & vbCrLf &
                  "		            NxN.EndEmpresa_Id," & vbCrLf &
                  "		            NxN.Cliente_Id," & vbCrLf &
                  "		            NxN.EndCliente_Id," & vbCrLf &
                  "		            NxN.EntradaSaida_Id," & vbCrLf &
                  "		            NxN.Serie_Id," & vbCrLf &
                  "		            NxN.Nota_Id" & vbCrLf &
                  "		       FROM NotasXNotas NxN " & vbCrLf &
                  "		      UNION ALL" & vbCrLf &
                  "		      --DESTINO" & vbCrLf &
                  "		     SELECT NxN.Empresa_Id," & vbCrLf &
                  "		            NxN.EndEmpresa_Id," & vbCrLf &
                  "		            NxN.Cliente_Id," & vbCrLf &
                  "		            NxN.EndCliente_Id," & vbCrLf &
                  "		            NxN.EntradaSaida_Id," & vbCrLf &
                  "		            NxN.Serie_Id," & vbCrLf &
                  "		            NxN.Nota_Id," & vbCrLf &
                  "		            NxN.OrigemEmpresa_Id," & vbCrLf &
                  "		            NxN.OrigemEndEmpresa_Id," & vbCrLf &
                  "		            NxN.OrigemCliente_Id," & vbCrLf &
                  "		            NxN.OrigemEndCliente_Id," & vbCrLf &
                  "		            NxN.OrigemEntradaSaida_Id," & vbCrLf &
                  "		            NxN.OrigemSerie_Id," & vbCrLf &
                  "		            NxN.OrigemNota_Id" & vbCrLf &
                  "		       FROM NotasXNotas NxN " & vbCrLf &
                  "	       ) AS NxN " & vbCrLf &
                  "  INNER JOIN NotasFiscais NF" & vbCrLf &
                  "     ON NF.Empresa_Id      = NxN.OrigemEmpresa_Id" & vbCrLf &
                  "    AND NF.EndEmpresa_Id   = NxN.OrigemEndEmpresa_Id" & vbCrLf &
                  "    AND NF.Cliente_id      = NxN.OrigemCliente_Id" & vbCrLf &
                  "    AND NF.EndCliente_id   = NxN.OrigemEndCliente_Id" & vbCrLf &
                  "    AND NF.EntradaSaida_Id = NxN.OrigemEntradaSaida_Id" & vbCrLf &
                  "    AND NF.Serie_Id        = NxN.OrigemSerie_Id" & vbCrLf &
                  "    AND NF.Nota_Id         = NxN.OrigemNota_Id" & vbCrLf &
                  "	 WHERE NxN.Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                  "	   AND NxN.EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                  "	   AND NxN.Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                  "	   AND NxN.EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                  "	   AND NxN.EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "	   AND NxN.Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                  "	   AND NxN.Nota_Id         = " & Me.Codigo & vbCrLf
            ds = Banco.ConsultaDataSet(sql, "NxN")

            Notas = New List(Of NotaFiscal)
            If Not ds Is Nothing AndAlso ds.Tables(0).Rows.Count > 0 Then
                For Each row As DataRow In ds.Tables(0).Rows
                    Dim nf As New NotaFiscal
                    nf.CodigoEmpresa = row("OrigemEmpresa_Id")
                    nf.EnderecoEmpresa = row("OrigemEndEmpresa_Id")
                    nf.CodigoCliente = row("OrigemCliente_Id")
                    nf.EnderecoCliente = row("OrigemEndCliente_Id")
                    nf.EntradaSaida = IIf(row("OrigemEntradaSaida_Id") = "E", Negocio.eEntradaSaida.Entrada, Negocio.eEntradaSaida.Saida)
                    nf.Serie = row("OrigemSerie_Id")
                    nf.Codigo = row("OrigemNota_Id")

                    Notas.Add(New NotaFiscal(nf))

                Next
            End If
            Return Notas
        End Get
    End Property



    '******************  ROMANEIO  ********************
    Public Property CriarRomaneio() As Boolean
        Get
            Return _CriarRomaneio
        End Get
        Set(ByVal value As Boolean)
            _CriarRomaneio = value
        End Set
    End Property

    Public Property CodigoRomaneio() As Integer
        Get
            Return _CodigoRomaneio
        End Get
        Set(ByVal value As Integer)
            _CodigoRomaneio = value
        End Set
    End Property

    Public Property Romaneio() As Negocio.Romaneio
        Get
            If _Romaneio Is Nothing And _CodigoRomaneio > 0 Then _Romaneio = New Negocio.Romaneio(_CodigoEmpresa, _EnderecoEmpresa, _CodigoRomaneio)
            Return _Romaneio
        End Get
        Set(ByVal value As Negocio.Romaneio)
            _Romaneio = value
        End Set
    End Property

    '************  AUTORIZACAO DE RETIRADA  ***********
    Public Property Autorizacao() As Negocio.AutorizacaoDeRetirada
        Get
            If _Autorizacao Is Nothing And _CodigoAutorizacao > 0 Then _Autorizacao = New Negocio.AutorizacaoDeRetirada(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.CodigoPedido, Me.CodigoAutorizacao, Me.SubOperacao.Classe)
            Return _Autorizacao
        End Get
        Set(ByVal value As Negocio.AutorizacaoDeRetirada)
            _Autorizacao = value
        End Set
    End Property

    Public Property CodigoAutorizacao() As Integer
        Get
            Return _CodigoAutorizacao
        End Get
        Set(ByVal value As Integer)
            _CodigoAutorizacao = value
            _Autorizacao = Nothing
        End Set
    End Property

    '************  LOCAL DE EMBARQUE  ***********
    Public Property CodigoLocalEmbarque() As String
        Get
            Return _CodigoLocalEmbarque
        End Get
        Set(ByVal value As String)
            _CodigoLocalEmbarque = value
            _LocalEmbarque = Nothing
        End Set
    End Property

    Public Property EndLocalEmbarque() As Integer
        Get
            Return _EndLocalEmbarque
        End Get
        Set(ByVal value As Integer)
            _EndLocalEmbarque = value
            _LocalEmbarque = Nothing
        End Set
    End Property

    Public ReadOnly Property LocalEmbarque() As Negocio.Cliente
        Get
            If _LocalEmbarque Is Nothing And _CodigoLocalEmbarque.Length > 0 Then _LocalEmbarque = New Negocio.Cliente(_CodigoLocalEmbarque, _EndLocalEmbarque)
            Return _LocalEmbarque
        End Get
    End Property

    '************  PAMCARD  ***********
    Public Property ContratoANTT() As String
        Get
            Return _ContratoANTT
        End Get
        Set(ByVal value As String)
            _ContratoANTT = value
        End Set
    End Property

    Public Property ProtocoloANTT() As String
        Get
            Return _ProtocoloANTT
        End Get
        Set(ByVal value As String)
            _ProtocoloANTT = value
        End Set
    End Property

    Public Property idPamcard() As String
        Get
            Return _idPamcard
        End Get
        Set(ByVal value As String)
            _idPamcard = value
        End Set
    End Property

    Public Property CartaoPgtoFrete() As String
        Get
            Return _CartaoPgtoFrete
        End Get
        Set(ByVal value As String)
            _CartaoPgtoFrete = value
        End Set
    End Property

    Public Property DataTermino() As DateTime
        Get
            Return _DataTermino
        End Get
        Set(ByVal value As DateTime)
            _DataTermino = value
        End Set
    End Property

    'EXPORTAção
    Public Property DadosDaExportacao() As Negocio.NotaFiscalXExportacao
        Get
            If _DadosDaExportacao Is Nothing Then
                _DadosDaExportacao = New Negocio.NotaFiscalXExportacao(Me)
                If _DadosDaExportacao.IUD = "N" Then _DadosDaExportacao = Nothing
            End If
            Return _DadosDaExportacao
        End Get
        Set(ByVal value As Negocio.NotaFiscalXExportacao)
            _DadosDaExportacao = value
        End Set
    End Property

    Public Property DadosDaExportacaoRE() As ListNotaFiscalXRE
        Get
            If _DadosDaExportacaoRE Is Nothing Then
                _DadosDaExportacaoRE = New ListNotaFiscalXRE(Me)
            End If
            Return _DadosDaExportacaoRE
        End Get
        Set(ByVal value As ListNotaFiscalXRE)
            _DadosDaExportacaoRE = value
        End Set
    End Property

    Public Property DadosDaExportacaoCE() As ListNotaFiscalXCE
        Get
            If _DadosDaExportacaoCE Is Nothing Then
                _DadosDaExportacaoCE = New ListNotaFiscalXCE(Me)
            End If
            Return _DadosDaExportacaoCE
        End Get
        Set(ByVal value As ListNotaFiscalXCE)
            _DadosDaExportacaoCE = value
        End Set
    End Property

    'IMPORTAção
    Public Property DadosDaImportacao() As Negocio.NotaFiscalXImportacao
        Get
            If _DadosDaImportacao Is Nothing Then
                _DadosDaImportacao = New Negocio.NotaFiscalXImportacao(Me)
                If _DadosDaImportacao.IUD = "N" Then _DadosDaImportacao = Nothing
            End If
            Return _DadosDaImportacao
        End Get
        Set(ByVal value As Negocio.NotaFiscalXImportacao)
            _DadosDaImportacao = value
        End Set
    End Property

    'FRETE
    Public Property CodigoTipoDeDocumentoFrete() As Integer
        Get
            If Me.TipoDeDocumentoFrete IsNot Nothing Then
                Return Convert.ToInt32(Me.TipoDeDocumentoFrete)
            End If
            Return Nothing
        End Get
        Set(ByVal value As Integer)
            _CodigoTipoDeDocumentoFrete = value
        End Set
    End Property

    Public Property TipoDeDocumentoFrete() As System.Nullable(Of eTipoDeDocumentoFrete)
        Get
            Return _TipoDeDocumentoFrete
        End Get
        Set(ByVal value As System.Nullable(Of eTipoDeDocumentoFrete))
            _TipoDeDocumentoFrete = value
        End Set
    End Property

    Public Property NotasReferenciais() As ListNotaFiscalReferencial
        Get
            'If _IUD = "U" AndAlso (_NotasReferenciais Is Nothing OrElse _NotasReferenciais.Count = 0) Then
            '    If Me.Itens IsNot Nothing AndAlso Me.Itens.Count > 0 Then
            '        _NotasReferenciais = New ListNotaFiscalReferencial(Me)
            '    End If

            'End If
            'If _CodigoSituacao = 4 AndAlso (_NotasReferenciais Is Nothing OrElse _NotasReferenciais.Count = 0) Then _NotasReferenciais = New ListNotaFiscalReferencial(Me, eTipoReferencial.NFC)

            Return _NotasReferenciais
        End Get
        Set(ByVal value As ListNotaFiscalReferencial)
            _NotasReferenciais = value
        End Set
    End Property

    'FAVORECIDO
    Public Property CodigoFavorecido() As String
        Get
            Return _CodigoFavorecido
        End Get
        Set(ByVal value As String)
            _CodigoFavorecido = value
        End Set
    End Property

    Public Property EnderecoFavorecido() As String
        Get
            Return _EnderecoFavorecido
        End Get
        Set(ByVal value As String)
            _EnderecoFavorecido = value
        End Set
    End Property

    'VERIFICAR A EXISTéNCIA DE FATURAS DE FRETES VINCULADAS A NF
    Public ReadOnly Property TemFaturaDeFrete As Boolean
        Get
            Dim ListaDeFaturasDeFretesXItens = New [Lib].Negocio.ListFaturasDeFretesXItens(Me)

            If Not ListaDeFaturasDeFretesXItens Is Nothing AndAlso ListaDeFaturasDeFretesXItens.Count > 0 Then
                Return True
            Else
                Return False
            End If
        End Get
    End Property

    'COMISSéES
    Public Property ComissoesXBaixas() As ListComissoesXBaixas
        Get
            If _ComissoesXBaixas Is Nothing Then _ComissoesXBaixas = New ListComissoesXBaixas(Me)
            Return _ComissoesXBaixas
        End Get
        Set(ByVal value As ListComissoesXBaixas)
            _ComissoesXBaixas = value
        End Set
    End Property

    'PERCURSOS
    Public Property NotaFiscalXPercursos() As ListNotaFiscalXPercurso
        Get
            If _NotaFiscalXPercursos Is Nothing Then _NotaFiscalXPercursos = New ListNotaFiscalXPercurso(Me)
            Return _NotaFiscalXPercursos
        End Get
        Set(ByVal value As ListNotaFiscalXPercurso)
            _NotaFiscalXPercursos = value
        End Set
    End Property

    'NOTA FISCAL ORIGINAL
    Public Property NotaFiscalOriginal As NotaFiscal
        Get
            If _NotaFiscalOriginal Is Nothing And IUD <> "I" Then
                _NotaFiscalOriginal = New NotaFiscal(Me)
            End If
            Return _NotaFiscalOriginal
        End Get
        Set(value As NotaFiscal)
            _NotaFiscalOriginal = value
        End Set
    End Property

    'ARQUIVO
    Public Property Arquivos As ListArquivo
        Get
            If _Arquivos Is Nothing OrElse Not _Arquivos.Count > 0 Then _Arquivos = New Negocio.ListArquivo(Me)
            Return _Arquivos
        End Get
        Set(value As ListArquivo)
            _Arquivos = value
        End Set
    End Property

    'CONFERENCIA DE NOTAS FISCAIS
    Public Property Conferencia As Boolean
        Get
            Return _Conferencia
        End Get
        Set(value As Boolean)
            _Conferencia = value
        End Set
    End Property

    Public Property UsuarioConferencia As String
        Get
            Return _UsuarioConferencia
        End Get
        Set(value As String)
            _UsuarioConferencia = value
        End Set
    End Property

    Public Property UsuarioConferenciaData As DateTime
        Get
            Return _UsuarioConferenciaData
        End Get
        Set(value As DateTime)
            _UsuarioConferenciaData = value
        End Set
    End Property

    'Peso de Chegada
    Public Property PesoDeChegada() As NotaFiscalXDestino
        Get
            If _PesoDeChegada Is Nothing Then _PesoDeChegada = New NotaFiscalXDestino(Me, True)
            Return _PesoDeChegada
        End Get
        Set(ByVal value As NotaFiscalXDestino)
            _PesoDeChegada = value
        End Set
    End Property

    'XML obsoleto
    'Public Property XML As String
    '    Get
    '        Return _XML
    '    End Get
    '    Set(value As String)
    '        _XML = value
    '    End Set
    'End Property

    'Saldos
    Public Property SaldosItensNota As ListSaldoPedido2015
        Get
            Return _SaldosItensNota
        End Get
        Set(value As ListSaldoPedido2015)
            _SaldosItensNota = value
        End Set
    End Property

    'Obriga NF Produtor
    Public Property ObrigaNFProdutor As Boolean
        Get
            Return _ObrigaNFProdutor
        End Get
        Set(value As Boolean)
            _ObrigaNFProdutor = value
        End Set
    End Property

    'Modo Complemento - Usado na NFE para Complemento de Nota Fiscal
    Public Property ModoComplemento As Boolean
        Get
            Return _ModoComplemento
        End Get
        Set(value As Boolean)
            _ModoComplemento = value
        End Set
    End Property

    Public Property TemRecusa As Boolean
        Get
            If _ChaveNFE IsNot Nothing Then _TemRecusa = VerRecusa()
            Return _TemRecusa
        End Get
        Set(value As Boolean)
            _TemRecusa = value
        End Set
    End Property

    Public Property CodigoNaturezaDeRendimento() As Integer
        Get
            Return _CodigoNaturezaDeRendimento
        End Get
        Set(ByVal value As Integer)
            _CodigoNaturezaDeRendimento = value
        End Set
    End Property

    Public Property NCMXML As String
        Get
            Return _NCMXML
        End Get
        Set(value As String)
            _NCMXML = value
        End Set
    End Property

    Public Property FecharTelaClassificacao As Boolean
        Get
            Return _FecharTelaClassificacao
        End Get
        Set(value As Boolean)
            _FecharTelaClassificacao = value
        End Set
    End Property

    Public Property ParteNomeProdNCMXML As String
        Get
            Return _ParteNomeProdNCMXML
        End Get
        Set(value As String)
            _ParteNomeProdNCMXML = value
        End Set
    End Property

    Public Property XMLImportado As Boolean
        Get
            Return _XMLImportado
        End Get
        Set(value As Boolean)
            _XMLImportado = value
        End Set
    End Property

    Public Property InvoiceNavio() As Integer
        Get
            Return _InvoiceNavio
        End Get
        Set(ByVal value As Integer)
            _InvoiceNavio = value
        End Set
    End Property

    Public Property MicSistemas As Boolean
        Get
            Return _MicSistemas
        End Get
        Set(value As Boolean)
            _MicSistemas = value
        End Set
    End Property

    Public Property VincularNotas As Boolean
        Get
            Return _VincularNotas
        End Get
        Set(value As Boolean)
            _VincularNotas = value
        End Set
    End Property

    Public Property CodigoPedidoMIC As String
        Get
            Return _CodigoPedidoMIC
        End Get
        Set(value As String)
            _CodigoPedidoMIC = value
        End Set
    End Property

    Public Property TemConfirmacaoDaOperacao As Boolean
        Get
            If _ChaveNFE IsNot Nothing Then _TemConfirmacaoDaOperacao = VerConfirmacaoDaOperacao()
            Return _TemConfirmacaoDaOperacao
        End Get
        Set(value As Boolean)
            _TemConfirmacaoDaOperacao = value
        End Set
    End Property

    Public Property DiferencaValorNFXProdutoXML As Decimal
        Get
            Return _DiferencaValorNFXProdutoXML
        End Get
        Set(value As Decimal)
            _DiferencaValorNFXProdutoXML = value
        End Set
    End Property

    Public Property TotalNotaValorModificado As Decimal
        Get
            Return _TotalNotaValorModificado
        End Get
        Set(value As Decimal)
            _TotalNotaValorModificado = value
        End Set
    End Property

    Public Property TotalBaseICMSXML As Decimal
        Get
            Return _TotalBaseICMSXML
        End Get
        Set(value As Decimal)
            _TotalBaseICMSXML = value
        End Set
    End Property

    Public Property TotalICMSXML As Decimal
        Get
            Return _TotalICMSXML
        End Get
        Set(value As Decimal)
            _TotalICMSXML = value
        End Set
    End Property

    Public Property CodigoRepresentante As String
        Get
            Return _CodigoRepresentante
        End Get
        Set(value As String)
            _CodigoRepresentante = value
            _Representante = Nothing
        End Set
    End Property

    Public Property EndRepresentante As Integer
        Get
            Return _EndRepresentante
        End Get
        Set(value As Integer)
            _EndRepresentante = value
            _Representante = Nothing
        End Set
    End Property

    Public ReadOnly Property Representante() As Negocio.Cliente
        Get
            If _Representante Is Nothing And _CodigoRepresentante.Length > 0 Then _Representante = New Negocio.Cliente(_CodigoRepresentante, _EndRepresentante)
            Return _Representante
        End Get
    End Property

    'NOTA FISCAL DE TERCEIROS
    Public Property NFT As Boolean
        Get
            Return _NFT
        End Get
        Set(value As Boolean)
            _NFT = value
        End Set
    End Property

    Public Property NotaDeTerceiro As Boolean
        Get
            Return _NotaDeTerceiro
        End Get
        Set(value As Boolean)
            _NotaDeTerceiro = value
        End Set
    End Property

    Public Property DataDevolucao() As DateTime
        Get
            Return _DataDevolucao
        End Get
        Set(ByVal value As DateTime)
            _DataDevolucao = value
        End Set
    End Property

    Public Property EmpresaEmissor As Boolean
        Get
            Return _EmpresaEmissor
        End Get
        Set(value As Boolean)
            _EmpresaEmissor = value
        End Set
    End Property

    Public Property CodigoTomador As String
        Get
            Return _CodigoTomador
        End Get
        Set(value As String)
            _CodigoTomador = value
            _Tomador = Nothing
        End Set
    End Property

    Public Property EnderecoTomador As Integer
        Get
            Return _EnderecoTomador
        End Get
        Set(value As Integer)
            _EnderecoTomador = value
            _Tomador = Nothing
        End Set
    End Property

    Public ReadOnly Property Tomador() As Negocio.Cliente
        Get
            If _Tomador Is Nothing And _CodigoTomador.Length > 0 Then _Tomador = New Negocio.Cliente(_CodigoTomador, _EnderecoTomador)
            Return _Tomador
        End Get
    End Property

    Public Property NotaOrigemImportacaoXML As List(Of NotaFiscal)
        Get
            Return _NotaOrigemImportacaoXML
        End Get
        Set(value As List(Of NotaFiscal))
            _NotaOrigemImportacaoXML = value
        End Set
    End Property

    Public Property ContratoDeFreteCTe As Boolean
        Get
            Return _ContratoDeFreteCTe
        End Get
        Set(value As Boolean)
            _ContratoDeFreteCTe = value
        End Set
    End Property

    Public Property EmitindoCTe As Boolean
        Get
            Return _EmitindoCTe
        End Get
        Set(value As Boolean)
            _EmitindoCTe = value
        End Set
    End Property

    Public Property TemPedagioCTe As Boolean
        Get
            Return _TemPedagioCTe
        End Get
        Set(value As Boolean)
            _TemPedagioCTe = value
        End Set
    End Property


#End Region

#Region "Methods"
    Public Sub VerificarIndice()
        'Dim fix As Negocio.Fixacao = Nothing

        'If _CodigoPedido = 0 Then
        '    _IndiceNota = New Cotacao(2, _DataNota).Indice 'PTAX
        'End If

        'If _CodigoTipoDeDocumento = 1 _
        'AndAlso _CodigoPedido > 0 _
        'AndAlso Pedido IsNot Nothing _
        'AndAlso Pedido.SubOperacao IsNot Nothing _
        'AndAlso Pedido.SubOperacao.Classe = Negocio.eClassesOperacoes.AFIXAR _
        'AndAlso Not Itens Is Nothing _
        'AndAlso Itens.Count > 0 _
        'AndAlso Itens(0).CodigoFixacao > 0 Then
        '    fix = Pedido.Itens(0).Fixacoes.Where(Function(s) s.Codigo = Itens(0).CodigoFixacao).FirstOrDefault
        'End If

        'If Not fix Is Nothing AndAlso fix.IndiceFixado > 0 Then
        '    _IndiceNota = fix.IndiceFixado
        'ElseIf _CodigoPedido > 0 AndAlso Pedido IsNot Nothing AndAlso Pedido.IndexadorFixo Then
        '    _IndiceNota = Pedido.IndiceFixado
        'Else
        '    If _CodigoTipoDeDocumento = 1 AndAlso _CodigoPedido > 0 Then
        '        _IndiceNota = New Cotacao(Pedido.CodigoIndexador, _DataNota).Indice
        '    End If
        'End If
    End Sub

    Public Sub AtualizaTotais()
        Dim Prod, ProdMoeda, Liq, LiqMoeda, VIPI, VIcms, BIcmsST, VIcmsST, Vol, Num, Fr As Decimal
        Dim BIcms As Decimal = 0
        Dim PesoB, PesoL, desc, aduan, seg As Decimal

        For Each item As Negocio.NotaFiscalXItem In Me.Itens.Where(Function(s) s.IUD <> "D")

            PesoB += item.PesoBruto
            PesoL += item.PesoLiquido

            Vol += item.Volumes

            Num += item.Numeracao
            Prod += item.ValorTotal

            'ProdMoeda += item.ValorTotalMoeda
            'LiqMoeda += item.ValorLiquidoMoeda

            For Each enc As NotaFiscalXItemXEncargo In item.Encargos.Where(Function(s) s.IUD <> "D").OrderBy(Function(s) s.Codigo)

                Select Case enc.Codigo
                    Case "IPI"
                        VIPI += enc.Valor

                    Case "ICMS"

                        If Me.MicSistemas AndAlso Me.TotalBaseICMSXML > 0 AndAlso Me.TotalICMSXML > 0 Then
                            BIcms = Me.TotalBaseICMSXML
                            VIcms = Me.TotalICMSXML
                            enc.Base = Me.TotalBaseICMSXML
                            enc.Valor = Me.TotalICMSXML
                        Else
                            BIcms += IIf(enc.Valor > 0, enc.Base, 0)
                            VIcms += enc.Valor
                        End If

                    Case "ICMS A REC."
                        BIcms += IIf(enc.Valor > 0, enc.Base, 0)
                        VIcms += enc.Valor
                    Case "ICMS-ST"
                        BIcmsST += IIf(enc.Valor > 0, enc.Base, 0)
                        VIcmsST += enc.Valor
                    Case "LIQUIDO"
                        Liq += enc.Valor
                        item.ValorLiquido = enc.Valor
                    Case "FRETES"
                        Fr += enc.Valor
                    Case "DESCONTOS"
                        desc += enc.Valor
                    Case "ADUANEIRAS"
                        aduan += enc.Valor
                    Case "DESP.ADUANEIRAS"
                        aduan += enc.Valor
                    Case "SEGURO"
                        seg += enc.Valor
                End Select

                If enc.Codigo = "PRODUTO" AndAlso BIcms = 0 AndAlso (enc.SituacaoTributaria = 51 Or enc.SituacaoTributaria = 151 Or enc.SituacaoTributaria = 651) Then
                    BIcms += enc.Valor
                End If

                If enc.Codigo = "IPI A RECUP." AndAlso enc.Sinal = "+" Then
                    VIPI += enc.Valor
                End If

            Next

            ProdMoeda += item.ValorTotalMoeda
            LiqMoeda += item.ValorLiquidoMoeda

        Next

        Me.TotalProduto = Prod
        Me.TotalProdutoMoeda = ProdMoeda

        Me.TotalNota = IIf(Liq < 0, 0, Liq)
        Me.TotalNotaMoeda = IIf(LiqMoeda < 0, 0, LiqMoeda)

        Me.ValorIPI = VIPI
        Me.ValorBaseIcms = BIcms
        Me.ValorIcms = VIcms
        Me.ValorBaseIcmsST = BIcmsST
        Me.ValorIcmsST = VIcmsST

        If Vol > 0 Then
            Me.Quantidade = Vol
        End If

        Me.Numero = Num
        Me.ValorFrete = Fr
        Me.ValorAduaneira = aduan
        Me.ValorDesconto = desc
        Me.ValorSeguro = seg

        If Me.MicSistemas = False Then
            Me.PesoBruto = PesoB
            Me.PesoLiquido = PesoL
        End If

    End Sub

    'Private Sub AtualizaItens(ByVal opcao As Integer)
    '    VerificarIndice()
    '    For Each item As Negocio.NotaFiscalXItem In Itens
    '        Select Case opcao
    '            Case 1
    '                If Not SubOperacao IsNot Nothing AndAlso SubOperacao.Devolucao Then
    '                    Dim ValorUnitario As Decimal

    '                    If Pedido IsNot Nothing AndAlso Pedido.Moeda IsNot Nothing AndAlso Pedido.Moeda.Classificacao = Negocio.eTiposMoeda.MoedaEstrangeira Then
    '                        ValorUnitario = Funcoes.ConverteMoeda(item.UnitarioMoeda, IndiceNota, eTiposMoeda.Oficial)
    '                    Else
    '                        ValorUnitario = item.Unitario
    '                    End If

    '                    item.Unitario = ValorUnitario.ToString("N10")
    '                    If item.Unitario > 0 And item.QuantidadeFiscal > 0 Then item.ValorTotal = (item.QuantidadeFiscal / item.Produto.BaseCalculo) * ValorUnitario

    '                    item.CarregandoEncargos = True
    '                    item.Encargos = New ListNotaFiscalXItemXEncargo(item)
    '                    item.CarregandoEncargos = False
    '                End If
    '            Case 2
    '                item.CodigoSubOperacao = Me.CodigoSubOperacao
    '            Case 3
    '                item.CarregandoEncargos = True
    '                item.Encargos = New ListNotaFiscalXItemXEncargo(item)
    '                item.CarregandoEncargos = False
    '        End Select
    '    Next
    'End Sub

    Private Function AtualizaNumeradorNota() As Numerador
        Dim n As Numerador = Nothing
        If _CodigoTipoDeDocumento = 12 Then
            n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 38)
            Me.Codigo = n.Sequencia + 1
            Me.Serie = n.Serie
        ElseIf _CodigoTipoDeDocumento = 4 Then
            n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 40)
            Me.Codigo = n.Sequencia + 1
            Me.Serie = n.Serie
        ElseIf _CodigoTipoDeDocumento = eTipoDeDocumento.NFC_e Then
            If _NossaEmissao Then
                n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 33)
                Me.Codigo = n.Sequencia + 1
                Me.Serie = n.Serie
            End If
        ElseIf _CodigoTipoDeDocumento = 2 OrElse _CodigoTipoDeDocumento = 8 OrElse _CodigoTipoDeDocumento = 10 OrElse _CodigoTipoDeDocumento = 14 Then
            If _NossaEmissao Then
                n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 35)
                Me.Codigo = n.Sequencia + 1
                Me.Serie = n.Serie
            ElseIf NFG AndAlso SubOperacao IsNot Nothing AndAlso SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                _EntradaSaida = eEntradaSaida.Entrada
            ElseIf NFG AndAlso SubOperacao IsNot Nothing AndAlso SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                _EntradaSaida = eEntradaSaida.Saida
            End If
        ElseIf SubOperacao IsNot Nothing AndAlso SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
            If _NossaEmissao Then
                If Retroativa Then
                    n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 21)
                Else
                    n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 20)
                End If
                Me.Codigo = n.Sequencia + 1
                Me.Serie = n.Serie
            Else
                Me.Codigo = NotaProdutor
                Me.Serie = SerieNotaProdutor
            End If
        Else
            If _NossaEmissao Then
                If Retroativa Then
                    n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 31)
                Else
                    n = New Numerador(Empresa.Codigo, Empresa.CodigoEndereco, 30)
                End If

                Me.Codigo = n.Sequencia + 1
                Me.Serie = n.Serie
                _NossaEmissao = True
                _Eletronica = Empresa.Empresa.NotaFiscalEletronica
            ElseIf NotaProdutor > 0 Then
                _Codigo = NotaProdutor
                _Serie = SerieNotaProdutor
            End If
        End If

        'DESCOMENTEN A LINHA 2107 e 2108 PARA FAZER UMA NOTA COM DATA RETROATIVA, não ESQUEéA DE COMENTAR APéS A EMISSão.
        '_Nota = 3
        '_Serie = "3"

        Return n
    End Function

    Public Function AtualizarDeposito() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim strSQL As String

        Razao.ExcluiContabilizacaoNotaSQL(Sqls)

        strSQL = " Update NotasFiscais set" & vbCrLf &
                 "     Deposito            ='" & Me.CodigoDeposito & "'" & vbCrLf &
                 "	  ,EndDeposito         = " & Me.EnderecoDeposito & vbCrLf &
                 "	  ,Destino             ='" & Me.CodigoDestino & "'" & vbCrLf &
                 "	  ,EndDestino          = " & Me.EnderecoDestino & vbCrLf &
                 "	  ,UsuarioAlteracao    ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                 "	  ,UsuarioAlteracaoData ='" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                 "	Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                 "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                 "	  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                 "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                 "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                 "    and Nota_Id         = " & Me.Codigo & vbCrLf &
                 "    and Serie_Id        ='" & Me.Serie & "';"
        Sqls.Add(strSQL)

        For Each item As Negocio.NotaFiscalXItem In Me.Itens
            strSQL = " Update NotasFiscaisXItens set" & vbCrLf &
                     "     Deposito    ='" & Me.CodigoDeposito & "'" & vbCrLf &
                     "	  ,EndDeposito = " & Me.EnderecoDeposito & vbCrLf &
                     "	Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                     "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                     "	  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                     "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                     "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                     "    and Nota_Id         = " & Me.Codigo & vbCrLf &
                     "    and Serie_Id        ='" & Me.Serie & "';"
            Sqls.Add(strSQL)
        Next

        If CodigoRomaneio > 0 Then
            strSQL = " Update Romaneios set" & vbCrLf &
                     "    Deposito    ='" & Me.CodigoDeposito & "'" & vbCrLf &
                     "	 ,EndDeposito = " & Me.EnderecoDeposito & "" & vbCrLf &
                     "	Where Empresa_Id    ='" & Me.Romaneio.CodigoEmpresa & "'" & vbCrLf &
                     "    and EndEmpresa_Id = " & Me.Romaneio.EnderecoEmpresa & vbCrLf &
                     "    and Romaneio_Id   = " & Me.Romaneio.Codigo & ";"

            Sqls.Add(strSQL)

            strSQL = " Update Pesagem Set" & vbCrLf &
                     "     Deposito             ='" & Me.CodigoDeposito & "'" & vbCrLf &
                     "	  ,EndDeposito          = " & Me.EnderecoDeposito & vbCrLf &
                     "    ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                     "    ,UsuarioAlteracaoData ='" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                     "	Where Empresa_Id    ='" & _Romaneio.CodigoEmpresa & "'" & vbCrLf &
                     "    and EndEmpresa_Id = " & _Romaneio.EnderecoEmpresa & vbCrLf &
                     "    and Pesagem_id    = (Select Pesagem_id" & vbCrLf &
                     "                           From RomaneiosXPesagens" & vbCrLf &
                     "                          where Empresa_Id    = '" & _Romaneio.CodigoEmpresa & "'" & vbCrLf &
                     "                            and EndEmpresa_Id = " & _Romaneio.EnderecoEmpresa & vbCrLf &
                     "                            and Romaneio_id   = " & _Romaneio.Codigo & ");"
            Sqls.Add(strSQL)
        End If

        Razao.ContabilizarNotaSql(Sqls)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AtualizarCIFFOB() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim strSQL As String

        strSQL = " Update NotasFiscais set " & vbCrLf &
                 "      CIFFOB               = '" & Me.CIFFOB.ToString & "'," & vbCrLf &
                 "      UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                 "      UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                 "  Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                 "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                 "	  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                 "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                 "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                 "    and Nota_Id         = " & Me.Codigo & vbCrLf &
                 "    and Serie_Id        ='" & Me.Serie & "';" & vbCrLf
        Sqls.Add(strSQL)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AtualizarObservacao() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim strSQL As String

        strSQL = " Update NotasFiscais set " & vbCrLf &
                 "      Observacoes                = '" & Me.Observacoes & "'," & vbCrLf &
                 "      ObservacoesDeEmbarque      = '" & Me.ObservacoesDeEmbarque & "'," & vbCrLf &
                 "      ObservacoesControleInterno = '" & Me.ObservacoesControleInterno & "'," & vbCrLf &
                 "      UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                 "      UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                 "  Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                 "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                 "	  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                 "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                 "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                 "    and Nota_Id         = " & Me.Codigo & vbCrLf &
                 "    and Serie_Id        ='" & Me.Serie & "';" & vbCrLf
        Sqls.Add(strSQL)

        If Me.NFG Then
            Dim cLote As Integer = 0
            If Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
                Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) OrElse
                Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) OrElse
                Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) OrElse
                Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) OrElse
                Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) OrElse
                Me.Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
                cLote = 21
            Else
                cLote = 9
            End If

            Me.Razao.ExcluiContabilizacaoNotaSQL(Sqls, cLote)
            Me.Razao.ContabilizarNotaSql(Sqls, cLote)
        Else
            Me.Razao.ExcluiContabilizacaoNotaSQL(Sqls)
            Me.Razao.ContabilizarNotaSql(Sqls)
        End If

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function AtualizaHoraNota() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim sql As String = String.Empty

        sql = "Update NotasFiscais" & vbCrLf &
        "   set UsuarioInclusaoData     = '" & Me.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf &
        "      UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
        "      UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'," & vbCrLf &
        "      ObservacoesControleInterno  = 'Ajustado hora da nota fiscal porque servidor estava com a data errada.'" & vbCrLf &
        " Where Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
        "   And EndEmpresa_Id   =  " & Me.EnderecoEmpresa & vbCrLf &
        "   And Cliente_Id      = '" & Me.CodigoCliente & "'" & vbCrLf &
        "   And EndCliente_Id   =  " & Me.EnderecoCliente & vbCrLf &
        "   And EntradaSaida_Id = '" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
        "   And Serie_Id        =  " & Me.Serie & vbCrLf &
        "   And Nota_Id         =  " & Me.Codigo & vbCrLf

        Sqls.Add(sql)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else : Return False
        End If
    End Function

    Public Function AtualizaDatasNota() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim sql As String = String.Empty

        Me.Razao.ExcluiContabilizacaoNotaSQL(Sqls)

        sql = "Update NotasFiscais" & vbCrLf &
                "   set DataDaNota          = '" & Me.DataNota.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                "       Movimento           = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'," & vbCrLf &
                "      UsuarioAlteracao     = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                "      UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                " Where Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                "   And EndEmpresa_Id   =  " & Me.EnderecoEmpresa & vbCrLf &
                "   And Cliente_Id      = '" & Me.CodigoCliente & "'" & vbCrLf &
                "   And EndCliente_Id   =  " & Me.EnderecoCliente & vbCrLf &
                "   And EntradaSaida_Id = '" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                "   And Serie_Id        = '" & Me.Serie & "'" & vbCrLf &
                "   And Nota_Id         =  " & Me.Codigo & vbCrLf

        Sqls.Add(sql)

        Me.Razao.ContabilizarNotaSql(Sqls)

        If Not Me.Romaneio Is Nothing AndAlso Me.Romaneio.Codigo > 0 Then
            sql = "Update Romaneios" & vbCrLf &
                    "   set Movimento           = '" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                    " Where Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                    "   And EndEmpresa_Id   =  " & Me.EnderecoEmpresa & vbCrLf &
                    "   And Romaneio_Id   =  " & Me.Romaneio.Codigo

            Sqls.Add(sql)
        End If

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else : Return False
        End If
    End Function

    Public Function AtualizaFinalidadeNota() As Boolean
        Dim Banco As New AcessaBanco
        Dim sql As String = String.Empty

        sql = "Update NotasFiscais" & vbCrLf &
                "   set Finalidade                 =  " & Me.CodigoFinalidade & "," & vbCrLf &
                "       ObservacoesControleInterno = '" & Me.ObservacoesControleInterno & "'," & vbCrLf &
                "       UsuarioAlteracao           = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                "       UsuarioAlteracaoData       = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                " Where Empresa_Id      = '" & Me.CodigoEmpresa & "'" & vbCrLf &
                "   And EndEmpresa_Id   =  " & Me.EnderecoEmpresa & vbCrLf &
                "   And Cliente_Id      = '" & Me.CodigoCliente & "'" & vbCrLf &
                "   And EndCliente_Id   =  " & Me.EnderecoCliente & vbCrLf &
                "   And EntradaSaida_Id = '" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                "   And Serie_Id        =  " & Me.Serie & vbCrLf &
                "   And Nota_Id         =  " & Me.Codigo & vbCrLf

        If Banco.GravaBanco(sql) Then
            Return True
        Else : Return False
        End If
    End Function

    Public Function VerRecusa() As Boolean
        Dim Banco As New AcessaBanco
        Dim strSQL As String

        strSQL = "Select ChaveNFe " & vbCrLf &
                 "from ManifestoNFE" & vbCrLf &
                 "Where CodigoEvento = '210240'" &
                 "  and ChaveNFe = '" & Me.ChaveNFE & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Recusa")

        If ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function VerConfirmacaoDaOperacao() As Boolean
        Dim Banco As New AcessaBanco
        Dim strSQL As String

        strSQL = "Select ChaveNFe " & vbCrLf &
                 "from ManifestoNFE" & vbCrLf &
                 "Where CodigoEvento = '210200'" &
                 "  and ChaveNFe = '" & Me.ChaveNFE & "'"

        Dim ds As DataSet = Banco.ConsultaDataSet(strSQL, "Recusa")

        If ds.Tables(0).Rows.Count = 0 Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        'If FinanceiroNovo AndAlso Not String.IsNullOrWhiteSpace(Empresa.Empresa.Servidor) AndAlso UsuarioServidor.NomeServidor <> Empresa.Empresa.Servidor Then
        '    Banco = New AcessaBanco(2, Empresa.Empresa.Servidor)
        'End If
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)

        'If FinanceiroNovo Then getSqlException(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    'Nao deveria existir essa function
    Public Function Salvar(ByRef Sqls As ArrayList) As Boolean
        Sqls.Clear()
        SalvarSql(Sqls)
        Return True
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String = ""

        'If FinanceiroNovo AndAlso (Me.IUD = "C" Or Me.IUD = "D") AndAlso Not Me.NFG Then
        '    Me.Titulos.ReajFinanceiro = New ReajusteFinanceiro(Me)
        '    If Me.Pedido.Troca AndAlso Me.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Baixa).Count > 0 Then
        '        Me.Titulos.ReajFinanceiro.ReajustaTroca()
        '    ElseIf Me.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
        '        Me.Titulos.ReajFinanceiro.ReajustaAFixar()
        '    Else
        '        Me.Titulos.ReajFinanceiro.Reajusta()
        '    End If
        'End If

        Select Case Me.IUD
            Case "I"
                '*******************************************************
                '**** Procedimento para obter o numerador da nota  *****
                '*******************************************************

                Dim NumNota As New Numerador

                If XMLImportado = False Then

                    NumNota = AtualizaNumeradorNota()

                    If Me.NossaEmissao Then
                        'COMENTE A LINHA 2219 PARA FAZER UMA NOTA COM DATA RETROATIVA, não ESQUEéA DE DESCOMENTAR APéS A EMISSão.
                        'DEPOIS Vé NO MéTODO AtualizaNumeradorNota() LINHA 2107 e 2108 para acertar a numeração.
                        Sqls.Add(NumNota.IncrementarNumeradorSql())
                    End If

                End If

                '*******************************************************
                '**** Sql de Insert Da Nota Fiscal   *******************
                '*******************************************************
                strSQL &= " Insert Into NotasFiscais " & vbCrLf &
                          " (Empresa_Id, EndEmpresa_Id, " & vbCrLf &
                          "  Cliente_Id, EndCliente_Id, " & vbCrLf &
                          "  EntradaSaida_Id, Serie_Id, Nota_Id, TipoDeDocumento, " & vbCrLf &
                          "  Formulario, Pedido, Procuracao, " & vbCrLf &
                          "  Operacao, SubOperacao, Finalidade, " & vbCrLf &
                          "  Movimento, DataDaNota, NossaEmissao, " & vbCrLf &
                          "  SerieNotadoProdutor, NumeroNotadoProdutor, " & vbCrLf &
                          "  Deposito, EndDeposito, " & vbCrLf &
                          "  Destino, EndDestino, " & vbCrLf &
                          "  Transbordo, EndTransbordo, " & vbCrLf &
                          "  Agenciador, EndAgenciador, " & vbCrLf &
                          "  Entrega, EndEntrega, " & vbCrLf &
                          "  CIFFOB, Observacoes, Eletronica, " & vbCrLf &
                          "  Autorizacao," & vbCrLf &
                          "  UsuarioInclusao, UsuarioInclusaoData, " & vbCrLf

                If UsuarioAlteracao.Length > 0 Then strSQL &= "UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf

                strSQL &= "  Situacao,ObservacoesDoProduto,ObservacoesDeEmbarque,ObservacoesControleInterno,EstadoDoCliente,LocalEmbarque,EndLocalEmbarque, " & vbCrLf &
                          "  ContratoANTT, ProtocoloANTT, CartaoPgtoFrete, DataTermino, idPamcard, TipoDeDocumentoFrete, Favorecido, EndFavorecido, ProprietarioDaMercadoria, EndProprietarioDaMercadoria, " & vbCrLf &
                          "  NFG, Conferencia, UsuarioConferencia, UsuarioConferenciaData, Troca, CodigoNaturezaDeRendimento, InvoiceNavio, Representante,EndRepresentante, NFT, Tomador, EndTomador, ContratoDeFrete) " & vbCrLf &
                          "Values ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ", " & vbCrLf &
                                                     "'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ", " & vbCrLf &
                                                     "'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & "," & Me.CodigoTipoDeDocumento & ", " & vbCrLf &
                                                     " " & Me.Formulario & ", " & Me.CodigoPedido.ToSqlNULL & ", " & Me.CodigoProcuracao.ToSqlNULL & ", " & vbCrLf &
                                                     " " & Me.CodigoOperacao & ", " & Me.CodigoSubOperacao & ", " & Me.CodigoFinalidade & ", " & vbCrLf &
                                                     "'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.DataNota.ToString("yyyy-MM-dd") & "','" & IIf(Me.NossaEmissao, "S", "N") & "' , " & vbCrLf &
                                                     "'" & Me.SerieNotaProdutor & "', " & IIf(Me.NotaProdutor = 0, "Null", Me.NotaProdutor) & ", " & vbCrLf &
                                                     "'" & Me.CodigoDeposito & "', " & Me.EnderecoDeposito & ", " & vbCrLf &
                                                     "'" & Me.CodigoDestino & "', " & Me.EnderecoDestino & ", " & vbCrLf &
                                                     "'" & Me.CodigoTransbordo & "', " & Me.EnderecoTransbordo & ", " & vbCrLf &
                                                     "'" & Me.CodigoAgenciador & "', " & Me.EnderecoAgenciador & ", " & vbCrLf &
                                                     "'" & Me.CodigoEntrega & "', " & Me.EnderecoEntrega & ", " & vbCrLf &
                                                     "'" & Me.CIFFOB.ToString & "', '" & Me.Observacoes & "', '" & IIf(Me.Eletronica, "S", "N") & "', " & vbCrLf &
                                                     " " & Me.CodigoAutorizacao & "," & vbCrLf &
                                                     "'" & Me.UsuarioInclusao & "','" & Me.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf

                If _UsuarioAlteracao.Length > 0 Then strSQL &= "'" & Me.UsuarioAlteracao & "','" & Me.DataAlteracao.ToString("yyyy-MM-dd HH:mm:ss") & "'," & vbCrLf

                strSQL &= "'" & Me.CodigoSituacao & "','" & Me.ObservacoesDoProduto & "','" & Me.ObservacoesDeEmbarque & "','" & Me.ObservacoesControleInterno & "'" & vbCrLf &
                          ",'" & Me.Cliente.CodigoEstado & "'" & vbCrLf &
                          "," & IIf(Me.CodigoLocalEmbarque.Length > 0, "'" & Me.CodigoLocalEmbarque & "'", "NULL") & vbCrLf &
                          "," & IIf(Me.CodigoLocalEmbarque.Length > 0, Me.EndLocalEmbarque, "NULL") & vbCrLf &
                          ",'" & Me.ContratoANTT & "','" & Me.ProtocoloANTT & "','" & Me.CartaoPgtoFrete & "','" & Me.DataTermino.ToString("yyyy-MM-dd") & "','" & _idPamcard & "', " & vbCrLf

                If Me.TipoDeDocumentoFrete IsNot Nothing Then strSQL &= "'" & Me.TipoDeDocumentoFrete & "'," Else strSQL &= "null," & vbCrLf

                If Not String.IsNullOrWhiteSpace(Me.CodigoFavorecido) Then
                    strSQL &= "'" & Me.CodigoFavorecido & "'," & Me.EnderecoFavorecido & ","
                Else
                    strSQL &= "null,null," & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(Me.CodigoProprietarioDaMercadoria) Then strSQL &= "'" & Me.CodigoProprietarioDaMercadoria & "'," & Me.EnderecoProprietarioDaMercadoria & "," Else strSQL &= "null,null," & vbCrLf

                strSQL &= IIf(_NFG, 1, 0) & "," & vbCrLf &
                          IIf(Me.Conferencia, 1, IIf(Me.Empresa.Empresa.ConferenciaNFE, 0, "NULL")) & "," & vbCrLf &
                          IIf(Me.Conferencia AndAlso Not String.IsNullOrWhiteSpace(UsuarioConferencia), UsuarioConferencia & "," & Me.UsuarioConferenciaData.ToString("yyyy-MM-dd HH:mm:ss"), "NULL,NULL") & vbCrLf &
                          "," & IIf(Me.Troca, 1, 0) & "," & Me.CodigoNaturezaDeRendimento & "," & Me.InvoiceNavio & ""

                If Not String.IsNullOrWhiteSpace(Me.CodigoRepresentante) Then
                    strSQL &= ", '" & Me.CodigoRepresentante & "'," & Me.EndRepresentante & vbCrLf
                Else
                    strSQL &= ",NULL,NULL" & vbCrLf
                End If

                strSQL &= ", " & IIf(_NFT, 1, 0) & vbCrLf

                If Not String.IsNullOrWhiteSpace(Me.CodigoTomador) Then
                    strSQL &= ", '" & Me.CodigoTomador & "'," & Me.EnderecoTomador
                Else
                    strSQL &= ",NULL,NULL"
                End If

                strSQL &= ", " & IIf(_ContratoDeFreteCTe, 1, 0) & vbCrLf

                strSQL &= ")" & vbCrLf

                Sqls.Add(strSQL)
                '***********************************************************************
                '**** Procedimento para Salvar as tabelas relacionadas com a Nota  *****
                '*******  Produtos, Encargos, Vencimentos, Laudo, NFE Pendente  ********
                '***********************************************************************
                SalvarTabelasRelacionadasSql(Sqls)

                If Me.NotaDeTerceiro = True AndAlso Me.Arquivos.Count > 0 Then
                    For Each arq In Me.Arquivos
                        If arq.IUD = "I" Then
                            arq.CodigoEmpresa = Me.CodigoEmpresa
                            arq.EnderecoEmpresa = Me.EnderecoEmpresa
                            arq.CodigoCliente = Me.CodigoCliente
                            arq.EnderecoCliente = Me.EnderecoCliente
                            arq.CodigoNota = Me.Codigo
                            arq.Serie = Me.Serie
                            arq.CodigoPedido = Me.CodigoPedido
                            arq.SalvarSql(Sqls)
                            arq.IUD = "U"
                        End If
                    Next
                End If

            Case "U"
                strSQL = " Update NotasFiscais set " & vbCrLf &
                         "      Procuracao             = " & Me.CodigoProcuracao.ToSqlNULL & vbCrLf &
                         "     ,Operacao               = " & Me.CodigoOperacao & vbCrLf &
                         "     ,SubOperacao            = " & Me.CodigoSubOperacao & vbCrLf &
                         "     ,Movimento              ='" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         "     ,DataDaNota             ='" & Me.DataNota.ToString("yyyy-MM-dd") & "'" & vbCrLf &
                         "     ,TipoDeDocumento        = " & Me.CodigoTipoDeDocumento & vbCrLf &
                         "     ,Deposito               ='" & Me.CodigoDeposito & "'" & vbCrLf &
                         "     ,EndDeposito            = " & Me.EnderecoDeposito & vbCrLf &
                         "     ,Destino                ='" & Me.CodigoDestino & "'" & vbCrLf &
                         "     ,EndDestino             = " & Me.EnderecoDestino & vbCrLf &
                         "     ,LocalEmbarque          ='" & Me.CodigoLocalEmbarque & "'" & vbCrLf &
                         "     ,EndLocalEmbarque       = " & Me.EndLocalEmbarque & vbCrLf &
                         "     ,Transbordo             ='" & Me.CodigoTransbordo & "'" & vbCrLf &
                         "     ,EndTransbordo          = " & Me.EnderecoTransbordo & vbCrLf &
                         "     ,Entrega                ='" & Me.CodigoEntrega & "'" & vbCrLf &
                         "     ,EndEntrega             = " & Me.EnderecoEntrega & vbCrLf &
                         "     ,Observacoes            ='" & Me.Observacoes & "'" & vbCrLf &
                         "     ,Eletronica             ='" & IIf(Me.Eletronica, "S", "N") & "'" & vbCrLf &
                         "     ,Situacao               = " & Me.CodigoSituacao & vbCrLf &
                         "	   ,UsuarioAlteracao       ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                         "	   ,UsuarioAlteracaoData   = getdate() " & vbCrLf &
                         "     ,ContratoANTT           ='" & Me.ContratoANTT & "'" & vbCrLf &
                         "     ,ProtocoloANTT          ='" & Me.ProtocoloANTT & "'" & vbCrLf &
                         "     ,idPamcard              ='" & Me.idPamcard & "'" & vbCrLf &
                         "	   ,Conferencia            = " & IIf(Me.NFG, 1, 0) & vbCrLf &
                         "	   ,UsuarioConferencia     = " & IIf(Me.NFG, "'" & Me.UsuarioConferencia & "'", 0) & vbCrLf &
                         "	   ,UsuarioConferenciaData = " & IIf(Me.Conferencia, "'" & Me.UsuarioConferenciaData.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & vbCrLf &
                         "     ,Troca                  = " & IIf(Me.Troca, 1, 0) & vbCrLf &
                         "     ,CodigoNaturezaDeRendimento = " & Me.CodigoNaturezaDeRendimento & vbCrLf &
                         "     ,InvoiceNavio           = " & Me.InvoiceNavio & vbCrLf

                If Not String.IsNullOrWhiteSpace(Me.CodigoFavorecido) Then
                    strSQL &= "     ,Favorecido    ='" & Me.CodigoFavorecido & "'" & vbCrLf &
                              "     ,EndFavorecido = " & Me.EnderecoFavorecido & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(Me.CodigoRepresentante) Then
                    strSQL &= "     ,Representante    = '" & Me.CodigoRepresentante & "'" & vbCrLf &
                              "     ,EndRepresentante = " & Me.EndRepresentante & vbCrLf
                Else
                    strSQL &= "     ,Representante    = NULL" & vbCrLf &
                              "     ,EndRepresentante = NULL" & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(Me.CodigoProprietarioDaMercadoria) Then
                    strSQL &= "     ,ProprietarioDaMercadoria    ='" & Me.CodigoProprietarioDaMercadoria & "'" & vbCrLf &
                              "     ,EndProprietarioDaMercadoria = " & Me.EnderecoProprietarioDaMercadoria & vbCrLf
                End If

                If Not String.IsNullOrWhiteSpace(Me.CodigoTomador) Then
                    strSQL &= "     ,Tomador    = '" & Me.CodigoTomador & "'" & vbCrLf &
                              "     ,EndTomador = " & Me.EnderecoTomador & vbCrLf
                Else
                    strSQL &= "     ,Tomador    = NULL" & vbCrLf &
                              "     ,EndTomador = NULL" & vbCrLf
                End If

                strSQL &= "  Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                          "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                          "	  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                          "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                          "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                          "    and Nota_Id         = " & Me.Codigo & vbCrLf &
                          "    and Serie_Id        ='" & Me.Serie & "';" & vbCrLf
                Sqls.Add(strSQL)

                For Each itemNF As Negocio.NotaFiscalXItem In _Itens.Where(Function(I) I.IUD <> "D")
                    'strSQL = " Update NotasFiscaisxitens set " & vbCrLf & _
                    '         "     Deposito         ='" & Me.CodigoDeposito & "'" & vbCrLf & _
                    '         "    ,EndDeposito      = " & Me.EnderecoDeposito & vbCrLf & _
                    '         "    ,NumeroPecas      = " & itemNF.NumeroPecas & vbCrLf & _
                    '         "  Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                    '         "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf & _
                    '         "	  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf & _
                    '         "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf & _
                    '         "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf & _
                    '         "    and Nota_Id         = " & Me.Codigo & vbCrLf & _
                    '         "    and Serie_Id        ='" & Me.Serie & "';" & vbCrLf
                    'Sqls.Add(strSQL)

                    'If itemNF.Encargos IsNot Nothing AndAlso itemNF.Encargos.Count > 0 Then
                    '    itemNF.Encargos.SalvarSql(Sqls)
                    'End If

                    If Not itemNF.IUD = "I" Then itemNF.IUD = Me.IUD
                Next

                If CodigoRomaneio > 0 Then
                    strSQL = " Update romaneios set " & vbCrLf &
                             "     Deposito         ='" & Me.CodigoDeposito & "'" & vbCrLf &
                             "    ,EndDeposito      = " & Me.EnderecoDeposito & vbCrLf &
                             "    ,Destino          ='" & Me.CodigoDestino & "'" & vbCrLf &
                             "    ,EndDestino       = " & Me.EnderecoDestino & vbCrLf &
                             "    ,Transbordo       ='" & Me.CodigoTransbordo & "'" & vbCrLf &
                             "    ,EndTransbordo    = " & Me.EnderecoTransbordo & vbCrLf &
                             "  FROM NotasFiscaisXRomaneios nfxr" & vbCrLf &
                             " INNER JOIN Romaneios " & vbCrLf &
                             "    ON nfxr.Empresa_Id    = Romaneios.Empresa_Id " & vbCrLf &
                             "   AND nfxr.EndEmpresa_Id = Romaneios.EndEmpresa_Id " & vbCrLf &
                             "   AND nfxr.Romaneio_Id   = Romaneios.Romaneio_Id " & vbCrLf &
                             " Where nfxr.Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                             "   and nfxr.EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                             "   and nfxr.Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                             "   and nfxr.EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                             "   and nfxr.EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                             "   and nfxr.Nota_Id         = " & Me.Codigo & vbCrLf &
                             "   and nfxr.Serie_Id        ='" & Me.Serie & "';" & vbCrLf
                    Sqls.Add(strSQL)


                    strSQL = " Update Pesagem Set Deposito = '" & Me.CodigoDeposito & "'," & vbCrLf &
                             "	               EndDeposito = " & Me.EnderecoDeposito & "," & vbCrLf &
                             "            UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "'," & vbCrLf &
                             "        UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                             "	Where Empresa_Id    ='" & Romaneio.CodigoEmpresa & "'" & vbCrLf &
                             "    and EndEmpresa_Id = " & Romaneio.EnderecoEmpresa & vbCrLf &
                             "    and Pesagem_id    = (Select Pesagem_id" & vbCrLf &
                             "                           From RomaneiosXPesagens" & vbCrLf &
                             "                          where Empresa_id    ='" & Romaneio.CodigoEmpresa & "'" & vbCrLf &
                             "                            and EndEmpresa_id = " & Romaneio.EnderecoEmpresa & vbCrLf &
                             "                            and Romaneio_id   = " & Romaneio.Codigo & ");"

                    Sqls.Add(strSQL)
                End If

                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                If Serie = "REC" Then
                    strSQL = "UPDATE NotasFiscais set Situacao = 3" & vbCrLf &
                             " WHERE Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                             "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                             "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                             "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                             "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                             "   and Nota_Id         = " & Me.Codigo & vbCrLf &
                             "   and Serie_Id        ='" & Me.Serie & "';"
                    Sqls.Add(strSQL)

                    If TemNotaTroca Then
                        SalvarNotasXNotas(Sqls)
                    End If
                Else
                    SalvarTabelasRelacionadasSql(Sqls)
                    strSQL = "DELETE NotasFiscais" & vbCrLf &
                             " WHERE  Empresa_Id     ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                             "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                             "	 and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                             "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                             "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                             "   and Nota_Id         = " & Me.Codigo & vbCrLf &
                             "   and Serie_Id        ='" & Me.Serie & "';"
                    Sqls.Add(strSQL)
                End If
            Case "C"
                If NFG Then
                    strSQL = "UPDATE NotasFiscais Set" & vbCrLf &
                             "    Situacao = 2 " & vbCrLf &
                             "	 ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                             "	 ,UsuarioAlteracaoData = getdate() " & vbCrLf &
                             " WHERE Empresa_Id      ='" & Me.CodigoEmpresa & "'" &
                             "   AND EndEmpresa_Id   = " & Me.EnderecoEmpresa &
                             "   AND Cliente_Id      ='" & Me.CodigoCliente & "'" &
                             "   AND EndCliente_Id   = " & Me.EnderecoCliente &
                             "   AND EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" &
                             "   AND Serie_Id        ='" & Me.Serie & "'" &
                             "   AND Nota_Id         = " & Me.Codigo
                    Sqls.Add(strSQL)

                    For Each row In VencimentosNota
                        row.IUD = "D"
                        row.SalvarSql(Sqls)
                        Sqls.Add(VencimentosNota.NotaxTituloSql(row.Codigo, "D"))
                    Next

                    Razao.ExcluiContabilizacaoNotaSQL(Sqls)

                ElseIf Eletronica And ChaveNFE.Length > 0 And ProtocoloNota.Length > 0 Then
                    strSQL = "INSERT INTO NFEPendencias (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, TpEmis, ChaveNfe, ObservacoesFiscais, DadosAdicionais, Protocolo)" &
                             "VALUES ('" & Me.CodigoEmpresa & "', " & Me.EnderecoEmpresa & ", " &
                             "'" & Me.CodigoCliente & "', " & Me.EnderecoCliente & ", " &
                             "'" & Me.EntradaSaida.ToString.Substring(0, 1) & "', '" & Me.Serie & "', " & Me.Codigo & ", " &
                             "'" & Me.DataNota.ToString("yyyy-MM-dd") & "', '" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "', " &
                             "'" & Me.Usuario & "', " &
                             "1, '" & Me.ChaveNFE & "', '" & Funcoes.EliminarCaracteresEspeciais(RTrim(Me.ObservacaoCancelamento)) & "', " &
                             "'CANCELAR', '" & Me.ProtocoloNota & "');"
                    Sqls.Add(strSQL)

                    strSQL = "UPDATE NotasFiscais Set" & vbCrLf &
                             "    Situacao             = 7 " & vbCrLf &
                             "	 ,UsuarioAlteracao     ='" & HttpContext.Current.Session("ssNomeUsuario") & "'" & vbCrLf &
                             "	 ,UsuarioAlteracaoData = getdate() " & vbCrLf &
                             " WHERE Empresa_Id      ='" & Me.CodigoEmpresa & "'" &
                             "   AND EndEmpresa_Id   = " & Me.EnderecoEmpresa &
                             "   AND Cliente_Id      ='" & Me.CodigoCliente & "'" &
                             "   AND EndCliente_Id   = " & Me.EnderecoCliente &
                             "   AND EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" &
                             "   AND Serie_Id        ='" & Me.Serie & "'" &
                             "   AND Nota_Id         = " & Me.Codigo
                    Sqls.Add(strSQL)

                End If
        End Select
    End Sub

    Public Sub SalvaAlteracaoDaNotaFiscal(ByRef Sqls As ArrayList)
        SalvarTabelasRelacionadasSql(Sqls)
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        'ITENS DA NOTA
        If Not Itens Is Nothing Then Itens.SalvarSql(Sqls)
        SalvarEmbalagens(Sqls)

        If Not Transportador Is Nothing AndAlso Transportador.Codigo.Length > 0 Then

            If XMLImportado Then

                Dim verificarTransportador As New Cliente(RemoveMask(Transportador.Codigo), 0)

                If verificarTransportador.Codigo Is Nothing OrElse verificarTransportador.Codigo.Length = 0 Then

                    Transportador.SalvarSql(Sqls)

                Else

                    'Inserimos o tipo transportador
                    If Transportador.Tipos.Where(Function(x) x.CodigoTipo = 7 And x.IUD = "I").Count() > 0 Then
                        Transportador.IUD = "U_TIPO"
                    End If

                End If

                Dim objPlaca As New [Lib].Negocio.Placa(Me.PlacaTransportador)

                If Not PlacaDetalhes Is Nothing Then
                    If objPlaca Is Nothing OrElse objPlaca.Placa01.Length = 0 OrElse objPlaca.IUD = "I" Then
                        objPlaca.Placa01 = PlacaDetalhes.Placa01
                        objPlaca.EstadoPlaca01 = PlacaDetalhes.EstadoPlaca01
                        objPlaca.SalvarSql(Sqls)
                    End If
                End If

            Else

                If Me.IUD = "I" Then
                    Transportador.SalvarSql(Sqls)
                End If

            End If

            SalvarTransportador(Sqls)

        End If

        If Not Me.PesoDeChegada Is Nothing Then
            Me.PesoDeChegada.SalvarSql(Sqls)
        End If

        'Caso tenha nr. na Chave deve ser gravada - Furlan - 25/08/2014
        ' If (NFG) OrElse (SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso ChaveNFE.Length > 0) Then
        If ChaveNFE.Length > 0 Then
            SalvarNFERealizada(Sqls)
        ElseIf Me.NFG Then
            SalvarNFERealizada(Sqls)
        End If

        'EXPORTAção
        If Not Me.DadosDaExportacao Is Nothing Then
            Me.DadosDaExportacao.IUD = Me.IUD
            Me.DadosDaExportacao.SalvarSql(Sqls)
        End If

        If Not Me.DadosDaExportacaoRE Is Nothing AndAlso Me.DadosDaExportacaoRE.Count > 0 Then Me.DadosDaExportacaoRE.SalvarSql(Sqls)

        'IMPORTAção
        If Not Me.DadosDaImportacao Is Nothing Then
            Me.DadosDaImportacao.IUD = Me.IUD
            Me.DadosDaImportacao.SalvarSql(Sqls)
        End If

        'COMISSéES
        If Me.NFG AndAlso Me.ComissoesXBaixas IsNot Nothing AndAlso Me.ComissoesXBaixas.Count > 0 Then
            Me.ComissoesXBaixas.SalvarSql(Sqls)
        End If

        'VENCIMENTOS
        'If FinanceiroNovo Then
        '    If SubOperacao.Financeiro AndAlso Titulos IsNot Nothing Then 'AndAlso (Titulos.Count > 0 OrElse Titulos.AdiantamentosAbertos.Count > 0) Then
        '        Titulos.SalvarSql(Sqls, False)
        '    End If
        'Else
        If SubOperacao.Financeiro OrElse VencimentosNota.Count > 0 Then
            If NFG Then
                If Me.IUD = "D" Then
                    For Each row In VencimentosNota.Where(Function(s) Not s.CodigoProvisao = 1)
                        row.IUD = "D"
                        row.SalvarSql(Sqls)
                        Sqls.Add(VencimentosNota.NotaxTituloSql(row.Codigo, "D"))
                    Next
                Else
                    If VencimentosNota.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.ValorDoDocumento) > 0 Then VencimentosNota.SalvarSQL(Sqls, False)
                End If
                'ElseIf SubOperacao.Devolucao AndAlso Pedido IsNot Nothing AndAlso Pedido.MomentoFinanceiro = eTipoFaturamento.NotaFiscal AndAlso SubOperacao.PrecoFixo AndAlso _
                '    (SubOperacao.Classe = eClassesOperacoes.VENDAS OrElse SubOperacao.Classe = eClassesOperacoes.COMPRAS) Then
                '    'Bloqueado porque a rotina não está fazendo correto para mais de 1 item - furlan - 31-07-2015
                '    If Me.Itens.Count = 1 Then
                '        VencimentosNota.DevolucaoNota(Sqls)
                '        VencimentosNota.SalvarSQL(Sqls)
                '    End If
            ElseIf (Me.SubOperacao.Classe = eClassesOperacoes.AFIXAR Or Me.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES) AndAlso
                Me.VencimentosPedido IsNot Nothing AndAlso Me.VencimentosPedido.Count > 0 Then
                For Each row In VencimentosPedido
                    If Me.CodigoEmpresa = row.CodigoEmpresaPedido AndAlso
                        Me.EnderecoEmpresa = row.EndEmpresaPedido AndAlso
                        Me.CodigoPedido = row.CodigoPedido AndAlso
                        Me.Itens(0).CodigoFixacao = row.CodigoPedidoFixacao _
                        AndAlso row.CodigoProvisao = 3 Then
                        Dim sql As String
                        If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                            sql = "Update ContasAPagar Set Provisao = 2 " & vbCrLf &
                                "Where Registro_id = " & row.Codigo
                        Else
                            sql = "Update ContasAReceber Set Provisao = 2 " & vbCrLf &
                                "Where Registro_id = " & row.Codigo
                        End If
                        Sqls.Add(sql)
                    End If
                Next
            ElseIf Me.SubOperacao.Devolucao Then
                VencimentosNota.SalvarSQL(Sqls, False)
            Else
                If VencimentosNota IsNot Nothing AndAlso VencimentosNota.Count > 0 Then
                    If Me.IUD = "D" Then VencimentosNota.ExcluirNota()

                    If Me.IUD = "I" Then
                        Dim parcela As Integer = 0
                        Dim parcelas As Integer = VencimentosNota.Count
                        For Each tit In VencimentosNota
                            parcela += 1
                            tit.Historico = "REF. NF " & Me.Codigo & "-" & Me.Serie & "-" & IIf(Me.EntradaSaida = eEntradaSaida.Saida, "S", "E") & ", Parcela " & parcela & "/" & parcelas & ", Pedido: " & Me.CodigoPedido & " / " & Me.Cliente.Nome
                        Next
                    End If

                    VencimentosNota.SalvarSQL(Sqls, False)
                End If

                If Not CarregandoNota AndAlso VencimentosPedido IsNot Nothing AndAlso VencimentosNota.Count > 0 Then
                    VencimentosPedido.SalvarSQL(Sqls)
                End If
            End If
        ElseIf Me.SubOperacao.Classe = eClassesOperacoes.AFIXAR Or Me.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
            If Me.VencimentosPedido IsNot Nothing AndAlso Me.VencimentosPedido.Count > 0 Then
                For Each row In VencimentosPedido
                    If Me.CodigoEmpresa = row.CodigoEmpresaPedido AndAlso
                        Me.EnderecoEmpresa = row.EndEmpresaPedido AndAlso
                        Me.CodigoPedido = row.CodigoPedido AndAlso
                        Me.Itens(0).CodigoFixacao = row.CodigoPedidoFixacao _
                        AndAlso row.CodigoProvisao = 3 Then
                        Dim sql As String
                        If Me.Pedido.SubOperacao.EntradaSaida = eEntradaSaida.Entrada Then
                            sql = "Update ContasAPagar Set Provisao = 2 " & vbCrLf &
                                "Where Registro_id = " & row.Codigo
                        Else
                            sql = "Update ContasAReceber Set Provisao = 2 " & vbCrLf &
                                "Where Registro_id = " & row.Codigo
                        End If
                        Sqls.Add(sql)
                    End If
                Next
            End If
        End If
        'End If

        'NOTAS X NOTAS - LISTA DE NOTAS ASSOCIADAS
        If Me.TipoDeDocumentoFrete Is Nothing _
        OrElse (Me.TipoDeDocumentoFrete IsNot Nothing AndAlso Me.TipoDeDocumentoFrete <> eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro) _
            OrElse (Me.TipoDeDocumentoFrete IsNot Nothing AndAlso Me.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.PrestacaoServicoTerceirosSemFinanceiro AndAlso Me.NotasTrocaOrigem IsNot Nothing AndAlso Me.NotasTrocaOrigem.Any(Function(s) s.TipoDeDocumentoFrete IsNot Nothing AndAlso s.TipoDeDocumentoFrete = eTipoDeDocumentoFrete.Circulacao)) Then

            If NotasTrocaOrigem IsNot Nothing AndAlso NotasTrocaOrigem.Count > 0 Then
                For Each Me._NotaTrocaOrigem In NotasTrocaOrigem
                    SalvarNotasXNotas(Sqls)
                Next
            Else
                If TemNotaTroca Then
                    SalvarNotasXNotas(Sqls)
                End If
            End If
        End If


        If Not CarregandoNota AndAlso Empresa.Empresa.NotaFiscalEletronica AndAlso Eletronica AndAlso NossaEmissao Then
            SalvarNFEPendencias(Sqls)
        End If

        If NFG Then
            Dim cLote As Integer = 0
            If Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC) OrElse
                Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CTRC_SEM_NF) OrElse
                 Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_TOM) OrElse
                Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E) OrElse
                Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.CT_E_OUT) OrElse
                Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.ComplementoDeFrete) OrElse
                Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Estadia) Then
                cLote = 21
            ElseIf Razao.NotaFiscalRazao.CodigoTipoDeDocumento = CInt(eTipoDeDocumento.Nota) Then
                cLote = 0
            Else
                cLote = 9
            End If

            If IUD = "I" Then
                Razao.ContabilizarNotaSql(Sqls, cLote)
            ElseIf IUD = "U" Then
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                Razao.ContabilizarNotaSql(Sqls, cLote)
                LancamentosContabeis = Nothing
            ElseIf IUD = "D" Then
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
            End If
        Else
            If Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CTRC) _
                    AndAlso Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ReciboDeFrete) _
                    AndAlso Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.ComplementoDeFrete) _
                    AndAlso Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Estadia) _
                    AndAlso Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CT_E) _
                    AndAlso Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.CT_E_OUT) _
                    AndAlso Razao.NotaFiscalRazao.CodigoTipoDeDocumento <> CInt(eTipoDeDocumento.Anulacao) Then
                If IUD = "I" Then
                    Razao.ContabilizarNotaSql(Sqls)
                ElseIf IUD = "U" Then
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                    Razao.ContabilizarNotaSql(Sqls)
                    LancamentosContabeis = Nothing
                ElseIf IUD = "D" Then
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                End If
            Else
                If IUD = "D" Then
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 9)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 10)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 11)
                    Razao.ExcluiContabilizacaoNotaSQL(Sqls, 21)
                End If
            End If
        End If

        If Me.NotasReferenciais IsNot Nothing AndAlso Me.NotasReferenciais.Count > 0 Then
            If Me.IUD = "U" Then
                For Each item In Me.NotasReferenciais
                    item.IUD = "D"
                    item.SalvarSql(Sqls)

                    item.IUD = "I"
                    item.SalvarSql(Sqls)
                Next
            Else
                For Each item In Me.NotasReferenciais
                    item.IUD = Me.IUD
                    item.SalvarSql(Sqls)
                Next
                'Me.NotasReferenciais.SalvarSql(Sqls)
            End If
        End If

        If Me.NotaFiscalXPercursos IsNot Nothing AndAlso Me.NotaFiscalXPercursos.Count > 0 Then
            Me.NotaFiscalXPercursos.SalvarSql(Sqls)
        End If

        If Me.CriarRomaneio Then
            Me.Romaneio.IUD = Me.IUD
            Me.Romaneio.CodigoEmpresa = _CodigoEmpresa
            Me.Romaneio.EnderecoEmpresa = _EnderecoEmpresa

            Me.Romaneio.CodigoOperacao = _CodigoOperacao
            Me.Romaneio.CodigoSubOperacao = _CodigoSubOperacao

            If Me.CodigoDeposito.Length > 0 Then
                Me.Romaneio.CodigoDeposito = _CodigoDeposito
                Me.Romaneio.EnderecoDeposito = _EnderecoDeposito
            Else
                Me.Romaneio.CodigoDeposito = _CodigoEmpresa
                Me.Romaneio.EnderecoDeposito = _EnderecoEmpresa
            End If

            If Me.CodigoDestino.Length > 0 Then
                Me.Romaneio.CodigoDestino = _CodigoDestino
                Me.Romaneio.EnderecoDestino = _EnderecoDestino
            Else
                Me.Romaneio.CodigoDestino = _CodigoCliente
                Me.Romaneio.EnderecoDestino = _EnderecoCliente
            End If

            If Me.CodigoTransbordo.Length > 0 Then
                Me.Romaneio.CodigoTransbordo = _CodigoTransbordo
                Me.Romaneio.EnderecoTransbordo = _EnderecoTransbordo
            Else
                Me.Romaneio.CodigoTransbordo = ""
                Me.Romaneio.EnderecoTransbordo = 0
            End If

            Me.Romaneio.Processo = "NOTA FISCAL"
            Me.Romaneio.EntradaSaida = _EntradaSaida.ToString.Substring(0, 1)
            Me.Romaneio.CodigoAutorizacao = _CodigoAutorizacao
            Me.Romaneio.Movimento = _Movimento
            Me.Romaneio.CodigoPedido = _CodigoPedido
            Me.Romaneio.CodigoProduto = _Itens(0).CodigoProduto

            Dim N As New Negocio.Numerador(_CodigoEmpresa, _EnderecoEmpresa, 110)
            Sqls.Add(N.IncrementarNumeradorSql)
            Me.CodigoRomaneio = N.Sequencia + 1
            Me.Romaneio.Codigo = _CodigoRomaneio
            Romaneio.SalvarSql(Sqls, False)
            SalvarNotasXRomaneio(Sqls)

        ElseIf Me.CodigoRomaneio > 0 Then
            Romaneio.Codigo = Me.CodigoRomaneio
            If IUD = "D" Then
                SalvarNotasXRomaneio(Sqls)
                If Romaneio.Pesagens.Count = 0 Then
                    Romaneio.IUD = "D"
                    Romaneio.SalvarSql(Sqls, False)
                End If
            ElseIf IUD = "I" And Romaneio.Pesagens.Count = 0 Then
                Romaneio.IUD = "I"
                Romaneio.SalvarSql(Sqls, False)
                SalvarNotasXRomaneio(Sqls)
            Else
                SalvarNotasXRomaneio(Sqls)
            End If
        End If

        'Utilizado para lanéar o peso de chegada automaticamente quando a nf for de entrada FOB e tiver romaneio
        If Not Me.NFG AndAlso Not Me.Romaneio Is Nothing AndAlso Not Me.Romaneio.Pesagens Is Nothing AndAlso Me.Romaneio.Pesagens.Count > 0 AndAlso Me.EntradaSaida = eEntradaSaida.Entrada AndAlso Me.TipoFreteSefaz = eTiposFrete.FOB Then

            If Me.IUD = "I" Then
                Me.PesoDeChegada.IUD = "I"
                Me.PesoDeChegada.PesoBruto = Me.Romaneio.PesoBruto
                Me.PesoDeChegada.PesoLiquido = Me.Romaneio.PesoLiquido
                Me.PesoDeChegada.Movimento = Me.Movimento
                Me.PesoDeChegada.Sinistro = False
                Me.PesoDeChegada.TarifaFrete = 0
                Me.PesoDeChegada.SalvarSql(Sqls)
            ElseIf Me.IUD = "D" Then
                Me.PesoDeChegada.IUD = "D"
                Me.PesoDeChegada.SalvarSql(Sqls)
            End If

        End If

        If Me.Arquivos IsNot Nothing AndAlso Me.Arquivos.Count > 0 Then
            For Each arq As [Lib].Negocio.Arquivo In Me.Arquivos
                Select Case Me.IUD
                    Case "I"
                        arq.IUD = "I"
                        arq.CodigoEmpresa = Me.CodigoEmpresa
                        arq.EnderecoEmpresa = Me.EnderecoEmpresa
                        arq.CodigoCliente = Me.CodigoCliente
                        arq.EnderecoCliente = Me.EnderecoCliente
                        arq.CodigoNota = Me.Codigo
                        arq.Serie = Me.Serie
                        arq.CodigoPedido = Me.CodigoPedido
                        arq.SalvarSql(Sqls)
                    Case "U"
                        arq.CodigoEmpresa = Me.CodigoEmpresa
                        arq.EnderecoEmpresa = Me.EnderecoEmpresa
                        arq.CodigoCliente = Me.CodigoCliente
                        arq.EnderecoCliente = Me.EnderecoCliente
                        arq.CodigoNota = Me.Codigo
                        arq.Serie = Me.Serie
                        arq.CodigoPedido = Me.CodigoPedido
                        arq.SalvarSql(Sqls)
                    Case "D"
                        arq.IUD = "D"
                        arq.CodigoEmpresa = Me.CodigoEmpresa
                        arq.EnderecoEmpresa = Me.EnderecoEmpresa
                        arq.CodigoCliente = Me.CodigoCliente
                        arq.EnderecoCliente = Me.EnderecoCliente
                        arq.CodigoNota = Me.Codigo
                        arq.Serie = Me.Serie
                        arq.CodigoPedido = Me.CodigoPedido
                        arq.SalvarSql(Sqls)
                End Select
            Next
        End If
    End Sub

    Public Function RemoveMask(ByVal pCpfCnpj As String) As String
        'Return pCpfCnpj.Replace(".", "").Replace("-", "").Replace("/", "").Replace("(", "").Replace(")", "").Replace(" ", "")
        Return String.Join("", System.Text.RegularExpressions.Regex.Split(pCpfCnpj, "[^\d]"))
    End Function

    Private Sub SalvarEmbalagens(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "Insert Into NotasXEmbalagens (Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Quantidade, Especie, Marca, Numero, PesoBruto, PesoLiquido)" & vbCrLf &
                      "Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & "," & Str(Me.Quantidade) & ",'" & Me.Especie & "','" & Me.Marca & "','" & Me.Numero & "'," & Str(Me.PesoBruto) & "," & Str(Me.PesoLiquido) & ")"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete From NotasXEmbalagens" & vbCrLf &
                      " Where Empresa_Id     ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "  and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                      "  and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                      "  and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                      "  and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "  and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                      "  and Nota_Id         = " & Me.Codigo
                Sqls.Add(sql)
        End Select
    End Sub

    Public Sub SalvarNotasXNotas(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "Insert Into NotasXNotas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, OrigemEmpresa_Id, OrigemEndEmpresa_Id, OrigemCliente_Id, OrigemEndCliente_Id, OrigemEntradaSaida_Id, OrigemSerie_Id, OrigemNota_Id)"
                sql &= "Values('" & NotaTrocaDestino.CodigoEmpresa & "'," & NotaTrocaDestino.EnderecoEmpresa & ",'" & NotaTrocaDestino.CodigoCliente & "'," & NotaTrocaDestino.EnderecoCliente & ",'" & NotaTrocaDestino.EntradaSaida.ToString.Substring(0, 1) & "','" & NotaTrocaDestino.Serie & "'," & NotaTrocaDestino.Codigo & ",'" & NotaTrocaOrigem.CodigoEmpresa & "'," & NotaTrocaOrigem.EnderecoEmpresa & ",'" & NotaTrocaOrigem.CodigoCliente & "'," & NotaTrocaOrigem.EnderecoCliente & ",'" & NotaTrocaOrigem.EntradaSaida.ToString.Substring(0, 1) & "','" & NotaTrocaOrigem.Serie & "'," & NotaTrocaOrigem.Codigo & ")"
                Sqls.Add(sql)


            Case "U"

                sql = "Delete NotasXNotas" & vbCrLf &
                      " Where Empresa_Id      ='" & NotaTrocaDestino.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & NotaTrocaDestino.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & NotaTrocaDestino.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & NotaTrocaDestino.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & NotaTrocaDestino.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & NotaTrocaDestino.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & NotaTrocaDestino.Codigo & vbCrLf &
                      "   and OrigemEmpresa_Id      ='" & NotaTrocaOrigem.CodigoEmpresa & "'" & vbCrLf &
                      "   and OrigemEndEmpresa_Id   = " & NotaTrocaOrigem.EnderecoEmpresa & vbCrLf &
                      "   and OrigemCliente_Id      ='" & NotaTrocaOrigem.CodigoCliente & "'" & vbCrLf &
                      "   and OrigemEndCliente_Id   = " & NotaTrocaOrigem.EnderecoCliente & vbCrLf &
                      "   and OrigemEntradaSaida_Id ='" & NotaTrocaOrigem.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and OrigemSerie_Id        ='" & NotaTrocaOrigem.Serie & "'" & vbCrLf &
                      "   and OrigemNota_Id         = " & NotaTrocaOrigem.Codigo & vbCrLf
                Sqls.Add(sql)

                sql = "Insert Into NotasXNotas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, OrigemEmpresa_Id, OrigemEndEmpresa_Id, OrigemCliente_Id, OrigemEndCliente_Id, OrigemEntradaSaida_Id, OrigemSerie_Id, OrigemNota_Id)"
                sql &= "Values('" & NotaTrocaDestino.CodigoEmpresa & "'," & NotaTrocaDestino.EnderecoEmpresa & ",'" & NotaTrocaDestino.CodigoCliente & "'," & NotaTrocaDestino.EnderecoCliente & ",'" & NotaTrocaDestino.EntradaSaida.ToString.Substring(0, 1) & "','" & NotaTrocaDestino.Serie & "'," & NotaTrocaDestino.Codigo & ",'" & NotaTrocaOrigem.CodigoEmpresa & "'," & NotaTrocaOrigem.EnderecoEmpresa & ",'" & NotaTrocaOrigem.CodigoCliente & "'," & NotaTrocaOrigem.EnderecoCliente & ",'" & NotaTrocaOrigem.EntradaSaida.ToString.Substring(0, 1) & "','" & NotaTrocaOrigem.Serie & "'," & NotaTrocaOrigem.Codigo & ")"
                Sqls.Add(sql)

            Case "D"
                sql = "Delete NotasXNotas" & vbCrLf &
                      " Where Empresa_Id      ='" & NotaTrocaDestino.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & NotaTrocaDestino.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & NotaTrocaDestino.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & NotaTrocaDestino.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & NotaTrocaDestino.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & NotaTrocaDestino.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & NotaTrocaDestino.Codigo & vbCrLf &
                      "   and OrigemEmpresa_Id      ='" & NotaTrocaOrigem.CodigoEmpresa & "'" & vbCrLf &
                      "   and OrigemEndEmpresa_Id   = " & NotaTrocaOrigem.EnderecoEmpresa & vbCrLf &
                      "   and OrigemCliente_Id      ='" & NotaTrocaOrigem.CodigoCliente & "'" & vbCrLf &
                      "   and OrigemEndCliente_Id   = " & NotaTrocaOrigem.EnderecoCliente & vbCrLf &
                      "   and OrigemEntradaSaida_Id ='" & NotaTrocaOrigem.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and OrigemSerie_Id        ='" & NotaTrocaOrigem.Serie & "'" & vbCrLf &
                      "   and OrigemNota_Id         = " & NotaTrocaOrigem.Codigo & vbCrLf
                Sqls.Add(sql)
        End Select
    End Sub

    Private Sub SalvarNFERealizada(ByRef Sqls As ArrayList)
        Dim sql As String

        Select Case Me.IUD
            Case "I"
                sql = " Insert into NFERealizadas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, ChaveNfe, ObservacoesFiscais, DadosAdicionais)" & vbCrLf &
                      " Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "','" & HttpContext.Current.Session("ssNomeUsuario") & "','" & Me.ChaveNFE & "','" & Me.Observacoes & " " & Me.ObservacoesDeEmbarque & "','" & Me.ObservacoesDoProduto & "');"
                Sqls.Add(sql)
            Case "U"
                If NFG Then

                    sql = "Delete NFERealizadas" & vbCrLf &
                              " where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                              "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                              "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                              "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                              "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                              "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                              "   and Nota_Id         = " & Me.Codigo
                    Sqls.Add(sql)

                    If (String.IsNullOrEmpty(Me.NotaFiscalOriginal.ChaveNFE) AndAlso Not String.IsNullOrEmpty(Me.ChaveNFE)) Then
                        sql = " Insert into NFERealizadas(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, ChaveNfe, SegCodBarra, ObservacoesFiscais, DadosAdicionais)" & vbCrLf &
                              " Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Now().ToString("yyyy-MM-dd HH:mm:ss") & "','" & HttpContext.Current.Session("ssNomeUsuario") & "','" & Me.ChaveNFE & "','" & Me.SegCodBarra & "','" & Me.Observacoes & " " & Me.ObservacoesDeEmbarque & "','" & Me.ObservacoesDoProduto & "');"
                    ElseIf _ChaveNFE.Length > 0 Then
                        sql = "update NFERealizadas set" & vbCrLf &
                              "   ChaveNfe = '" & Me.ChaveNFE & "'," & vbCrLf &
                              "   SegCodBarra = '" & Me.SegCodBarra & "'" & vbCrLf &
                              " where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                              "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                              "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                              "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                              "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                              "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                              "   and Nota_Id         = " & Me.Codigo
                        'Else
                        '    sql = "Delete NFERealizadas" & vbCrLf &
                        '          " where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                        '          "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                        '          "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                        '          "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                        '          "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                        '          "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                        '          "   and Nota_Id         = " & Me.Codigo
                    End If

                    Sqls.Add(sql)

                End If

            Case "D"

                sql = "Delete NFERealizadas" & vbCrLf &
                      " where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & Me.Codigo
                Sqls.Add(sql)


                'COMENTEI PORQUE ESTAVA REMOVENDO OS XML'S IMPORTADOS, NÃO APARECEIA MAIS PARA LANÇAR NOVAMENTE - FURLAN - 12/08/2024
                'sql = "Delete DocumentoXML " & vbCrLf &
                '      " where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                '      "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                '      "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                '      "   and Numero_Id       = " & Me.Codigo & vbCrLf &
                '      "   and Chave_Id        = '" & Me.ChaveNFE & "'"
                'Sqls.Add(sql)

        End Select
    End Sub

    Private Sub SalvarTransportador(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "Insert Into NotasFiscaisXTransportadores(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Proprietario, EndProprietario, Motorista, EndMotorista, Placa)" & vbCrLf &
                      "Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & ",'" & Me.CodigoTransportador & "'," & Me.EnderecoTransportador & ","
                If PlacaDetalhes Is Nothing Then
                    sql &= "'',0,'');"
                Else

                    If PlacaDetalhes.Motorista Is Nothing Then
                        sql &= "'',0,'" & PlacaDetalhes.Placa01 & "');"
                    Else
                        sql &= "'" & PlacaDetalhes.Motorista.Codigo & "'," & PlacaDetalhes.Motorista.CodigoEndereco & ",'" & PlacaDetalhes.Placa01 & "');"
                    End If

                End If
                Sqls.Add(sql)
            Case "D"
                sql = "Delete NotasFiscaisXTransportadores" & vbCrLf &
                      " Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                      "   and Nota_Id         = " & Me.Codigo
                Sqls.Add(sql)
        End Select
    End Sub

    Public Sub AtualizarObservacoes()

        Dim Obs As New System.Text.StringBuilder()

        ' Acumuladores para CBS e IBS
        Dim totalBaseCBS As Decimal = 0D
        Dim totalValorCBS As Decimal = 0D
        Dim aliquotaCBS As Decimal = 0D

        Dim totalBaseIBS As Decimal = 0D
        Dim totalValorIBS As Decimal = 0D
        Dim aliquotaIBS As Decimal = 0D

        For Each item As Negocio.NotaFiscalXItem In Me.Itens
            For Each enc As NotaFiscalXItemXEncargo In item.Encargos

                If enc.Codigo <> "LIQUIDO" Then

                    ' Descrição fiscal
                    If Not String.IsNullOrEmpty(enc.DescricaoFiscal) AndAlso
                   Not Obs.ToString().Contains(enc.DescricaoFiscal) Then

                        Obs.Append(" | " & enc.DescricaoFiscal)
                    End If

                    ' Observação fiscal
                    If Not String.IsNullOrEmpty(enc.ObservacaoFiscal) AndAlso
                   Not Obs.ToString().Contains(enc.ObservacaoFiscal) Then

                        Obs.Append(" | " & enc.ObservacaoFiscal)
                    End If

                    ' Acumula CBS
                    If enc.Codigo = "CBS" Then
                        totalBaseCBS += enc.Base
                        totalValorCBS += enc.Valor
                        If aliquotaCBS = 0D Then
                            aliquotaCBS = enc.Percentual
                        End If
                    End If

                    ' Acumula IBS
                    If enc.Codigo = "IBS" Then
                        totalBaseIBS += enc.Base
                        totalValorIBS += enc.Valor
                        If aliquotaIBS = 0D Then
                            aliquotaIBS = enc.Percentual
                        End If
                    End If

                End If

            Next
        Next

        If totalBaseIBS > 0D Then
            Dim textoIBS As String = String.Format(
            " | Tributação IBS: Base {0:C} - Alíquota {1:N1}% - Valor {2:C} (EC 132/2023)",
            totalBaseIBS, aliquotaIBS, totalValorIBS)

            If Not Obs.ToString().Contains("Tributação IBS:") Then
                Obs.Append(textoIBS)
            End If
        End If

        If totalBaseCBS > 0D Then
            Dim textoCBS As String = String.Format(
            " | Tributação CBS: Base {0:C} - Alíquota {1:N1}% - Valor {2:C} (EC 132/2023)",
            totalBaseCBS, aliquotaCBS, totalValorCBS)

            If Not Obs.ToString().Contains("Tributação CBS:") Then
                Obs.Append(textoCBS)
            End If
        End If

        _Observacoes = Obs.ToString()
    End Sub

    Private Sub SalvarNFEPendencias(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert into NFEPendencias(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Data, Hora, Usuario, TpEmis, ObservacoesFiscais, DadosAdicionais)" & vbCrLf &
                      " Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & ",'" & Me.Movimento.ToString("yyyy-MM-dd") & "','" & Me.DataInclusao.ToString("yyyy-MM-dd HH:mm:ss") & "','" & HttpContext.Current.Session("ssNomeUsuario") & "',1 ,'" & Me.Observacoes & " " & Me.ObservacoesDeEmbarque & "','" & Me.ObservacoesDoProduto & "');"
                Sqls.Add(sql)
            Case "D"
                sql = "Delete from NFEPendencias" & vbCrLf &
                      " Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "   and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                      "   and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                      "   and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                      "   and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "   and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                      "   and Nota_Id         ='" & Me.Codigo & "';"
                Sqls.Add(sql)
        End Select
    End Sub

    Public Sub SalvarNotasXRomaneio(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert into NotasFiscaisXRomaneios(Empresa_Id, EndEmpresa_Id, Cliente_Id, EndCliente_Id, EntradaSaida_Id, Serie_Id, Nota_Id, Romaneio_Id)" & vbCrLf &
                      " Values('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & ",'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.EntradaSaida.ToString.Substring(0, 1) & "','" & Me.Serie & "'," & Me.Codigo & "," & Me.CodigoRomaneio & ");"
                Sqls.Add(sql)
            Case "D"
                sql = " Delete NotasFiscaisXRomaneios" & vbCrLf &
                      "  Where Empresa_Id      ='" & Me.CodigoEmpresa & "'" & vbCrLf &
                      "    and EndEmpresa_Id   = " & Me.EnderecoEmpresa & vbCrLf &
                      "    and Cliente_Id      ='" & Me.CodigoCliente & "'" & vbCrLf &
                      "    and EndCliente_Id   = " & Me.EnderecoCliente & vbCrLf &
                      "    and EntradaSaida_Id ='" & Me.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                      "    and Serie_Id        ='" & Me.Serie & "'" & vbCrLf &
                      "    and Nota_Id         = " & Me.Codigo & vbCrLf &
                      "    and Romaneio_Id     = " & Me.CodigoRomaneio
                Sqls.Add(sql)
        End Select
    End Sub

    Public Function AtualizarTrocaDeNota() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Dim strSQL As String = ""

        SalvarNotasXNotas(Sqls)

        If Me.IUD = "I" Then
            strSQL = "Update NotasFiscais set Troca = 1, " & vbCrLf
        ElseIf Me.IUD = "D" Then
            strSQL = "Update NotasFiscais set Troca = 0, " & vbCrLf
        End If

        strSQL &= "       UsuarioAlteracao = '" & HttpContext.Current.Session("ssNomeUsuario") & "', " & vbCrLf &
                  "       UsuarioAlteracaoData = '" & Now().ToString("yyyy-MM-dd") & "'" & vbCrLf &
                  " Where Empresa_Id      ='" & NotaTrocaOrigem.CodigoEmpresa & "'" & vbCrLf &
                  "   and EndEmpresa_Id   = " & NotaTrocaOrigem.EnderecoEmpresa & vbCrLf &
                  "   and Cliente_Id      ='" & NotaTrocaOrigem.CodigoCliente & "'" & vbCrLf &
                  "   and EndCliente_Id   = " & NotaTrocaOrigem.EnderecoCliente & vbCrLf &
                  "   and EntradaSaida_Id ='" & NotaTrocaOrigem.EntradaSaida.ToString.Substring(0, 1) & "'" & vbCrLf &
                  "   and Serie_Id        ='" & NotaTrocaOrigem.Serie & "'" & vbCrLf &
                  "   and Nota_Id         = " & NotaTrocaOrigem.Codigo

        Sqls.Add(strSQL)

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub VerificaCartaoPamcard()
        Dim sql As String = "SELECT CartaoPgtoFrete " & vbCrLf &
                            "FROM  NotasFiscais " & vbCrLf &
                            "WHERE TipoDeDocumento = 2 AND Situacao in (1,4) AND CartaoPgtoFrete = '" & Me.CartaoPgtoFrete & "'"

        Dim Banco As New AcessaBanco
        Dim ds As New DataSet

        ds = Banco.ConsultaDataSet(sql, "NotaXPamcard")

        If ds Is Nothing OrElse ds.Tables(0).Rows.Count = 0 Then
            CartaoNovo = True
        Else
            CartaoNovo = False
        End If
    End Sub

    Public Sub getSqlException(ByRef Sqls As ArrayList)
        Dim sql As String = "BEGIN TRY " & vbCrLf &
               "DECLARE @HORA_BLOQUEIO AS DATETIME = DATEADD(MINUTE, 3, (SELECT VERSAOHORARIOBLOQUEIO FROM PEDIDOS WHERE PEDIDO_ID = '" & Me.Pedido.Codigo & "' AND EMPRESA_ID = '" & Me.Pedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.Pedido.EnderecoEmpresa & "')) " & vbCrLf &
               "PRINT 'HORA_ATUAL: ' + CAST(GETDATE() AS VARCHAR); " & vbCrLf &
               "PRINT 'HORA_BLOQUEIO: ' + CAST(@HORA_BLOQUEIO AS VARCHAR); " & vbCrLf &
               "IF (GETDATE() > @HORA_BLOQUEIO) " & vbCrLf &
               "BEGIN " & vbCrLf &
               "RAISERROR ('POR FAVOR, ATUALIZE O SALDO FINANCEIRO PARA REALIZAR ESTA Ação!', " & vbCrLf &
               "16, " & vbCrLf &
               "1); " & vbCrLf &
               "END " & vbCrLf &
               "UPDATE PEDIDOS SET VERSAOHORARIOBLOQUEIO = NULL WHERE PEDIDO_ID = '" & Me.Pedido.Codigo & "' AND EMPRESA_ID = '" & Me.Pedido.CodigoEmpresa & "' AND ENDEMPRESA_ID = '" & Me.Pedido.EnderecoEmpresa & "'; " & vbCrLf &
               "END TRY " & vbCrLf &
               "BEGIN CATCH " & vbCrLf &
               "DECLARE @ErrorMessage NVARCHAR(4000); " & vbCrLf &
               "DECLARE @ErrorSeverity INT; " & vbCrLf &
               "DECLARE @ErrorState INT; " & vbCrLf &
               "SELECT " & vbCrLf &
               "@ErrorMessage = ERROR_MESSAGE(), " & vbCrLf &
               "@ErrorSeverity = ERROR_SEVERITY(), " & vbCrLf &
               "@ErrorState = ERROR_STATE(); " & vbCrLf &
               "RAISERROR (@ErrorMessage, " & vbCrLf &
               "@ErrorSeverity, " & vbCrLf &
               "@ErrorState); " & vbCrLf &
               "END CATCH;"
        Sqls.Add(sql)
    End Sub

    Public Function ConferenciaNotas() As ListNotasFiscais
        Dim strSQL As String = "SELECT DISTINCT Empresa_Id, EndEmpresa_Id, Nota_Id, Serie_Id, Cliente_Id, EndCliente_Id FROM VW_ConferenciaNotas "
        Dim objBanco As New AcessaBanco
        Dim listNF As New ListNotasFiscais
        Dim dsNotasFiscais As DataSet = objBanco.ConsultaDataSet(strSQL, "NotasFiscais")
        For Each row As DataRow In dsNotasFiscais.Tables(0).Rows
            Dim NF As New Negocio.NotaFiscal
            NF.Codigo = row("Nota_Id")
            NF.Serie = row("Serie_Id")
            NF.CodigoEmpresa = row("Empresa_Id")
            NF.EnderecoEmpresa = row("EndEmpresa_Id")
            NF.CodigoCliente = row("Cliente_Id")
            NF.EnderecoCliente = row("EndCliente_Id")
            listNF.Add(NF)
        Next
        Return listNF
    End Function

    Public Function ValidaNotaRetroativa() As String
        Dim sql As String
        Dim objBanco As New AcessaBanco
        Dim resposta As String = String.Empty

        sql = "SELECT 1" & vbCrLf &
              " FROM NotasFiscais" & vbCrLf &
              " WHERE NossaEmissao  ='S'" & vbCrLf &
              "   AND Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf &
              "   AND EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "   AND Serie_Id      ='" & Me.Serie & "'" & vbCrLf &
              "   AND Movimento     >'" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf
        Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "NotasFiscais")

        If ds.Tables(0).Rows.Count >= 1 Then
            resposta = "Jé existe(m) nota(s) com data superior a esta!"
        End If

        If DateDiff("d", Me.Movimento, Now) = 0 Then
            resposta = "A Nota Fiscal Retroativa não pode ser na data atual."
        End If

        If DateDiff("d", Me.Movimento, Now) > Me.Empresa.Empresa.DiasNFRetroativa + System.DateTime.DaysInMonth(Year(Now), Month(Now.AddMonths(-1))) Then
            resposta = "Esta data ultrapassa os limites para Notas Fiscais Retroativas."
        End If

        If Me.Movimento.Month <> Month(Now()) AndAlso Me.Movimento.Year = Year(Now()) Then
            If DateDiff("d", Me.Movimento, Now) > Me.Empresa.Empresa.DiasNFRetroativa Then
                resposta = "A quantidade de dias para Notas Fiscais retroativas no més subsequente é limitada a " & Me.Empresa.Empresa.DiasNFRetroativa & ". (Configurado no cadastro da Empresa)"
            End If
        End If

        Return resposta

    End Function

    Public Function ValidarNotaFiscal(ByVal dsXML As DataSet, ByVal bTabImportacao As Boolean, ByRef erroMsg As String, ByVal emissaoNossaImportacaoXML As Boolean) As Boolean

        'Verificar se foi feito a recusa da Nota Fiscal
        If Me.TemRecusa Then
            erroMsg = "Nota Fiscal não pode ser lançada pois a mesma foi lançada como recusada."
            Return False
        End If

        'Verifica se o operação está com CLASSE cadastrada.
        If String.IsNullOrWhiteSpace(Me.Operacao.CodigoClasse) Then
            erroMsg = "Operação: " & Me.CodigoOperacao & " - " & Me.Operacao.Descricao & ". está sem Classe cadastrada!"
            Return False
        End If

        If Me.MicSistemas = False AndAlso Me.NotaDeTerceiro = False AndAlso Not Funcoes.VerificaAcesso(Me.Empresa.Codigo, Me.Empresa.CodigoEndereco, Me.Movimento, "NOTAS FISCAIS") Then
            erroMsg = "Movimento da Nota Fiscal jé Fechado para esta data."
            Return False
        End If

        If Me.MicSistemas = False AndAlso Me.NotaDeTerceiro = False AndAlso Not Funcoes.VerificaAcesso(Me.Empresa.Codigo, Me.Empresa.CodigoEndereco, Me.Movimento, "CONTABIL") Then
            erroMsg = "Movimento Contébil jé Fechado para esta data."
            Return False
        End If

        If Not Me.NossaEmissao AndAlso Me.EntradaSaida = eEntradaSaida.Entrada AndAlso Me.Eletronica AndAlso Me.Empresa.Empresa.ObrigaChaveNf AndAlso Not Me.ChaveNFE.Length = 44 Then
            erroMsg = "Obrigatério a informação da Chave da Nota Eletrônica do Fornecedor."
            Return False
        End If

        'Solicitação via e-mail em 04/09/2023 pela Nutri - Furlan - 14/09/2023
        If Me.CodigoEmpresa = "05366261000224" AndAlso Me.EntradaSaida = eEntradaSaida.Saida AndAlso Me.CodigoLocalEmbarque.Length = 0 Then
            erroMsg = "Local de Embarque não foi informado, o mesmo é obrigatério para o Estado do MS."
            Return False
        End If

        'Verifica se a suboperação tem a Situação PIS/Cofins cadastrada.

        'Verifica se a soma das notas de relacionadas na devolucao batem com a nota de Devolucao.
        If Me.SubOperacao.Devolucao AndAlso Not Me.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") _
            AndAlso ([Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.DEPOSITOS Or
                     [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.COMPRAS Or
                     [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.COMPRASAORDEM Or
                     [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.EXPORTACOES Or
                     [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.VENDAS Or
                     [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.VENDASAORDEM) Then

            For Each item In Me.Itens

                If item.CodigoEmbalagem = 1 AndAlso item.QuantidadeFiscal > 0 Then
                    'ainda vou testar - Furlan - 23/08/2022
                    'Alterado dia 24/05/25 - Felipe
                    Dim vString As String = CStr(item.QuantidadeFiscal.ToString("F4"))
                    Dim vDecimal As Decimal = CDec(vString.Split(",")(1))

                    If vDecimal > 0 Then
                        erroMsg = "Produto a GRANEL não pode ter casa decimal." & vbCrLf
                        Return False
                    End If

                    If item.QuantidadeFisica > 0 Then
                        vString = CStr(item.QuantidadeFisica.ToString("F4"))
                        vDecimal = CDec(vString.Split(",")(1))

                        If vDecimal > 0 Then
                            erroMsg = "Produto a GRANEL não pode ter casa decimal." & vbCrLf
                            Return False
                        End If
                    End If
                End If

                If Me.SubOperacao.QuantidadeFiscal AndAlso item.NotasDevolucao.Count = 0 Then
                    erroMsg = "não foi selecionado nenhuma nota fiscal do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome
                    Return False
                End If

                If dsXML Is Nothing AndAlso Me.SubOperacao.QuantidadeFiscal AndAlso item.NotasDevolucao.SomaQtde <> item.QuantidadeFiscal Then
                    erroMsg = "A Soma das Quantidades das notas devolvidas nao batem com a quantidade do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
                    Return False
                End If

                If dsXML Is Nothing AndAlso Math.Round(item.NotasDevolucao.SomaVlr, 2, MidpointRounding.AwayFromZero) <> Math.Round(item.ValorTotal, 2, MidpointRounding.AwayFromZero) Then
                    erroMsg = "A Soma dos Valores das notas devolvidas não batem com o Valor do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
                    Return False
                End If

                Dim qDev As Decimal = 0
                Dim vDev As Decimal = 0
                Dim vDevLiq As Decimal = 0

                For Each dev In item.NotasDevolucao
                    qDev += dev.QuantidadeDevolucao
                    vDev += dev.ValorDevolucao
                    If Me.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse Not Me.SubOperacao.Financeiro Then
                        dev.ValorLiquidoDevolucao = dev.ValorDevolucao
                    End If
                    vDevLiq += dev.ValorLiquidoDevolucao
                Next

                If qDev <> item.QuantidadeFiscal Then
                    erroMsg = "A Soma das Quantidades das notas devolvidas nao batem com a quantidade do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
                    Return False
                End If

                If vDev <> item.ValorTotal Then
                    erroMsg = "A Soma dos Valores das notas devolvidas não batem com o Valor do item Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
                    Return False
                End If

                'COMENTADO PARA REVISARMOS OUTRA MANEIRA DE PEGAR O LIQUIDO DA DEVOLUÇÃO
                'If vDevLiq <> item.ValorLiquido Then
                '    erroMsg = "A Soma dos Valores Líquidos das notas devolvidas não batem com o Valor do item Líquido Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
                '    Return False
                'End If

                'If vDevLiq <> item.Encargos.EncLiquido.Valor Then
                '    erroMsg = "A Soma dos Valores Líquidos das notas devolvidas não batem com o Valor do item Líquido Devolvido: " & item.CodigoProduto & " - " & item.Produto.Nome & " - Sequência: " & item.Sequencia
                '    Return False
                'End If

                If Me.Empresa.CodigoEstado = "PR" AndAlso Me.EntradaSaida = eEntradaSaida.Saida AndAlso Me.ValorIcms = 0 Then
                    If String.IsNullOrEmpty(item.OperacaoEstado.CodigoBeneficio) Then
                        erroMsg = "Código do Benefício não foi informado, verifique na Operação X Encargos o cédigo: " & item.OperacaoEstado.Codigo
                        Return False
                    End If
                End If

                If Me.Empresa.CodigoEstado = "PR" AndAlso Me.EntradaSaida = eEntradaSaida.Saida AndAlso Me.ValorIcms > 0 Then
                    If Not String.IsNullOrEmpty(item.OperacaoEstado.CodigoBeneficio) Then
                        erroMsg = "Código do Benefício não pode ser informado, verifique na Operação X Encargos o cédigo: " & item.OperacaoEstado.Codigo
                        Return False
                    End If
                End If
            Next
        End If

        'Repete a Validacao da chave da nota caso seja eletronica S e nao seja nossa emissao
        If Me.Eletronica And Not Me.NossaEmissao Then

            Dim valida As Boolean = True

            If (Mid(Me.ChaveNFE, 7, 14) = "02935843000105" Or
                Mid(Me.ChaveNFE, 7, 14) = "01409655000180" Or
                Mid(Me.ChaveNFE, 7, 14) = "78393592000146" Or
                Mid(Me.ChaveNFE, 7, 14) = "82951310000156" Or
                Mid(Me.ChaveNFE, 7, 14) = "87958674000181" Or
                Mid(Me.ChaveNFE, 7, 14) = "76416890000189" Or
                Mid(Me.ChaveNFE, 7, 14) = "12200192000169" Or
                Mid(Me.ChaveNFE, 7, 14) = "58290502000184" Or
                Mid(Me.ChaveNFE, 7, 14) = "03507415000578") Then
                valida = False
            End If

            Me.ChaveNFE = Me.ChaveNFE.Replace(".", "")

            If valida AndAlso Me.CodigoCliente.Length = 11 AndAlso Not Mid(Me.ChaveNFE, 10, 11) = Me.CodigoCliente AndAlso emissaoNossaImportacaoXML = False Then
                erroMsg = "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf valida AndAlso Me.CodigoCliente.Length = 14 AndAlso Not Mid(Me.ChaveNFE, 7, 14) = Me.CodigoCliente AndAlso emissaoNossaImportacaoXML = False Then
                erroMsg = "Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Me.Cliente.Municipio.EstadoIbge = 0 Then
                erroMsg = "Tem algum problema no cadastro do municipio do Cliente, código do IBGE não pode ser zero!"
                Return False
            ElseIf valida AndAlso Not CInt(Left(Me.ChaveNFE, 2)) = Me.Cliente.Municipio.EstadoIbge And emissaoNossaImportacaoXML = False Then
                erroMsg = "Estado do Cliente na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not CInt(Mid(Me.ChaveNFE, 26, 9)) = CInt(Me.Codigo) Then
                erroMsg = "Número da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not CInt(Mid(Me.ChaveNFE, 23, 3)) = CInt(Me.Serie) Then
                erroMsg = "Série da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
            ElseIf Not Mid(Me.ChaveNFE, 3, 2) = String.Format("{0:yy}", CDate(Me.DataNota)) Then
                erroMsg = "Ano da Nota na Chave Eletrônica diferente do informado na Nota Fiscal."
                Return False
                'ElseIf Not Mid(Me.ChaveNFE, 5, 2) = String.Format("{0:MM}", CDate(txtDataDeEmissao.Text)) Then
                '    erroMsg = "Més da Nota na Chave Eletrénica diferente do informado na Nota Fiscal."
                '    Return False
            End If

        End If

        'Verifica obrigação da Nota de Produtor (NUTRI) - 27/08/2020 - FURLAN
        If Me.Empresa.Empresa.ObrigaNfProdutor AndAlso Me.ObrigaNFProdutor AndAlso Not Me.CodigoFinalidade = 5 Then

            If Me.NotaTrocaOrigem Is Nothing Then
                erroMsg = "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente."
                Return False
            ElseIf Me.NotaTrocaOrigem.Itens.Count = 0 Then
                erroMsg = "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente."
                Return False
            ElseIf Me.NotaTrocaOrigem.Itens(0).QuantidadeFiscal = 0 AndAlso Me.NotaTrocaOrigem.Itens(0).ValorTotal = 0 Then
                erroMsg = "Nota Fiscal do Produtor não foi selecionada, faça a busca novamente."
                Return False
            ElseIf String.IsNullOrEmpty(Me.Cliente.InscricaoEstadual) OrElse Me.Cliente.InscricaoEstadual.Contains("ISENTO") Then
                erroMsg = "Inscrição Estadual do Produtor não foi informada, verifique no Cadastro de Clientes."
                Return False
            End If
        End If

        'Obrigatorio para a Baxi a informação de embarque para saida
        If Me.EntradaSaida = eEntradaSaida.Saida AndAlso Me.CodigoLocalEmbarque.ToString.Length = 0 AndAlso Me.NossaEmissao AndAlso (Mid(Me.CodigoEmpresa, 1, 8) = "40938762" Or Mid(Me.CodigoEmpresa, 1, 8) = "49673784") Then
            erroMsg = "Local de Embarque não foi selecionado"
            Return False
        End If

        'furlan - não deixar gravar encargo caso não tenha conta débito ou crédito - 18/07/2014
        '*********************************************************************************************************************************************************
        If Me.SubOperacao.Contabil AndAlso Not Left(Me.CodigoEmpresa, 8) = "04854422" AndAlso Not Left(Me.CodigoEmpresa, 8) = "03189063" Then
            For Each item In Me.Itens
                If String.IsNullOrWhiteSpace(item.SubOperacao.CodigoGrupoContas) Then
                    erroMsg = "Operação " & Me.CodigoOperacao & "-" & Me.CodigoSubOperacao & " não possui conta contábil, verifique."
                    Return False
                End If

                For Each enc In item.Encargos
                    If Not enc.Codigo = "LIQUIDO" AndAlso enc.Valor > 0 Then
                        If String.IsNullOrWhiteSpace(enc.OperacaoEncargo.CodigoDebitaConta) AndAlso String.IsNullOrWhiteSpace(enc.OperacaoEncargo.CodigoCreditaConta) AndAlso enc.OperacaoEncargo.Sinal <> "=" Then
                            erroMsg = "Produto " & item.CodigoProduto & " ENCARGO " & enc.Codigo & " não possui conta DéBITO/CRéDITO, verifique."
                            Return False
                        End If
                    End If
                Next
            Next
        End If
        'ATé AQUI

        If Me.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES AndAlso Me.Empresa.Empresa.RegistroDeExportacao Then
            If Me.DadosDaExportacao Is Nothing Then
                erroMsg = "A Exportacao Exige um Registro de Exportacao, para Emissao da Nota."
                Return False
            End If
        End If

        'Comentado devido ao FinanceiroNovo que seré retirado do processo
        'If Me.FinanceiroNovo AndAlso Me.SubOperacao.Financeiro AndAlso Not Me.Pedido.SubOperacao.Classe.Equals(eClassesOperacoes.AFIXAR) _
        '   AndAlso Not (Not Me.Pedido.FinanceiroNovo AndAlso Me.Pedido.MomentoFinanceiro = eTipoFaturamento.Pedido) Then
        '    If (Me.Titulos.Count + Me.Titulos.AdiantamentosAbertos.Count = 0) Then
        '        erroMsg = "A suboperação possui financeiro, porém não hé programação financeira para a nota fiscal!"
        '        Return False
        '    End If

        '    If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
        '        If (Me.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa) + Me.Titulos.Sum(Function(T) T.Valores.EncargoValorDocumento.Valor)) <> Me.TotalNota Then
        '            erroMsg = "A suboperação possui financeiro, porém o valor financeiro programado não corresponde com o valor da nota fiscal!"
        '            Return False
        '        End If
        '    Else
        '        If (Me.Titulos.AdiantamentosAbertos.Sum(Function(T) T.VlrBaixa) + Me.Titulos.Sum(Function(T) T.Valores.EncargoValorDocumento.Valor)) <> Math.Round(Me.TotalNota / Me.IndiceNota, 2, MidpointRounding.AwayFromZero) Then
        '            erroMsg = "A suboperação possui financeiro, porém o valor financeiro programado não corresponde com o valor da nota fiscal!"
        '            Return False
        '        End If
        '    End If
        'End If

        If Not Me.SubOperacao.Devolucao AndAlso
           Me.SubOperacao.Financeiro AndAlso
           Me.Pedido.MomentoFinanceiro = 3 Then

            If Left(Me.Empresa.Codigo, 8) = "24450490" AndAlso Not Me.SubOperacao.QuantidadeFiscal AndAlso Not Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                'NÃO FAZ NADA, LIBERADO NOTA FISCAL PARA COMPLEMENTO DE PEDIDO EM DÓLAR PARA RTGRÃOS - FURLAN - 23/10/2024
            ElseIf Me.VencimentosNota Is Nothing Then
                erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
                Return False
            ElseIf Me.VencimentosNota.Count = 0 Then
                erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
                Return False
            ElseIf Not Me.VencimentosNota.Sum(Function(s) s.ValorDoDocumento) = Me.TotalNota Then
                erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
                Return False
            End If
        End If

        If Not Me.VencimentosNota Is Nothing AndAlso Me.VencimentosNota.Count > 0 Then
            If Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then

                If Me.VencimentosNota.Any(Function(s) s.Prorrogacao = Date.Today) Then
                    erroMsg = "Data de Vencimento não pode ser igual a data de hoje!"
                    Return False
                End If

                If Me.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR AndAlso
                    Me.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then

                    If Not Me.VencimentosNota.Where(Function(s) s.CodigoProvisao = 2).Sum(Function(s) s.ValorDoDocumento) = Me.Pedido.Itens(0).Fixacoes.Where(Function(s) s.Codigo = Me.Itens(0).CodigoFixacao).FirstOrDefault.TotalOficial Then
                        erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
                        Return False
                    End If

                ElseIf Not Me.VencimentosNota.Where(Function(s) s.CodigoProvisao = 2).Sum(Function(s) s.ValorDoDocumento) = Me.TotalNota Then
                    erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
                    Return False
                End If
            Else
                If (Left(Me.CodigoEmpresa, 8) = "24450490" OrElse Left(Me.CodigoEmpresa, 8) = "44979506") AndAlso Not Me.SubOperacao.QuantidadeFiscal AndAlso Me.SubOperacao.FinalidadeDaNota = 2 Then
                    ''LIBERA PARA COMPLEMENTO RTGRÃO - FURLAN - 28/01/2025
                    ''LIBERA PARA COMPLEMENTO VERDE  - FURLAN - 13/05/2025
                ElseIf Math.Abs(Math.Round(Me.VencimentosNota.Where(Function(s) s.CodigoProvisao = 2).Sum(Function(s) s.MoedaValorDoDocumento), 2, MidpointRounding.AwayFromZero) - Math.Round(Me.TotalNotaMoeda, 2, MidpointRounding.AwayFromZero)) > 1 Then
                    erroMsg = "Verifique o valor financeiro programado, não corresponde com o valor Total da Nota Fiscal!"
                    Return False
                End If
            End If

            Dim dataParcela = Today()
            Dim primeiro = True

            If Left(Me.Empresa.Codigo, 8) = "24450490" AndAlso Not Me.SubOperacao.QuantidadeFiscal AndAlso Not Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial Then
                'NÃO FAZ NADA, LIBERADO PARA COMPLEMENTO DE PEDIDO EM DÓLAR PARA RTGRÃOS - FURLAN - 23/10/2024
            Else
                For Each parcela In Me.VencimentosNota
                    If primeiro AndAlso Me.VencimentosNota.Count > 1 AndAlso parcela.Prorrogacao >= dataParcela Then
                        dataParcela = parcela.Prorrogacao
                        primeiro = False
                    End If

                    If Me.Empresa.Empresa.ControlaDataMovimentoNFG AndAlso parcela.Prorrogacao < dataParcela Then
                        erroMsg = "Verifique o(s) vencimento(s) no financeiro programado pois a ordem não está correspondente!"
                        Return False
                    ElseIf parcela.ValorDoDocumento = 0 Then
                        erroMsg = "Verifique o(s) vencimento(s), valor programado no tétulo " & parcela.Codigo & " não pode ser 0(ZERO)."
                        Return False
                    End If

                    dataParcela = parcela.Prorrogacao
                Next
            End If

        End If

        If Not Me.NotaDeTerceiro AndAlso Not Me.VencimentosPedido Is Nothing Then

            Dim saldoEmProvisao As Decimal
            For Each tit As [Lib].Negocio.Titulo In Me.VencimentosPedido
                If tit.CodigoProvisao = 3 Then
                    saldoEmProvisao += tit.ValorLiquido
                End If
            Next

        End If

        If Me.EntradaSaida = eEntradaSaida.Saida And Me.SubOperacao.QuantidadeFisico AndAlso Not Me.Romaneio Is Nothing Then
            Dim qtde As Decimal
            For Each item As [Lib].Negocio.NotaFiscalXItem In Me.Itens
                qtde += item.QuantidadeFiscal
            Next

            If Me.Itens.Count = 1 Then
                If Me.Itens(0).Produto.Unidade = "KG" OrElse Me.Itens(0).Produto.Unidade = "KGS" OrElse Me.Itens(0).Produto.Unidade = "TON" Then
                    If Me.Romaneio.PesoLiquido <> qtde Then
                        erroMsg = "Peso Fisico e Fiscal não pode ser Diferente."
                        Return False
                    End If
                End If
            End If
        End If

        If Me.EntradaSaida = eEntradaSaida.Entrada And Me.SubOperacao.QuantidadeFisico AndAlso Not Me.Romaneio Is Nothing Then
            Dim qtde As Decimal
            For Each item As [Lib].Negocio.NotaFiscalXItem In Me.Itens
                qtde += item.QuantidadeFisica
            Next
            If Me.Romaneio.PesoLiquido <> qtde Then
                erroMsg = "Peso Físico da Nota não pode ser diferente do Romaneio, limpe e rafaéa o processo."
                Return False
            End If
        End If

        If Not Me.NotaDeTerceiro AndAlso String.IsNullOrWhiteSpace(Me.CodigoDeposito) Then
            erroMsg = "Selecione o Deposito"
            Return False
        End If

        If Not Me.NotaDeTerceiro AndAlso String.IsNullOrWhiteSpace(Me.CodigoDestino) Then
            erroMsg = "Selecione o Deposito de origem/destino"
            Return False
        End If

        If Not Left(Me.CodigoEmpresa, 8) = "24450490" AndAlso Not [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.IMPORTACOES AndAlso Me.Cliente.CodigoEstado = "EX" AndAlso Me.CodigoLocalEmbarque.ToString.Length = 0 Then
            erroMsg = "Local de Embarque não foi selecionado"
            Return False
        End If

        For Each row As [Lib].Negocio.NotaFiscalXItem In Me.Itens
            If row.Produto.ControlarNumeroDoLote AndAlso row.SubOperacao.ControlarNumeroDoLote OrElse (row.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso row.SubOperacao.ControlarNumeroDoLote AndAlso (row.Produto.CodigoGrupo = "10101" Or row.Produto.CodigoGrupo = "10102" Or row.Produto.CodigoGrupo = "30101" Or row.Produto.CodigoGrupo = "30102")) Then

                If Me.SubOperacao.EstoqueFisico Then
                    If Me.Itens(0).Produto.Unidade = "KG" OrElse Me.Itens(0).Produto.Unidade = "KGS" OrElse Me.Itens(0).Produto.Unidade = "TON" Then
                        'não FAZ NADA - LIBERADO PARA BAXI - FURLAN - 08/04/2024
                    ElseIf row.QuantidadeFisica <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
                        erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
                        Return False
                    End If
                Else
                    If row.QuantidadeFiscal <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
                        erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
                        Return False
                    End If
                End If

                For Each nLote In row.Lotes
                    If nLote.Validade < Me.Movimento Then
                        erroMsg = "Validade informada no Lote não pode ser menor que a data da Nota Fiscal, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome & " lote " & nLote.Lote
                        Return False
                    End If
                Next

            End If

            If row.Produto.ControlarNumeroDoLote AndAlso row.SubOperacao.ControlarNumeroDoLote OrElse (row.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso row.SubOperacao.Devolucao AndAlso row.SubOperacao.ControlarNumeroDoLote AndAlso (row.Produto.CodigoGrupo = "10101" Or row.Produto.CodigoGrupo = "10102" Or row.Produto.CodigoGrupo = "30101" Or row.Produto.CodigoGrupo = "30102")) Then

                If Me.SubOperacao.EstoqueFisico Then
                    If Me.Itens(0).Produto.Unidade = "KG" OrElse Me.Itens(0).Produto.Unidade = "KGS" OrElse Me.Itens(0).Produto.Unidade = "TON" Then
                        'não FAZ NADA - LIBERADO PARA BAXI - FURLAN - 08/04/2024
                    ElseIf row.QuantidadeFisica <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
                        erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
                        Return False
                    End If
                Else
                    If row.QuantidadeFiscal <> row.Lotes.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.Quantidade) Then
                        erroMsg = "Quantidade do Lote informada diferente da quantidade do Produto, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome
                        Return False
                    End If
                End If

                For Each nLote In row.Lotes
                    If nLote.Validade < Me.Movimento Then
                        erroMsg = "Validade informada no Lote não pode ser menor que a data da Nota Fiscal, verifique o item " & row.CodigoProduto & "-" & row.Produto.Nome & " lote " & nLote.Lote
                        Return False
                    End If
                Next
            End If

            If row.QuantidadeFiscal > row.SaldoPedidoFiscal And (row.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS Or row.SubOperacao.Classe = eClassesOperacoes.AFIXAR) And row.SubOperacao.Devolucao = True Then
                erroMsg = "A Quantidade informada não pode ser maior que o saldo do Pedido para Devolução"
                Return False
            End If

            If row.Produto.ControlarLote And row.Classificacao = "" And Me.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
                erroMsg = row.Produto.Nome & ", Este Produto é Controlado por Lote, informe o lote ao qual o produto pertence."
                Return False
            End If

            If row.Produto.ControlarEmbalagem And row.CodigoEmbalagem = 0 And Me.SubOperacao.Classe <> eClassesOperacoes.GLOBAL And (Me.SubOperacao.QuantidadeFiscal Or Me.SubOperacao.QuantidadeFisico) Then
                erroMsg = row.Produto.Nome & ", Este Produto é Controlado por Embalagem, informe a Embalagem o qual o produto é Comercializado."
                Return False
            End If

            If row.Produto.ControlarPecas AndAlso row.SubOperacao.ControlarPecas AndAlso row.NumeroPecas = 0 Then
                erroMsg = row.Produto.Nome & "não foi informado a quantidade de Peéas/Meios do Produto."
                Return False
            End If

            'USADO PARA LIBERAR OPERAééES DE IMPORTAção DA RT GRãoS - FURLAN - 26/03/2024
            Dim validaIMP As Boolean = True
            If Left(Me.CodigoEmpresa, 8) = "24450490" AndAlso Me.SubOperacao.EstoqueFisico = False Then
                validaIMP = False
            End If

            If validaIMP And row.CFOP > 3000 AndAlso row.CFOP < 4000 And bTabImportacao Then
                If Me.DadosDaImportacao Is Nothing Then
                    erroMsg = "Declaração de Importação não foi Informada."
                    Return False
                ElseIf Trim(Me.DadosDaImportacao.NumeroDeclaracaoImportacao).Length < 11 Then
                    erroMsg = "Declaração de Importação do Produto " & row.CodigoProduto & " deve ter 11 dégitos."
                    Return False
                ElseIf Trim(Me.DadosDaImportacao.LocalDesembarqueImportacao).Length = 0 Then
                    erroMsg = "Local do Desembarque da Declaração de Importação do Produto " & row.CodigoProduto & " deve ser informado."
                    Return False
                ElseIf Me.DadosDaImportacao.EstadoDesembarqueImportacao.Length = 0 Then
                    erroMsg = "Estado do Desembarque da Declaração de Importação do Produto " & row.CodigoProduto & " deve ser informado."
                    Return False
                ElseIf Trim(Me.DadosDaImportacao.CodigoFabricante).Length = 0 Then
                    erroMsg = "Fabricante que está na Declaração de Importação do Produto " & row.CodigoProduto & " deve ser informado."
                    Return False
                End If
            End If
        Next

        If Not Me.SubOperacao.Devolucao AndAlso
            Me.EntradaSaida = eEntradaSaida.Saida AndAlso
            Me.Itens.Where(Function(s) s.IUD <> "D").Sum(Function(s) s.QuantidadeFiscal) = 0 AndAlso
            (Me.NotasReferenciais Is Nothing OrElse Me.NotasReferenciais.Count = 0) Then

            'TRATANDO PARA NOTA DE TRANSFERéNCIA DE ICMS - FURLAN - 28/11/2023
            If Not Left(Me.CodigoEmpresa, 8) = "24450490" AndAlso Not Me.SubOperacao.Descricao.Contains("TRANSFERENCIA DE CREDITO DE ICMS") Then
                If Me.CFOP.Codigo = 5602 OrElse Me.CFOP.Codigo = 5605 Then
                    'LIBERA TRANSF. SDO CREDOR DE ICMS P/OUTRO ESTAB.DA EMPR. - FURLAN 03/04/2025
                Else
                    erroMsg = "Nota fiscal Referencial não foi selecionada."
                    Return False
                End If
            End If
        End If

        If Not Left(Me.CodigoEmpresa, 8) = "24450490" AndAlso Me.CodigoFinalidade = 28 AndAlso (Me.NotasReferenciais Is Nothing OrElse Me.NotasReferenciais.Count = 0) Then
            If Me.CFOP.Codigo = 5602 OrElse Me.CFOP.Codigo = 5605 Then
                'LIBERA TRANSF. SDO CREDOR DE ICMS P/OUTRO ESTAB.DA EMPR. - FURLAN 03/04/2025
            Else
                erroMsg = "Nota fiscal Referencial não foi selecionada."
                Return False
            End If
        End If

        'NOTA FISCAL COMPLEMENTAR TEM QUE TER UMA REFERÊNCIA - FURLAN - 06/08/2024
        If Not Me.SubOperacao.Devolucao AndAlso Me.Eletronica AndAlso Me.NossaEmissao AndAlso Not Me.SubOperacao.FinalidadeDaNota = 3 AndAlso Me.Itens(0).QuantidadeFiscal = 0 AndAlso (Me.NotasReferenciais Is Nothing OrElse Me.NotasReferenciais.Count = 0) Then

            If Me.CFOP.Codigo = 5602 OrElse Me.CFOP.Codigo = 5605 Then
                'LIBERA TRANSF. SDO CREDOR DE ICMS P/OUTRO ESTAB.DA EMPR. - FURLAN 03/04/2025
            Else
                erroMsg = "Nota fiscal Referencial não foi selecionada."
                Return False
            End If
        End If

        If Me.Itens(0).CodigoFixacao = 0 AndAlso
           Me.Itens(0).Produto.ControlarPesagem AndAlso
           Me.CodigoRomaneio = 0 AndAlso
           Me.SubOperacao.QuantidadeFisico AndAlso
           Me.SubOperacao.Classe <> eClassesOperacoes.GLOBAL Then
            erroMsg = "é Obrigatorio a seleção de um Romaneio ou fazer a classificaéao do produto, Lembrando que a classificação pela nota gera um Romaneio Fisico = não."
            Return False
        End If

        'If Not Session("TotalProcuracao" & HID.Value) Is Nothing Then
        If Me.ProcuracaoSaldoPedido > 0 And Me.ProcuracaoSaldoPedido < Me.Itens(0).SaldoValorOficial Then
            erroMsg = "Verifique o saldo referente as procuraéées. não é permitido valores negativos."
            Return False
        End If
        'End If

        If Me.CodigoRomaneio > 0 And
                   Me.CriarRomaneio = False And
                   (Me.CodigoFinalidade = 20 OrElse Me.CodigoFinalidade = 22) And
                   Me.EntradaSaida.ToString.Substring(0, 1) = "S" Then
            erroMsg = "Devolução fisica não pode ser usada com a finalidade de compra cédigo 20 ou 22."
            Return False
        End If

        If Me.CodigoRomaneio > 0 And
                   Me.CriarRomaneio = True And
                   Me.SubOperacao.Devolucao = True And
                   (Me.CodigoFinalidade <> 20 AndAlso Me.CodigoFinalidade <> 22) And
                   Me.EntradaSaida.ToString.Substring(0, 1) = "S" And
                   (Me.CodigoAutorizacao = 0) And Not Me.TemNotaTroca Then
            erroMsg = "Devolução fisico = não, sé pode ser feita com a finalidade cédigo 20 ou 22."
            Return False
        End If

        If Me.NossaEmissao AndAlso Not Me.Transportador Is Nothing AndAlso Me.Transportador.Codigo.ToString.Length > 0 AndAlso Me.PlacaTransportador.ToString.Length = 0 Then
            erroMsg = "Placa não foi selecionada."
            Return False
        End If

        If Me.NossaEmissao AndAlso Not Me.CodigoTransportador Is Nothing AndAlso Me.CodigoTransportador.Length > 0 Then
            If Me.PlacaDetalhes Is Nothing Then
                erroMsg = "Placa não foi selecionada."
                Return False
            ElseIf Me.PlacaDetalhes.Placa01.Length = 0 Then
                erroMsg = "Placa não foi selecionada."
                Return False
            End If
        End If

        If AutorizacaoENecessaria(Me) AndAlso Not Me.CodigoAutorizacao > 0 Then
            erroMsg = "Vocé deve selecionar uma autorizacao de retirada, caso não exista crie uma."
            Return False
        End If

        If Me.CodigoAutorizacao > 0 Then
            If Me.SubOperacao.QuantidadeFiscal AndAlso Me.Autorizacao.SaldoFiscal < IIf(Me.Itens(0).PesoQuantidade = "P", Me.Itens(0).PesoFiscal, Me.Itens(0).QuantidadeFiscal) Then
                erroMsg = "Saldo Fiscal Insuficiente para geração da Nota, Saldo:" & Me.Autorizacao.SaldoFiscal & " Nota Fiscal: " & Me.Itens(0).PesoFiscal
                Return False
            End If
            If Me.SubOperacao.QuantidadeFisico AndAlso Me.Autorizacao.SaldoFisico < Me.Itens(0).PesoLiquido Then
                erroMsg = "Saldo Fésico Insuficiente para geração da Nota, Saldo:" & Me.Autorizacao.SaldoFisico & " Nota Fiscal: " & Me.Itens(0).PesoLiquido
                Return False
            End If
        End If

        'If CDate(txtDataDeEmissao.Text) > Now.Date Then
        If DateValue(CDate(Me.DataNota)) > DateValue(Now.Date) Then
            erroMsg = "A Data de Emissão não pode ser maior que a Data Atual."
            Return False
        End If

        'If CDate(txtDataDeEntrada.Text) < CDate(txtDataDeEmissao.Text) Then
        If DateValue(CDate(Me.Movimento)) < DateValue(CDate(Me.DataNota)) Then
            erroMsg = "A Data de Emissão não pode ser posterior a Data do Movimento."
            Return False
        End If

        If String.IsNullOrWhiteSpace(Me.Codigo) Then
            erroMsg = "Némero da nota não foi informado."
            Return False
        End If

        If CInt(Me.Codigo) = 0 Then
            erroMsg = "Némero da Nota Fiscal não pode ser 0."
            Return False
        End If

        If String.IsNullOrWhiteSpace(Me.Serie) Then
            erroMsg = "Serie não foi informada."
            Return False
        End If

        If Not Me.NotaDeTerceiro AndAlso Me.CodigoFinalidade = 0 Then
            erroMsg = "Finalidade não foi selecionada."
            Return False
        End If

        If Me.Eletronica AndAlso Not Me.NossaEmissao AndAlso Me.EntradaSaida.ToString.Substring(0, 1) = "E" AndAlso Me.ChaveNFE.Replace(".", "").Length <> 44 Then
            erroMsg = "Informe a chave da nota eletrénica, deve ter 44 dégitos."
            Return False
        End If

        If Me.Itens.Count = 0 Then
            erroMsg = "Verifique itens da Nota Fiscal."
            Return False
        End If

        'If DgEncargos.Rows.Count = 0 Then
        '    erroMsg = "Verifique encargos da Nota Fiscal."
        '    Return False
        'End If

        If dsXML Is Nothing AndAlso Not Me.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE ICMS") AndAlso
            Not Me.SubOperacao.Descricao.Contains("COMPLEMENTO DE PESO") AndAlso
            Not Me.SubOperacao.Descricao.Contains("COMPLEMENTO DE ICMS-ST") AndAlso
            Not Me.SubOperacao.Descricao.Contains("TRANSFERENCIA DE CREDITO DE ICMS") AndAlso
            Not Me.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE PRECO A FIXAR") AndAlso
            Not Me.SubOperacao.Descricao.Contains("COMPLEMENTACAO DE IPI") Then
            If Me.TotalNota <= 0 Then
                erroMsg = "Valor da Nota Fiscal não pode ser zero."
                Return False
            End If

            'ADICIONADO EM 02-04-2013 REMOVER EM 2014
            For Each row In Me.Itens
                If row.ValorTotal = 0 Then
                    erroMsg = row.Produto.Nome & " - Valor da Nota Fiscal não pode ser zero. ENTRE EM CONTATO COM A TI."
                    Return False
                    If row.Encargos.EncProduto.Valor = 0 Then
                        erroMsg = row.Produto.Nome & " - Encargo - Valor do Encargo da Nota Fiscal não pode ser zero. ENTRE EM CONTATO COM A TI."
                        Return False
                    End If
                End If
            Next
        End If

        For Each prd As [Lib].Negocio.NotaFiscalXItem In Me.Itens
            If prd.Encargos.EncProduto.ValorPeso = eValorPeso.Valor Then
                If prd.ValorTotal <> prd.Encargos.EncProduto.Base Then
                    erroMsg = "O Valor do produto " & prd.CodigoProduto & " - " & prd.Produto.Nome & " não confere com a Base De Calculo dos encargos. Cliquem em OK para atualizar os Totais do Produto e se o problema persisitr informe o TI."
                    Return False
                End If
            Else
                If prd.QuantidadeFiscal <> prd.Encargos.EncProduto.Base Then
                    erroMsg = "A Quantidade do produto " & prd.CodigoProduto & " - " & prd.Produto.Nome & " não confere com a Base De Calculo dos encargos, Cliquem em OK para atualizar os Totais do Produto e se o problema persisitr informe o TI."
                    Return False
                End If
            End If

            For Each enc In prd.Encargos
                If enc.Codigo = "ICMS" AndAlso enc.SituacaoTributaria = 0 AndAlso enc.Valor = 0 Then
                    erroMsg = "Valor do Icms Tributado Integralmente não foi encontrado no Produto " & prd.CodigoProduto & "-" & prd.Produto.Nome & "."
                    Return False
                End If
            Next

            If prd.Encargos.EncProduto.SituacaoTributariaIPI = 0 Or prd.Encargos.EncProduto.SituacaoTributariaIPI = 50 Then
                If Not prd.Encargos.Where(Function(s) s.Codigo.Contains("IPI")).Count > 0 Then
                    erroMsg = "Produto " & prd.CodigoProduto & "-" & prd.Produto.Nome & " com SituacaoTributariaIPI = " & prd.Encargos.EncProduto.SituacaoTributariaIPI & " não existe o Encargo IPI parametrizado. Verifique parametrização na OperaçãoXEncargos."
                    Return False
                ElseIf prd.Encargos.Where(Function(s) s.Codigo.Contains("IPI")).Sum(Function(s) s.Valor = 0) Then
                    erroMsg = "Produto " & prd.CodigoProduto & "-" & prd.Produto.Nome & " com SituacaoTributariaIPI = " & prd.Encargos.EncProduto.SituacaoTributariaIPI & " não pode ser Zero(0)."
                    Return False
                End If
            End If
        Next

        'If FinanceiroNovo AndAlso Me.Pedido.CondicaoPagamento IsNot Nothing AndAlso Me.Pedido.CondicaoPagamento.Antecipado AndAlso Me.TotalNota > Me.Titulos.AdiantamentosAbertos.ValorTotalInformadoParaBaixa Then
        '    erroMsg = "Para Pagamento/Recebimento Antecipado é necessério ter o valor: " & Me.TotalNota.ToString("N2") & " em Adiantamento"
        '    Return False
        'End If

        'If FinanceiroNovo AndAlso Not Me.SubOperacao.Devolucao AndAlso
        '    Me.SubOperacao.Financeiro AndAlso
        '    Me.TotalNota > Me.Pedido.Titulos.Where(Function(s) s.CodigoProvisao = eProvisao.Previsao).Sum(Function(s) s.Valores.EncargoValorLiquido.ValorOficial) + Me.Titulos.AdiantamentosAbertos.ValorTotalInformadoParaBaixa AndAlso
        '    Not Me.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR Then
        '    erroMsg = "Valor Total da Nota Fiscal não é compatével com o financeiro do pedido"
        '    Return False
        'End If

        'Obrigatéria informação do némero do processo de drawback para CFOP: 
        ' 7127: Venda de produção do estabelecimento sob o regime de drawback 
        ' 7211: Devolução de compras p/ industrialização sob o regime de drawback
        If (Me.CFOP.Codigo = 7127 OrElse Me.CFOP.Codigo = 7211) AndAlso (Me.DadosDaExportacao Is Nothing OrElse String.IsNullOrWhiteSpace(Me.DadosDaExportacao.NumAtoConcessorio)) Then
            erroMsg = "Para exportação em regime de Drawback é obrigatério o preenchimento do campo: Num. Ato Concessério (Drawbak), na aba exportação"
            Return False
        End If

        'Para NF de importação não preencher o campo Local de Embarque/Coleta.
        If [Enum].Parse(GetType(eClassesOperacoes), Me.Operacao.CodigoClasse.ToUpper) = eClassesOperacoes.IMPORTACOES AndAlso Not String.IsNullOrWhiteSpace(Me.CodigoLocalEmbarque) Then
            erroMsg = "O campo Local de Embarque/Coleta não deve ser preenchido em Nota Fiscal de importação"
            Return False
        End If

        'Para NF de exportação, obrigatério JABER/ZéLIO - 22/10/2018
        If Left(Me.CodigoEmpresa, 8) = "03189063" _
            AndAlso Me.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES _
            AndAlso (Me.DadosDaExportacao Is Nothing OrElse Me.DadosDaExportacao.FaturaExportacao.Length = 0) _
            AndAlso Me.EntradaSaida = eEntradaSaida.Saida Then

            erroMsg = "Némero da Fatura de Exportação é obrigatério"
            Return False
        End If

        'Validação Xml de Notas de Entrada'
        'If SessaoDsXML IsNot Nothing Then
        'erroMsg = DocumentoEletronico.ValidaNFEntradaComXML(Me, SessaoDsXML)
        '    If Not String.IsNullOrWhiteSpace(erroMsg) Then
        '        Return False
        '    End If
        'End If

        'Validar Totals dos Produtos e Total da Nota Fiscal
        If dsXML IsNot Nothing AndAlso Not Me.NotaDeTerceiro Then

            If Me.TemSegCodBarra AndAlso Not Me.SegCodBarra.Length = 36 Then
                erroMsg = "O Segundo Cédigo de Barras não está preenchido corretamente."
                Return False
            End If

            Dim XmlValorTotalProduto As Decimal = CDec(dsXML.Tables("ICMSTot").Rows(0)("vProd").ToString().Replace(".", ","))
            Dim XmlValorTotalNota As Decimal = CDec(dsXML.Tables("ICMSTot").Rows(0)("vNF").ToString().Replace(".", ","))
            Dim validaTotaisXML As Boolean = True

            If Me.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR AndAlso
                Me.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
                validaTotaisXML = False
            End If

            Dim temDesconto As Boolean = False

            If CDec(dsXML.Tables("ICMSTot").Rows(0)("vDesc").ToString().Replace(".", ",")) > 0 Then temDesconto = True

            If validaTotaisXML AndAlso Not Me.TotalProduto = XmlValorTotalProduto Then
                erroMsg = "O valor total do(s) Produto(s) no XML " & XmlValorTotalProduto & " está diferente do valor da NF " & Me.TotalProduto.ToString()
                Return False
            End If

            If validaTotaisXML AndAlso Not Me.TotalNota = XmlValorTotalNota Then
                Dim totalGeral As Decimal = Me.TotalNota

                For Each prd In Me.Itens
                    For Each enc In prd.Encargos

                        If Not enc.Codigo = "PRODUTO" AndAlso enc.Sinal = "+" Then
                            totalGeral -= enc.Valor
                        ElseIf enc.Sinal = "-" Then
                            totalGeral += enc.Valor
                        End If

                        'If enc.Codigo = "FUNRURAL" OrElse enc.Codigo = "SENAR" OrElse enc.Codigo = "FETHAB" OrElse enc.Codigo = "IAGRO" Then
                        '    totalGeral += enc.Valor
                        'End If
                    Next
                Next

                If Not temDesconto AndAlso Not totalGeral = XmlValorTotalNota Then
                    erroMsg = "O valor total da NF no XML " & XmlValorTotalNota & " está diferente do valor da NF " & Me.TotalNota.ToString()
                    Return False
                End If
            End If
        End If

        erroMsg = ""
        Return True
    End Function

    Public Function AutorizacaoENecessaria(ByRef objNF As NotaFiscal) As Boolean

        If Left(objNF.CodigoEmpresa, 8) = "24450490" Then 'NÃO OBRIGAR AUTORIZAÇÃO DE RETIRADA PARA RTGRÃOS - FURLAN - 09/07/2024
            Return False
        ElseIf objNF IsNot Nothing AndAlso objNF.SubOperacao IsNot Nothing AndAlso
                 objNF.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso
                 (objNF.SubOperacao.QuantidadeFisico OrElse objNF.SubOperacao.QuantidadeFiscal) AndAlso
                 objNF.SubOperacao.Classe <> eClassesOperacoes.GLOBAL AndAlso
                 objNF.Itens IsNot Nothing AndAlso objNF.Itens.Count > 0 AndAlso
                 (objNF.Itens(0).Produto.Agrupar = "N" Or objNF.Itens(0).Produto.AutorizacaoDeRetirada) AndAlso
                 (objNF.CodigoFinalidade <> 14 AndAlso objNF.CodigoFinalidade <> 20 AndAlso objNF.CodigoFinalidade <> 22) Then
            Return True
        Else
            Return False
        End If

    End Function

    Public Sub HoraSefaz()

        If Me.Empresa.Empresa.NotaFiscalEletronica AndAlso Me.Empresa.Empresa.NossaEmissao Then
            Dim easternZoneBR As TimeZoneInfo = TimeZoneInfo.FindSystemTimeZoneById(BuscarHoraServidorSefaz())
            Dim horaSefaz As DateTime = TimeZoneInfo.ConvertTimeFromUtc(DateTime.UtcNow, easternZoneBR)

            Me.DataInclusao = horaSefaz
        Else
            Me.DataInclusao = Date.Now
        End If

        If Me.NossaEmissao AndAlso Me.Eletronica Then
            Me.CodigoSituacao = 4
        Else
            Me.CodigoSituacao = 1
        End If

    End Sub

    Private Function BuscarHoraServidorSefaz() As String

        If Me.Empresa.Empresa.NotaFiscalEletronica AndAlso Me.Empresa.Empresa.NossaEmissao Then
            Dim fileEmpresa As String = "C:\NGS\xmlCopy\" & Me.Empresa.Codigo & ".xml"

            If Not File.Exists(fileEmpresa) Then
                Throw New Exception("Arquivo de Configuracão da Empresa não foi encontrado.")
            End If

            Dim xmlConf As New XmlDocument
            xmlConf.Load(fileEmpresa)

            'xmlConf.GetElementsByTagName("dados").GetNode("dados").ChildNodes.GetNode("horarioServidor").InnerText

            Dim dadosNodes As XmlNodeList = xmlConf.GetElementsByTagName("dados")

            ' Verifica se a lista contém pelo menos um né
            If dadosNodes.Count > 0 Then
                ' Acessa o primeiro né "dados" na lista
                Dim dadosNode As XmlNode = dadosNodes(0)

                ' Obtém a lista de nés filhos do né "dados"
                Dim childNodes As XmlNodeList = dadosNode.ChildNodes

                ' Itera pela lista de nés filhos para encontrar o né "horarioServidor"
                For Each childNode As XmlNode In childNodes
                    If childNode.Name = "horarioServidor" Then
                        ' Retorna o valor InnerText do né "horarioServidor"
                        Return childNode.InnerText
                    End If
                Next
            End If

        End If

        Return Now.ToString("yyyy-MM-ddTHH:mm:ss")

    End Function

    Public Sub CarregarDadosNotaDeTerceiro(ByVal sUsuario As String)

        Dim codigoProdutoTerceiro As String = "903010007"
        Dim produtoTerceiro As New Produto(codigoProdutoTerceiro)

        If produtoTerceiro Is Nothing OrElse String.IsNullOrEmpty(produtoTerceiro.Codigo) Then
            Throw New Exception(String.Format("O produto de terceiro {0} não está cadastrado!", codigoProdutoTerceiro))
            Exit Sub
        End If

        Dim op As New SubOperacao(70, 99)

        If op Is Nothing OrElse String.IsNullOrEmpty(op.Codigo) Then
            Throw New Exception(String.Format("A operação 70-99 PRESTACAO DE SERVICO - IMPORTAÇÃO NF TERCEIRO não está cadastrada!", codigoProdutoTerceiro))
            Exit Sub
        End If

        Dim Parametros As New OperacaoXEstado
        Parametros.Empresa = Left(Me.Empresa.Codigo, 8)
        Parametros.CodigoGrupoProduto = produtoTerceiro.CodigoGrupo
        Parametros.CodigoProduto = produtoTerceiro.Codigo
        Parametros.CodigoOperacao = op.CodigoOperacao
        Parametros.CodigoSubOperacao = op.Codigo
        Parametros.EstadoOrigem = Me.Empresa.CodigoEstado
        Parametros.EstadoDestino = Me.Cliente.CodigoEstado
        Parametros.InicioVigencia = Me.Movimento
        Dim OXE As New OperacaoXEstado(Parametros)

        If OXE.CodigoOperacao = 0 Then

            OXE = New [Lib].Negocio.OperacaoXEstado()
            OXE.ConsultarOperacaoNotaDeTerceiro(Parametros)
            OXE.IUD = "I"
            OXE.Ativo = True
            OXE.Empresa = Left(Me.Empresa.Codigo, 8)
            OXE.CodigoGrupoProduto = produtoTerceiro.CodigoGrupo
            OXE.CodigoProduto = produtoTerceiro.Codigo
            OXE.CodigoOperacao = op.CodigoOperacao
            OXE.CodigoSubOperacao = op.Codigo
            OXE.EstadoOrigem = Me.Empresa.CodigoEstado

            If Me.Empresa.CodigoEstado = Me.Cliente.CodigoEstado Then
                OXE.EstadoDestino = Me.Cliente.CodigoEstado
            Else
                OXE.EstadoDestino = Me.Cliente.Estado.Regiao
            End If

            OXE.InicioVigencia = Me.Movimento

            If UsuarioServidor.Usuario Is Nothing Then

                UsuarioServidor.CarregarInformacaoParaUsuarioServidor(sUsuario)

            End If

            Dim Banco As New AcessaBanco
            Dim Sqls As New ArrayList
            Dim Sql As String = ""

            OXE.SalvarSql(Sqls)

            If Not Banco.GravaBanco(Sqls) Then
                Throw New Exception("Não foi possível cadastrar a operação comercial!")
                Exit Sub
            Else
                Parametros.Empresa = Left(Me.Empresa.Codigo, 8)
                Parametros.CodigoGrupoProduto = produtoTerceiro.CodigoGrupo
                Parametros.CodigoProduto = produtoTerceiro.Codigo
                Parametros.CodigoOperacao = op.CodigoOperacao
                Parametros.CodigoSubOperacao = op.Codigo
                Parametros.EstadoOrigem = Me.Empresa.CodigoEstado
                Parametros.EstadoDestino = Me.Cliente.CodigoEstado
                Parametros.InicioVigencia = Me.Movimento
                OXE = New OperacaoXEstado(Parametros)
            End If

        End If

        Dim itemDaNota As [Lib].Negocio.NotaFiscalXItem
        itemDaNota = Me.Itens.FirstOrDefault()

        If Me.Itens.Count() = 0 Then
            Throw New Exception("A nota importada não tem item cadastrado!")
            Exit Sub
        End If

        Me.CodigoOperacao = op.CodigoOperacao
        Me.CodigoSubOperacao = op.Codigo

        With itemDaNota

            .IUD = "I"
            .CodigoProduto = produtoTerceiro.Codigo
            .PesoQuantidade = produtoTerceiro.PesoQuantidade
            .CodigoOperacao = op.CodigoOperacao
            .CodigoSubOperacao = op.Codigo
            .CodigoOperacaoEstado = OXE.Codigo

            itemDaNota.CarregandoEncargos = True
            .Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(itemDaNota, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})

            If Not .Encargos Is Nothing AndAlso .Encargos.Count() > 0 AndAlso Not .Encargos.EncProduto.GrupoDeContaDebito Is Nothing Then
                .TemCentroDeCusto = .Encargos.EncProduto.GrupoDeContaDebito.TemCentroDeCusto
                itemDaNota.CarregandoEncargos = False
            End If

            If itemDaNota.Encargos.Count() = 0 Then
                Throw New Exception("Encargos não cadastrado!")
                Exit Sub
            End If

        End With

        Dim objPedido As New Pedido(Me, 1)

        'NUMERADOR DOS PEDIDOS
        Dim SqlN As String = "exec sp_Numerador 'TERCEIROS'," & objPedido.EnderecoEmpresa & "," & [Lib].Negocio.eTiposNumerador.Pedido & ""
        Dim dsN As New DataSet
        Dim objBanco As New AcessaBanco

        dsN = objBanco.ConsultaDataSet(SqlN, "Numerador")

        Dim CodigoNumerador As Integer = 0
        If dsN IsNot Nothing AndAlso dsN.Tables(0).Rows.Count > 0 Then
            CodigoNumerador = dsN.Tables(0).Rows(0).Item(0)
        End If

        If Not CodigoNumerador > 0 Then
            Throw New Exception("Numerador de Pedidos não cadastrado!")
            Exit Sub
        End If

        objPedido.Codigo = CodigoNumerador
        Me.CodigoPedido = CodigoNumerador
        Me.Pedido = objPedido
        Me.CIFFOB = objPedido.FreteCIFFOB

        If Not objPedido.Salvar Then
            Throw New Exception("Não foi possível cadastrar o pedido!")
            Exit Sub
        End If

        For Each row In Me.VencimentosNota
            If row.CodigoProvisao = 1 Then
                Throw New Exception("Não pode ser alterado a Empresa com Financeiro Baixado - Título " & row.Codigo)
                Exit Sub
            Else
                If Not Me.SubOperacao.Financeiro Then row.IUD = "D"

                row.CodigoUnidadeDeNegocio = Me.CodigoUnidadeDeNegocio
                row.CodigoEmpresa = Me.CodigoEmpresa
                row.EnderecoEmpresa = Me.EnderecoEmpresa

                row.CodigoEmpresaPedido = Me.CodigoEmpresa
                row.EndEmpresaPedido = Me.EnderecoEmpresa
                row.CodigoPedido = Me.CodigoPedido

                row.CodigoCliente = Me.CodigoCliente
                row.EndCliente = Me.EnderecoCliente

                row.CodigoDestinatario = Me.CodigoCliente
                row.EndDestinatario = Me.EnderecoCliente
                row.NomeDoDestinatario = ""

                row.Movimento = Me.Movimento

                row.Historico = "PAGTO " & Trim(Me.TipoDeDocumento.Historico) & " " & Trim(Me.Codigo) & "-" & Trim(Me.Serie) & " - " & Trim(Me.Cliente.Nome) & " - " & Trim(row.Carteira.Descricao)
                row.Observacoes = Me.Observacoes
            End If
        Next

    End Sub

    Public Sub CarregarDetalhesPedido()

        Dim listPedidosSaldo As New ListSaldoPedido2015(ChecarParametros)
        Dim pedidoSaldo As New SaldoPedido2015
        'gridGlobalDireta.DataSource = objListSaldoPedidos.Where(Function(s) s.Tipo = 1)
        'gridGlobalDireta.DataBind()

        If listPedidosSaldo Is Nothing OrElse listPedidosSaldo.Count = 0 Then
            Throw New Exception(String.Format("Não foi encontrado nenhum pedido para o número do pedido MIC: {0}!", Me.CodigoPedidoMIC))
        End If

        If listPedidosSaldo.Count > 1 Then
            Throw New Exception(String.Format("Foi encontrado varios pedidos para o pedido MIC: {0}! Necessário verificação.", Me.CodigoPedidoMIC))
        End If

        pedidoSaldo = listPedidosSaldo.FirstOrDefault()

        If pedidoSaldo Is Nothing Then
            Throw New Exception(String.Format("Não foi encontrado saldo para o pedido: {0} ", Me.CodigoPedidoMIC))
        End If

        Me.Pedido = New Pedido(pedidoSaldo.Empresa.Codigo, pedidoSaldo.Empresa.CodigoEndereco, pedidoSaldo.CodigoPedido, Me.NotaDeTerceiro)
        'Me.CodigoPedido = Me.Pedido.Codigo

        Dim objCliente As New Cliente(Me.CodigoCliente, Me.EnderecoCliente)
        If objCliente.CodigoSituacao = 50 And pedidoSaldo.SubOperacao.Classe = eClassesOperacoes.VENDAS Then
            Throw New Exception("Cliente com Documentação pendente, Faturamento Bloqueado.")
        End If

        If Not pedidoSaldo Is Nothing AndAlso pedidoSaldo.Pedido.Troca Then
            If pedidoSaldo.Pedido.PedidoTroca.PedidoTroca Is Nothing Then
                Throw New Exception("Este pedido está vinculado ao pedido " & pedidoSaldo.Pedido.PedidoTroca.Codigo & " e esse não está marcado como troca.")
            End If
        End If

        If Not pedidoSaldo Is Nothing _
        AndAlso pedidoSaldo.Pedido.Troca _
        AndAlso pedidoSaldo.Pedido.CodigoPedidoTroca = 0 Then
            Throw New Exception("Este pedido está marcado como troca porém não está vinculado a outro pedido. Caso este não seja mais de troca então desmarque a opção troca.")
        End If

        If Not pedidoSaldo Is Nothing _
      AndAlso pedidoSaldo.Pedido.Troca _
      AndAlso pedidoSaldo.Pedido.CodigoPedidoTroca > 0 _
      AndAlso ((pedidoSaldo.Pedido.CodigoMoeda = eTiposMoeda.Oficial AndAlso Math.Abs(pedidoSaldo.Pedido.Itens.LiquidoOficial - pedidoSaldo.Pedido.PedidoTroca.Itens.LiquidoOficial) > 0.01) _
                OrElse (pedidoSaldo.Pedido.CodigoMoeda = eTiposMoeda.MoedaEstrangeira AndAlso Math.Abs(pedidoSaldo.Pedido.Itens.LiquidoOficial - pedidoSaldo.Pedido.PedidoTroca.Itens.LiquidoOficial) > 0.1)) Then

            Throw New Exception(String.Format("Quando o pedido é de Troca o Valor Liquido total dos itens dos pedidos devem ser iguais. Pedido: {0} Valor = {1} e Pedido de Troca: {2} Valor = {3}",
                          pedidoSaldo.Pedido.Codigo, pedidoSaldo.Pedido.Itens.LiquidoOficial, pedidoSaldo.Pedido.PedidoTroca.Codigo, pedidoSaldo.Pedido.PedidoTroca.Itens.LiquidoOficial))

            Exit Sub
        End If

        If Me.NossaEmissao Then
            Dim Mensagem As String = Me.NotaFiscalJaLancada()
            If Not String.IsNullOrWhiteSpace(Mensagem) Then
                Throw New Exception(Mensagem)
            End If
        End If

        Dim erroMsg As String = String.Empty
        Dim MsgAlerta As String = String.Empty
        Dim obj As [Lib].Negocio.IBaseEntity

        If pedidoSaldo.Itens Is Nothing OrElse pedidoSaldo.Itens.Count = 0 Then
            Throw New Exception("Sem item no pedido!")
        End If

        'Selecionamos o item
        pedidoSaldo.Itens(0).Selecionado = True

        If Not Me.Pedido.Depositos Is Nothing AndAlso Me.Pedido.Depositos.Count > 0 Then
            Me.CodigoDeposito = Me.Pedido.Depositos.FirstOrDefault().Deposito.Codigo
            Me.EnderecoDeposito = Me.Pedido.Depositos.FirstOrDefault().Deposito.CodigoEndereco
        End If

        Me.CodigoDestino = Me.CodigoDeposito
        Me.EnderecoDestino = Me.EnderecoDeposito

        Me.CodigoTransbordo = ""
        Me.EnderecoTransbordo = 0

        Me.CodigoLocalEmbarque = ""
        Me.EndLocalEmbarque = 0

        Me.CodigoEntrega = ""
        Me.EnderecoEntrega = 0

        Me.CodigoFinalidade = Me.Pedido.CodigoFinalidade

        Me.CodigoOperacao = pedidoSaldo.CodigoOperacao
        Me.CodigoSubOperacao = pedidoSaldo.CodigoSubOperacao
        'USADO PARA IMPORTAR NOTAS FISCAIS DE DEVOLUCÃO PROCESSO MIC CASO PRECISE - FURLAN 23/03/2026
        'Me.CodigoOperacao = 21
        'Me.CodigoSubOperacao = 4

        If CarregarItensComSaldo(pedidoSaldo, erroMsg, MsgAlerta) Then

            Dim itensAgrupadoPorProduto = From item In Me.Itens
                                          Group item By item.CodigoProduto Into Grupo = Group
                                          Select New With {
                                              .CodigoProduto = CodigoProduto,
                                              .SaldoPedidoFiscal = Grupo.Sum(Function(x) x.SaldoPedidoFiscal),
                                              .SaldoPedidoFisico = Grupo.Sum(Function(x) x.SaldoPedidoFisico),
                                              .SaldoValorMoeda = Grupo.Sum(Function(x) x.SaldoValorMoeda),
                                              .SaldoValorOficial = Grupo.Sum(Function(x) x.SaldoValorOficial),
                                              .Itens = Grupo.ToList()
                                          }


            For Each itemNF In Me.Itens

                itemNF.CodigoOperacao = Me.CodigoOperacao
                itemNF.CodigoSubOperacao = Me.CodigoSubOperacao

                itemNF.PesoFiscal = itemNF.QuantidadeFiscal
                itemNF.QuantidadeFisica = itemNF.QuantidadeFiscal

                For Each itemSaldo In itensAgrupadoPorProduto.Where(Function(x) x.CodigoProduto = itemNF.CodigoProduto)

                    If itemSaldo.SaldoPedidoFiscal > 0 Then

                        itemNF.SaldoPedidoFiscal = itemSaldo.SaldoPedidoFiscal
                        itemNF.SaldoPedidoFisico = itemSaldo.SaldoPedidoFisico
                        itemNF.SaldoValorMoeda = itemSaldo.SaldoValorMoeda
                        itemNF.SaldoValorOficial = itemSaldo.SaldoValorOficial

                    Else

                        itemNF.SaldoPedidoFiscal = 0
                        itemNF.SaldoPedidoFisico = 0
                        itemNF.SaldoValorMoeda = 0
                        itemNF.SaldoValorOficial = 0

                    End If


                    itemSaldo.SaldoPedidoFiscal -= itemNF.PesoFiscal
                    itemSaldo.SaldoPedidoFisico -= itemNF.PesoFiscal
                    itemSaldo.SaldoValorMoeda -= itemNF.ValorTotal
                    itemSaldo.SaldoValorOficial -= itemNF.ValorTotalMoeda

                Next

                'If _CodigoPedido = 0 Then
                '    _IndiceNota = New Cotacao(2, _DataNota).Indice 'PTAX
                'End If

                If Not Me.NotaDeTerceiro AndAlso Me.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then

                    Dim indiceDolar As Decimal

                    If Me.Pedido.CodigoIndexador = 99 OrElse Me.Pedido.IndexadorFixo Then
                        indiceDolar = Me.Pedido.IndiceFixado
                    Else
                        indiceDolar = New Cotacao(Me.Pedido.CodigoIndexador, Me.DataNota).Indice 'PTAX
                    End If

                    itemNF.UnitarioMoeda = itemNF.Unitario / indiceDolar
                    itemNF.ValorLiquidoMoeda = itemNF.ValorLiquido / indiceDolar
                    itemNF.ValorTotalMoeda = itemNF.ValorTotal / indiceDolar

                End If

                If itemNF.Encargos Is Nothing Then

                    'Poder instanciar os encargos dos itens que faltam

                End If

            Next

            If Not Me.NotaDeTerceiro AndAlso Me.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira Then
                Me.TotalNotaMoeda = Me.Itens.Sum(Function(x) x.ValorTotalMoeda)
                Me.TotalProdutoMoeda = Me.TotalNotaMoeda
                Me.TotalProdutoMoeda = Me.TotalNotaMoeda
            End If

        Else
            If erroMsg.Length > 0 Then
                If String.IsNullOrWhiteSpace(MsgAlerta) Then
                    Throw New Exception(erroMsg)
                Else
                    Throw New Exception(MsgAlerta)
                End If
            End If
        End If

    End Sub

    Public Function ChecarParametros(Optional pTipoPedido As String = "") As Hashtable

        '*******************************
        '***** PARAMETROS TRATATOS *****
        '*******************************
        'TipoPedido        nvarchar(5)  = NULL   1-Global/Direto 2-Afixar 3-Deposito
        'TipoApuracao      bit          = 0,     0 Sintetico - 1 Analitico        
        'Empresa           nvarchar(18) = NULL,
        'EndEmpresa        int          = NULL, 
        'Cliente           nvarchar(18) = NULL,        
        'EndCliente        int          = NULL, 
        'FilialDev         nvarchar(18) = NULL, 
        'EndFilialDev      int          = NULL, 
        'DataReferencia    datetime     = NULL, 
        'Safra             nvarchar(50) = NULL, 
        'Pedido            int          = NULL, 
        'Saldo             int          = NULL, -- 0 Sem Saldo, 1 Com Saldo, NULL Todos 
        'Fiscal            int          = NULL, -- 0 Fechado  , 1 Aberto   , NULL Todos
        'PeriodoInicial    datetime     = NULL,
        'PeriodoFinal      datetime     = NULL,
        'Operacao          int          = NULL,
        'SubOperacao       int          = NULL,
        'Produto           int          = NULL

        Dim parametros As New Hashtable

        parametros.Add("Empresa", Me.CodigoEmpresa)
        parametros.Add("EndEmpresa", Me.EnderecoEmpresa)
        parametros.Add("Cliente", CodigoCliente)

        If Me.MicSistemas And Me.NotaDeTerceiro = False Then
            parametros.Add("Contrato", Me.CodigoPedidoMIC)
        ElseIf Me.NotaDeTerceiro = True Then
            parametros.Add("Pedido", Me.CodigoPedido)
        Else
            parametros.Add("Pedido", Me.CodigoPedido)
        End If

        Return parametros
    End Function

    Public Sub Devolucao()

        Dim devolucaoMsg As String = String.Empty

        If Me.SubOperacao.Devolucao AndAlso Not Me.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then
            For Each item As [Lib].Negocio.NotaFiscalXItem In Me.Itens
                devolucaoMsg += " " & item.NotasDevolucao.MsgDevolucao & " "
            Next

            'Realiza o estorno nos tétulos através das NFés associadas a NF de Devolução
            If Not Me.VencimentosPedido Is Nothing AndAlso Me.VencimentosPedido.Count > 0 Then
                Me.VencimentosNota.ReajFinanceiro = New ReajusteFinanceiro(Me, False)
                Me.VencimentosNota.ReajFinanceiro.ReajustaNotaDeDevolucao()
            End If
        Else
            'Historico Financeiro da NF
            If Not Me.VencimentosPedido Is Nothing AndAlso Me.VencimentosPedido.Count > 0 AndAlso Not Me.SubOperacao.Classe = eClassesOperacoes.REMESSAS Then

                If Me.VencimentosPedido.Where(Function(x) x.CodigoProvisao = eProvisao.Provisao).Count > 0 Then
                    Dim provisao = Me.VencimentosPedido.Where(Function(x) x.CodigoProvisao = eProvisao.Provisao).First()

                    provisao.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(provisao, Me, 0, Me.TotalNota, "Retirado o valor de R$ : " & Me.TotalNota.ToString("C2", CultureInfo.CurrentCulture) & ", ref a NF: " & Me.Codigo & "-" & Me.Serie & "-" & IIf(Me.EntradaSaida = eEntradaSaida.Entrada, "E", "S")))
                    Me.VencimentosNota.ForEach(Sub(v) v.ListHistoricoFinanceiro.Add(New HistoricoFinanceiro(v, Me, provisao.Codigo, v.ValorDoDocumento, "Adicionado previsão " & v.Historico)))

                End If
            End If
        End If

        If devolucaoMsg.Length > 0 Then Me.ObservacoesDeEmbarque = Me.ObservacoesDeEmbarque & devolucaoMsg

        If Me.SubOperacao.Classe = eClassesOperacoes.EXPORTACOES Then
            If Not Me.DadosDaExportacaoRE Is Nothing Then
                Dim msgExportacao As String = ""
                Dim strVirgula As String = "| RE: "
                For Each row As NotaFiscalXRE In Me.DadosDaExportacaoRE
                    msgExportacao &= strVirgula & row.RegistroDeExportacao & " - " & row.UfProdutor
                    strVirgula = ", "
                Next
                If msgExportacao.Length > 0 Then Me.ObservacoesDeEmbarque = Me.ObservacoesDeEmbarque & msgExportacao
            End If
        End If

    End Sub

    Public Sub ObservacoesEmbarque()

        Me.ObservacoesDeEmbarque = Me.ObservacoesDeEmbarque & "|USUARIO EMISSOR: " & Me.UsuarioInclusao

        For Each item In Me.Itens
            If item.Encargos.Where(Function(s) s.Codigo).ToString.Contains("DESONERADO") Then
                If item.ObservacoesDoProduto.Length = 0 Then
                    item.ObservacoesDoProduto = "IMPOSTOS: ICMS DESONERADO " & item.Encargos.Where(Function(s) s.Codigo = "ICMS DESONERADO").Sum(Function(s) s.Valor = 0)
                Else
                    item.ObservacoesDoProduto = item.ObservacoesDoProduto & "|IMPOSTOS: ICMS DESONERADO " & item.Encargos.Where(Function(s) s.Codigo = "ICMS DESONERADO").Sum(Function(s) s.Valor)
                End If
            End If
        Next

    End Sub

    Public Function enviarEstoqueMinimo(ByVal objEstoqueMinimo As DataTable) As Boolean

        Dim estoqueOK As Boolean = True

        Dim lstMail As New List(Of String)

        Dim Filial As String = String.Empty
        Dim bodyHTML = String.Empty

        bodyHTML &= "<h1 style='clear: left; color: red; display: block; line-height: 105%; margin: 2px; padding: 7px 5px; position: relative; font-size: 150%; font-weight: bold;'>Alerta para o Estoque ménimo</h1>" & vbCrLf

        bodyHTML &= "<br/>Sr(s), <br/><br/>" & vbCrLf

        If objEstoqueMinimo.Rows.Count = 1 Then
            bodyHTML &= "Emitindo Nota Fiscal, o Produto abaixo merece atenção: <br/>" & vbCrLf
        Else
            bodyHTML &= "Emitindo Nota Fiscal, os Produtos abaixo merecem atenção: <br/>" & vbCrLf
        End If

        'bodyHTML &= "<ul>"
        If String.IsNullOrEmpty(Filial) Then Filial = Me.CodigoEmpresa

        bodyHTML &= "<table border='1' bordercolor='#000000' cellpadding='0' cellspacing='0'>" & vbCrLf

        bodyHTML &= "<tr>" & vbCrLf
        bodyHTML &= "<td>&nbsp;PRODUTO&nbsp;</td>" & vbCrLf
        bodyHTML &= "<td style='text-align:right;'>&nbsp;ESTOQUE MéNIMO&nbsp;</td>" & vbCrLf
        bodyHTML &= "<td style='text-align:right;'>&nbsp;FATURANDO&nbsp;</td>" & vbCrLf
        bodyHTML &= "<td style='text-align:right;'>&nbsp;SALDO&nbsp;</td>" & vbCrLf
        bodyHTML &= "</tr>" & vbCrLf


        For Each drItemE As DataRow In objEstoqueMinimo.Rows
            bodyHTML &= "<tr>" & vbCrLf
            bodyHTML &= "<td>&nbsp;" & drItemE("Produto") & "-" & drItemE("Nome") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "<td style='text-align:right;'>&nbsp;" & drItemE("EstoqueMinimo") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "<td style='text-align:right;'>&nbsp;" & drItemE("Faturando") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "<td style='text-align:right;'>&nbsp;" & drItemE("Saldo") & "&nbsp;</td>" & vbCrLf
            bodyHTML &= "</tr>" & vbCrLf
        Next

        bodyHTML &= "</table>" & vbCrLf

        bodyHTML &= "<br/><font size='2'>Este é um e-mail gerado automaticamente. Por favor, não responder.</font>" & vbCrLf

        bodyHTML &= "<br/><br/><font size='2'>é Copyright " & Now.ToString("yyyy") & " NGS Soluéées Ltda - Todos os direitos reservados</font>" & vbCrLf

        Dim objUsuario As New [Lib].Negocio.Usuario(Me.UsuarioInclusao)
        If objUsuario IsNot Nothing AndAlso Not String.IsNullOrWhiteSpace(objUsuario.Email) AndAlso objUsuario.Email.Trim() <> "&nbsp;" Then
            lstMail.Add(objUsuario.Email)
        End If

        Dim Sql As String = "SELECT u.Email " & vbCrLf &
                      "  FROM ConfiguracaoXUsuario cxu " & vbCrLf &
                      " INNER JOIN Usuarios u " & vbCrLf &
                      "    ON (u.Usuario_Id = cxu.Usuario_Id) " & vbCrLf &
                      " WHERE cxu.Etapa_Id = " & eEtapa.EstoqueMinimo

        Dim objBanco As New AcessaBanco
        Dim dsMail As DataSet = objBanco.ConsultaDataSet(Sql, "ConfiguracaoXUsuario")
        If dsMail IsNot Nothing AndAlso dsMail.Tables IsNot Nothing AndAlso dsMail.Tables.Count > 0 AndAlso dsMail.Tables(0).Rows.Count > 0 Then
            For Each row As DataRow In dsMail.Tables(0).Rows
                If Not String.IsNullOrWhiteSpace(row("Email")) Then
                    lstMail.Add(row("Email"))
                End If
            Next
        End If

        'bodyHTML &= "</ul>"

        Dim EmprMail = New [Lib].Negocio.Cliente(Filial, 0)
        Dim errorMsgMail As String = String.Empty
        Dim subject As String = String.Format("Empresa " & EmprMail.CodigoFormatado & " (" & EmprMail.Cidade & "/" & EmprMail.CodigoEstado & ") - Atenção para o Estoque mÍnimo - {0:dd/MM/yyyy HH:mm}", DateTime.Now)
        Dim smtp = Funcoes.GetSmtpSettings()
        Dim fromMail = Funcoes.GetFromMail()

        ServicePointManager.SecurityProtocol = CType(&HC00, SecurityProtocolType)

        If lstMail IsNot Nothing AndAlso lstMail.Count > 0 Then
            Funcoes.SendMail(fromMail, "NGS SOLUÇÕES", lstMail, subject, bodyHTML, smtp, errorMsgMail)
        End If

        Return True
    End Function

    Public Function PedidoBloqueado() As Boolean

        If Not Me.Pedido.Bloquear() Then
            Return True
        End If

        Return False

    End Function

    Public Function RealizarEstorno() As Boolean

        'REALIZA ESTORNO NO PEDIDO NAS DEVOLUCOES DE NOTAS GERAIS
        If FinanceiroNovo AndAlso Me.Pedido.TemNFG() AndAlso Me.SubOperacao.Devolucao Then
            Dim objPedido As New Pedido
            Dim i As Integer = 0
            objPedido = Me.Pedido
            objPedido.IUD = "U"
            For Each item As PedidoXItem In objPedido.Itens
                Dim codigoProduto = item.CodigoProduto
                item.IUD = "U"
                Dim notaXItem As New NotaFiscalXItem()
                notaXItem = Me.Itens.FirstOrDefault(Function(P) P.CodigoProduto = codigoProduto)

                If (notaXItem IsNot Nothing) Then
                    Dim objLancamento As New [Lib].Negocio.LancamentoItemPedido(item)

                    objLancamento.IUD = "I"
                    objLancamento.CodigoPedidoItem = item.Lancamentos.Count + 1
                    objLancamento.TipoLancamento = eTiposLancamentosPedidos.Estorno
                    objLancamento.Movimento = Me.Movimento
                    objLancamento.DataEntrega = Me.Movimento
                    objLancamento.QuantidadeFaturamento = notaXItem.QuantidadeFiscal
                    objLancamento.UnitarioOficialCompra = notaXItem.Unitario
                    objLancamento.UnitarioOficial = notaXItem.Unitario
                    objLancamento.UnitarioMoeda = notaXItem.UnitarioMoeda
                    objLancamento.TotalOficial = notaXItem.ValorTotal
                    objLancamento.TotalMoeda = notaXItem.UnitarioMoeda * notaXItem.QuantidadeFiscal
                    objLancamento.QuantidadeComercializacao = notaXItem.QuantidadeFiscal

                    item.Lancamentos.Add(objLancamento)
                    item.Encargos = Nothing
                    item.Encargos.CriaListar()

                    If Not item.Salvar() Then
                        Return False
                    End If
                End If
            Next

        End If

        Return True

    End Function

    Public Function VerificarLiberaCarregamento(ByVal pCodigoPedido As Integer) As Boolean

        If Me.Empresa.Empresa.LiberaCarregamento Then
            Dim verPedido As Pedido = New Pedido(Me.CodigoEmpresa, Me.EnderecoEmpresa, pCodigoPedido)

            If Not verPedido.LiberaCarregamento AndAlso Not verPedido.SubOperacao.Devolucao AndAlso verPedido.SubOperacao.EntradaSaida = eEntradaSaida.Saida Then
                Return False
            End If
        End If

        Return True

    End Function

    Public Function NotaFiscalJaLancada() As String

        Dim sql As String = String.Empty

        sql = "SELECT ISNULL(NFG,0) NFG " & vbCrLf &
              "  FROM NOTASFISCAIS" & vbCrLf &
              "  WHERE Empresa_Id = '" & Me.CodigoEmpresa & "'" & vbCrLf &
              "     AND EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "     AND Cliente_Id = '" & Me.CodigoCliente & "'" & vbCrLf &
              "     AND EndCliente_Id = " & Me.EnderecoCliente & vbCrLf &
              "     AND EntradaSaida_Id = '" & IIf(Me.EntradaSaida = eEntradaSaida.Entrada, "E", "S") & "' " & vbCrLf &
              "     AND Serie_Id= '" & Me.SerieNotaProdutor & "'" & vbCrLf &
              "     AND Nota_Id= " & Me.NotaProdutor & vbCrLf

        Dim banco As New AcessaBanco
        Dim Ds As DataSet = banco.ConsultaDataSet(sql, "NF")

        If Ds.Tables(0).Rows.Count > 0 Then
            If Ds.Tables(0).Rows(0)("NFG") Then
                Return "Existe uma Nota Fiscal Geral jé emitida com esse Némero"
            Else
                Return "Existe uma Nota Fiscal jé emitida com esse Némero"
            End If
        Else
            Return ""
        End If
    End Function

    Private Function CarregarItensComSaldo(ByVal pedidoSaldo As [Lib].Negocio.SaldoPedido2015, ByRef erroMsg As String, ByRef MsgAlerta As String) As Boolean

        Me.CarregandoItens = True

        For Each objItemNF As NotaFiscalXItem In Me.Itens

            If (objItemNF.CodigoProduto Is Nothing OrElse objItemNF.CodigoProduto.Length = 0) AndAlso Not Me.Pedido Is Nothing AndAlso Not Me.Pedido.Codigo = 0 Then
                objItemNF.CodigoProduto = Me.Pedido.Itens(0).CodigoProduto
            End If

            objItemNF.CodigoOperacao = Me.CodigoOperacao
            objItemNF.CodigoSubOperacao = Me.CodigoSubOperacao

            If Not Me.Pedido Is Nothing AndAlso Not Me.Pedido.Codigo = 0 Then
                objItemNF.CodigoOperacaoEstado = Me.Pedido.Itens(0).CodigoOperacaoXEstado
            End If

            objItemNF.CarregandoEncargos = True
            objItemNF.Encargos = New [Lib].Negocio.ListNotaFiscalXItemXEncargo(objItemNF, New List(Of eEtapaEncago) From {eEtapaEncago.Normal})
            objItemNF.CarregandoEncargos = False

            Me.CarregandoItens = False

            'Levar informarção do InfaDProd do Produto para NFE caso tenha - Furlan - 21-12-2022
            If objItemNF.Produto.InfaDProd.Length > 0 Then objItemNF.ObservacoesDoProduto = objItemNF.Produto.InfaDProd

            objItemNF.CodigoPedido = pedidoSaldo.CodigoPedido

            objItemNF.PesoQuantidade = objItemNF.Produto.PesoQuantidade

            objItemNF.Classificacao = pedidoSaldo.Classificacao
            objItemNF.CodigoEmbalagem = pedidoSaldo.CodigoEmbalagem
            objItemNF.CodigoEmbalagemIndea = pedidoSaldo.EmbalagemIndea
            objItemNF.CodigoTipoDeEmbalagem = pedidoSaldo.TipoDeEmbalagem
            objItemNF.CapacidadeEmbalagem = pedidoSaldo.CapacidadeEmbalagem

            '**********************************************************************************************************************************
            '***************************************** 3 - SALDOS  ****************************************************************************
            '**********************************************************************************************************************************
            'Verifica o saldo do pedido de acordo com a Operacao da Nota

            'GLOBAL é DEVOLUCAO
            If Me.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Me.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.QtdeEntregueFiscalGlobal - pedidoSaldo.QtdeEntregueFiscalRemessa
                objItemNF.SaldoPedidoFisico = pedidoSaldo.QtdeEntregueFiscalGlobal - pedidoSaldo.QtdeEntregueFiscalRemessa
                objItemNF.SaldoValorOficial = pedidoSaldo.VlrNotaOficialGlobalBruto - pedidoSaldo.VlrNotaOficialRemessaBruto
                objItemNF.SaldoValorMoeda = pedidoSaldo.VlrNotaMoedaGlobalBruto - pedidoSaldo.VlrNotaMoedaRemessaBruto
                'GLOBAL NAO é DEVOLUCAO
            ElseIf Me.SubOperacao.Classe = eClassesOperacoes.GLOBAL And Not Me.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.SaldoQtdeGlobalFiscal
                objItemNF.SaldoPedidoFisico = pedidoSaldo.SaldoQtdeGlobalFiscal
                objItemNF.SaldoValorOficial = pedidoSaldo.SaldoValorOficialGlobalDireto
                objItemNF.SaldoValorMoeda = pedidoSaldo.SaldoValorMoedaGlobalDireto
                'REMESSA é DEVOLUCAO
            ElseIf Me.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Me.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.QtdeEntregueFiscalRemessa
                objItemNF.SaldoPedidoFisico = pedidoSaldo.QtdeEntregueFisicoRemessa
                objItemNF.SaldoValorOficial = pedidoSaldo.VlrNotaOficialRemessaBruto
                objItemNF.SaldoValorMoeda = pedidoSaldo.VlrNotaMoedaRemessaBruto
                'REMESSA NAO é DEVOLUCAO
            ElseIf Me.SubOperacao.Classe = eClassesOperacoes.REMESSAS And Not Me.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.SaldoQtdeRemessaFiscal
                objItemNF.SaldoPedidoFisico = pedidoSaldo.SaldoQtdeRemessaFisica
                objItemNF.SaldoValorOficial = pedidoSaldo.SaldoValorOficialRemessa
                objItemNF.SaldoValorMoeda = pedidoSaldo.SaldoValorMoedaRemessa
                'AFIXAR é DEVOLUCAO
            ElseIf Me.SubOperacao.Classe = eClassesOperacoes.AFIXAR And Me.SubOperacao.Devolucao Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.QtdeEntregueFiscalAFixar - pedidoSaldo.QtdeFixacao
                objItemNF.SaldoPedidoFisico = pedidoSaldo.QtdeEntregueFisicoAFixar - pedidoSaldo.QtdeFixacao
                'Criado um novo campo na StoredProcedure spSaldoPedido para trazer o valor
                'total da fixação, mas correspondente ao unitério da NF. 
                'O campo VlrfixacaoNF seré utilizado quando for menor  do que o campo VlrfixacaoOficial senão este éltimo seré, para que o valor da devolução fique correto.
                objItemNF.SaldoValorOficial = pedidoSaldo.VlrNotaOficialAFixarBruto - IIf(pedidoSaldo.VlrFixacaoNF < pedidoSaldo.VlrFixacaoOficial, pedidoSaldo.VlrFixacaoNF, pedidoSaldo.VlrFixacaoOficial)
                objItemNF.SaldoValorMoeda = pedidoSaldo.VlrNotaMoedaAFixarBruto - pedidoSaldo.VlrFixacaoMoeda

                'Devolução utilizada ou para Reajuste de unitério ou para devolução de valor.
            ElseIf Me.SubOperacao.Devolucao And pedidoSaldo.Tipo = 1 AndAlso Not Me.SubOperacao.QuantidadeFiscal Then
                objItemNF.SaldoPedidoFiscal = Math.Abs(pedidoSaldo.SaldoQtdeDiretoFiscal)
                objItemNF.SaldoPedidoFisico = Math.Abs(pedidoSaldo.SaldoQtdeDiretoFisica)
                objItemNF.SaldoValorOficial = Math.Abs(pedidoSaldo.VlrNotaOficialDiretaBruto)
                objItemNF.SaldoValorMoeda = Math.Abs(pedidoSaldo.VlrNotaMoedaDiretaBruto)
            ElseIf Me.SubOperacao.Devolucao And pedidoSaldo.Tipo = 1 Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.QtdeEntregueFiscalDireta
                objItemNF.SaldoPedidoFisico = pedidoSaldo.QtdeEntregueFisicoDireta
                objItemNF.SaldoValorOficial = pedidoSaldo.VlrNotaOficialDiretaBruto
                objItemNF.SaldoValorMoeda = pedidoSaldo.VlrNotaMoedaDiretaBruto
            ElseIf Me.SubOperacao.Devolucao And pedidoSaldo.Tipo = 3 Then
                objItemNF.SaldoPedidoFiscal = pedidoSaldo.QtdeEntregueFiscalDeposito
                objItemNF.SaldoPedidoFisico = pedidoSaldo.QtdeEntregueFisicoDeposito
                objItemNF.SaldoValorOficial = pedidoSaldo.VlrNotaOficialDepositoBruto
                objItemNF.SaldoValorMoeda = pedidoSaldo.VlrNotaMoedaDepositoBruto
            Else
                'Verificar ainda pode nao estar contemplando tudo
                If pedidoSaldo.QtdeProgramada = 0 And pedidoSaldo.QtdeProgramadaComercializacao > 0 Then
                    objItemNF.SaldoPedidoFiscal = pedidoSaldo.SaldoQtdeComercializacao
                    objItemNF.SaldoPedidoFisico = pedidoSaldo.SaldoQtdeComercializacao
                    objItemNF.SaldoValorOficial = pedidoSaldo.SaldoValorOficialGlobalDireto
                    objItemNF.SaldoValorMoeda = pedidoSaldo.SaldoValorMoedaGlobalDireto
                Else
                    objItemNF.SaldoPedidoFiscal = pedidoSaldo.SaldoQtdeDiretoFiscal
                    objItemNF.SaldoPedidoFisico = pedidoSaldo.SaldoQtdeDiretoFisica
                    objItemNF.SaldoValorOficial = pedidoSaldo.SaldoValorOficialGlobalDireto
                    objItemNF.SaldoValorMoeda = pedidoSaldo.SaldoValorMoedaGlobalDireto
                End If
            End If

            '********************************************************************************************************************************************
            '***************************************** 3.1 - SALDOS vs ESTOQUE **************************************************************************
            '********************************************************************************************************************************************
            'Verifica o saldo do pedido de acordo com a Operacao da Nota não é maior que o saldo do produto em Estoque
            'Se o saldo Quantidade do produto em Estoque "Sem Embalagem" for menor que o saldo do pedido o saldo do pedido assume a Quantidade em Estoque
            'Acrescentei a classe CONTAEORDEM pois o estoque sé deve ser checado para a VENDA OU COMPRA A ORDEM - FURLAN - 22/03/16 

            If Me.SubOperacao.Classe <> eClassesOperacoes.GLOBAL _
            And Me.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
            And Me.SubOperacao.EntradaSaida = eEntradaSaida.Saida _
            And objItemNF.Produto.ControlarEstoque Then

                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objItemNF.CodigoProduto)
                SaldoProdutoEstoque.CarregarResumoSaldoEmEstoque(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.SubOperacao)

                If objItemNF.SubOperacao.EstoqueFisico Or objItemNF.SubOperacao.EstoqueFiscal Then

                    If Not objItemNF.SubOperacao.Devolucao AndAlso objItemNF.SubOperacao.EntradaSaida = eEntradaSaida.Saida AndAlso objItemNF.SubOperacao.EstoqueFiscal AndAlso pedidoSaldo.Produto.EstoqueMinimo > 0 Then

                        If (SaldoProdutoEstoque.SaldoFiscal - objItemNF.SaldoPedidoFiscal) < pedidoSaldo.Produto.EstoqueMinimo Then

                            'Verificar a nessecidade de alimentar o objeto com o estoque minimo, na tela de nota
                            'exite este processo, foi comentado, pq não foi comprovado a necessidade
                            'Dim ItemEstoqueMinimo As DatapedidoSaldo = CType(Session("objEstoqueMinimo" & HID.Value), DataTable).NewpedidoSaldo()

                            'ItemEstoqueMinimo("Produto") = pedidoSaldo.Produto.Codigo

                            'If Left(Me.CodigoEmpresa, 8) = "03189063" Then
                            '    ItemEstoqueMinimo("Nome") = pedidoSaldo.Produto.Nome & "-" & pedidoSaldo.Produto.Descricao
                            'Else
                            '    ItemEstoqueMinimo("Nome") = pedidoSaldo.Produto.Nome
                            'End If

                            'ItemEstoqueMinimo("EstoqueMinimo") = pedidoSaldo.Produto.EstoqueMinimo.ToString("N4")
                            'ItemEstoqueMinimo("Faturando") = objItemNF.SaldoPedidoFiscal
                            'ItemEstoqueMinimo("Saldo") = SaldoProdutoEstoque.SaldoFiscal - objItemNF.SaldoPedidoFiscal
                            'CType(Session("objEstoqueMinimo" & HID.Value), DataTable).pedidoSaldos.Add(ItemEstoqueMinimo)

                        End If
                    End If

                    If SaldoProdutoEstoque.Count = 0 Then
                        objItemNF.SaldoValorOficial = 0
                        objItemNF.SaldoValorMoeda = 0
                        objItemNF.SaldoPedidoFiscal = 0
                        objItemNF.SaldoPedidoFisico = 0
                        erroMsg &= "Produto " & pedidoSaldo.CodigoProduto & " - " & pedidoSaldo.NomeProduto & " sem estoque. \n"
                    Else
                        Dim faltaSaldo As Boolean = False

                        If objItemNF.SubOperacao.EstoqueFisico AndAlso objItemNF.SaldoPedidoFisico > SaldoProdutoEstoque.SaldoFisico Then
                            MsgAlerta &= "O Produto " & pedidoSaldo.CodigoProduto & " - " & pedidoSaldo.NomeProduto & " tem em Estoque FISICO -> " & SaldoProdutoEstoque(0).SaldoFisico.ToString() & " que é menor que o saldo do Pedido -> " & objItemNF.SaldoPedidoFisico & ". Seré liberado apenas o que tem em Estoque. \n"

                            faltaSaldo = True
                        End If

                        If objItemNF.SubOperacao.EstoqueFiscal AndAlso objItemNF.SaldoPedidoFiscal > SaldoProdutoEstoque.SaldoFiscal Then
                            MsgAlerta &= "O Produto " & pedidoSaldo.CodigoProduto & " - " & pedidoSaldo.NomeProduto & " tem em Estoque FISCAL -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & objItemNF.SaldoPedidoFiscal & ". Seré liberado apenas o que tem em Estoque. \n"

                            faltaSaldo = True
                        End If

                        If faltaSaldo Then
                            objItemNF.SaldoPedidoFiscal = SaldoProdutoEstoque.SaldoFiscal
                            objItemNF.SaldoPedidoFisico = SaldoProdutoEstoque.SaldoFisico
                            If pedidoSaldo.UnitarioOficial > 0 Then
                                objItemNF.SaldoValorOficial = Math.Round(SaldoProdutoEstoque.SaldoFiscal * pedidoSaldo.UnitarioOficial, 2, MidpointRounding.AwayFromZero)
                                objItemNF.SaldoValorMoeda = Math.Round(SaldoProdutoEstoque.SaldoFiscal * pedidoSaldo.UnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                            End If
                        End If
                    End If
                End If
            End If

            'Na Devolucao Apura o estoque do produto de acordo com o lote/classificacao que foi recebido
            'Acrescentei a classe CONTAEORDEM pois o estoque sé deve ser checado para a VENDA OU COMPRA A ORDEM - FURLAN - 22/03/16 
            If Me.SubOperacao.Devolucao _
             And Me.SubOperacao.Classe <> eClassesOperacoes.GLOBAL _
             And Me.SubOperacao.Classe <> eClassesOperacoes.CONTAEORDEM _
             And Me.SubOperacao.EntradaSaida = eEntradaSaida.Saida _
             And objItemNF.Produto.ControlarEstoque _
             And (objItemNF.Produto.ControlarEmbalagem Or objItemNF.Produto.ControlarLote) _
             And Me.SubOperacao.EstoqueFiscal _
            Then
                Dim SaldoProdutoEstoque As New [Lib].Negocio.ListSaldoProdutoEstoqueLoteEmbalagem(objItemNF.CodigoProduto)
                SaldoProdutoEstoque.CarregaSaldoDisponivelSaidaLoteClassificao(Me.SubOperacao, Me.CodigoEmpresa, Me.EnderecoEmpresa, objItemNF.Lote, objItemNF.Classificacao, objItemNF.CodigoEmbalagem, objItemNF.CodigoTipoDeEmbalagem, objItemNF.CapacidadeEmbalagem)

                If SaldoProdutoEstoque.Count = 0 Then
                    objItemNF.SaldoPedidoFiscal = 0
                    objItemNF.SaldoPedidoFisico = 0
                    objItemNF.SaldoValorOficial = 0
                    objItemNF.SaldoValorMoeda = 0
                    erroMsg &= "Produto não pode ser liberado, saldo em Estoque = 0. \n" & vbCrLf
                Else
                    Dim faltaSaldo As Boolean = False

                    If objItemNF.SubOperacao.EstoqueFiscal AndAlso objItemNF.SaldoPedidoFiscal > SaldoProdutoEstoque(0).SaldoFiscal Then
                        MsgAlerta &= "O Produto " & pedidoSaldo.CodigoProduto & " - " & pedidoSaldo.NomeProduto & " tem em Estoque FISCAL -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & SaldoProdutoEstoque(0).SaldoFiscal & ". Seré liberado apenas o que tem em Estoque. \n"

                        faltaSaldo = True
                    End If

                    If objItemNF.SubOperacao.EstoqueFisico AndAlso objItemNF.SaldoPedidoFisico > SaldoProdutoEstoque(0).SaldoFisico Then
                        MsgAlerta &= "O Produto " & pedidoSaldo.CodigoProduto & " - " & pedidoSaldo.NomeProduto & " tem em Estoque FISICO -> " & SaldoProdutoEstoque(0).SaldoFiscal.ToString() & " que é menor que o saldo do Pedido -> " & SaldoProdutoEstoque(0).SaldoFiscal & ". Seré liberado apenas o que tem em Estoque. \n"

                        faltaSaldo = True
                    End If

                    If faltaSaldo Then
                        objItemNF.SaldoPedidoFiscal = SaldoProdutoEstoque(0).SaldoFiscal
                        objItemNF.SaldoPedidoFisico = SaldoProdutoEstoque(0).SaldoFisico
                        objItemNF.SaldoValorOficial = Math.Round(SaldoProdutoEstoque(0).SaldoFiscal * pedidoSaldo.UnitarioOficial, 2, MidpointRounding.AwayFromZero)
                        objItemNF.SaldoValorMoeda = Math.Round(SaldoProdutoEstoque(0).SaldoFiscal * pedidoSaldo.UnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                    End If
                End If
            End If
            '********************
            '**** 3 - FIM *******
            '********************

            If Me.XMLImportado = False Then

                If pedidoSaldo.UnitarioOficial > 0 Then
                    objItemNF.Unitario = pedidoSaldo.UnitarioOficial
                End If

                If pedidoSaldo.UnitarioMoeda > 0 Then
                    objItemNF.UnitarioMoeda = pedidoSaldo.UnitarioMoeda
                End If


                If pedidoSaldo.XmlvProd > 0 Then
                    objItemNF.ValorTotal = pedidoSaldo.XmlvProd
                    objItemNF.ValorTotalMoeda = 0
                Else
                    'Quando a nota é somente valor, ela nao tem cotacao e o valor é somente em reais
                    If Not Me.SubOperacao.QuantidadeFiscal Then
                        'Pendente alimentar o valor
                        objItemNF.ValorTotal = objItemNF.SaldoValorOficial
                        objItemNF.ValorTotalMoeda = 0
                    Else
                        objItemNF.ValorTotal = Math.Round(objItemNF.QuantidadeFiscal * objItemNF.Unitario, 2, MidpointRounding.AwayFromZero)
                        objItemNF.ValorTotalMoeda = Math.Round(objItemNF.QuantidadeFiscal * objItemNF.UnitarioMoeda, 2, MidpointRounding.AwayFromZero)
                    End If
                End If

            End If


            '**********************************************************************************************************************************
            '*********************************************** OBSERVACOES DO PRODUTO ***********************************************************
            '**********************************************************************************************************************************
            If objItemNF.EmbalagemProduto IsNot Nothing AndAlso objItemNF.EmbalagemProduto.PesoVariavel Then
                objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.ProdutoLoteClassificacao.PesoSaco
                objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & IIf(objItemNF.EmbalagemProduto.PesoVariavel, objItemNF.ProdutoLoteClassificacao.PesoSaco, objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade) & " / " & objItemNF.Produto.DescricaoTecnica
            ElseIf objItemNF.CapacidadeEmbalagem > 0 Then
                objItemNF.QuantidadeDeEmbalagem = objItemNF.QuantidadeFiscal / objItemNF.CapacidadeEmbalagem
                objItemNF.ObservacoesDoProduto &= objItemNF.QuantidadeDeEmbalagem & " " & objItemNF.Embalagem.Descricao & " " & objItemNF.TipoDeEmbalagem.Descricao & " " & objItemNF.CapacidadeEmbalagem & " " & objItemNF.Produto.Unidade & " / " & objItemNF.Produto.DescricaoTecnica
            End If

            If objItemNF.Lote.Length > 0 Then
                Dim L As New [Lib].Negocio.Lote(objItemNF.Lote, objItemNF.CodigoProduto)
                If L.Tipo = 2 Then
                    objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                Else
                    objItemNF.ObservacoesDoProduto &= IIf(objItemNF.ObservacoesDoProduto.Length = 0, "", ", ") & "Lote: " & objItemNF.Lote & " GER " & L.Germinacao & IIf(L.Pureza = 0, "", " PUR " & L.Pureza) & " Peneira/Classif. " & objItemNF.Classificacao & " Renasem " & L.Renasem & " Validade " & L.DataValidade.ToString("dd-MM-yyyy")
                End If
            End If

            If Not String.IsNullOrWhiteSpace(pedidoSaldo.XmlinfAdProd) Then
                objItemNF.ObservacoesDoProduto &= pedidoSaldo.XmlinfAdProd
            End If
            '**********************************************************************************************************************************
            '**********************************************************************************************************************************

            If objItemNF.Encargos.Count = 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", não tem encargos cadastrados na Operação:" & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " \n"
            ElseIf objItemNF.CFOP = 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação:" & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " está sem CFOP informada \n"
            ElseIf objItemNF.CFOP < 5000 And Me.EntradaSaida = eEntradaSaida.Saida Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " com CFOP " & objItemNF.CFOP & " não pode ser usada na saéda. \n"
            ElseIf objItemNF.CFOP > 5000 And Me.EntradaSaida = eEntradaSaida.Entrada Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & ", Operação " & objItemNF.CodigoOperacao & "-" & objItemNF.CodigoSubOperacao & " com CFOP " & objItemNF.CFOP & " não pode ser usada na entrada. \n"
            ElseIf Not Me.NotaDeTerceiro AndAlso (Me.SubOperacao.QuantidadeFiscal And Me.Pedido.Moeda.Classificacao = eTiposMoeda.MoedaEstrangeira And objItemNF.SaldoValorMoeda <= 0 And objItemNF.SubOperacao.Classe <> eClassesOperacoes.AFIXAR) Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo 0 ou negativo não pode ser usado para emissão de Nota Fiscal. Saldo: " & objItemNF.SaldoValorMoeda & " Pedido em moeda Extrangeira. \n"
            ElseIf Not Me.NotaDeTerceiro AndAlso (IIf(Me.Pedido.Moeda.Classificacao = eTiposMoeda.Oficial, objItemNF.SaldoValorOficial, objItemNF.SaldoValorMoeda) <= 0 _
                And Me.SubOperacao.Classe <> eClassesOperacoes.DEPOSITOS _
                And Me.SubOperacao.Classe <> eClassesOperacoes.AFIXAR _
                And Me.SubOperacao.Classe <> eClassesOperacoes.FRETES _
                And Me.Operacao.CodigoClasse <> eClassesOperacoes.COMPRAS.ToString _
                And Not Me.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
                And Not Me.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
                And Me.SubOperacao.QuantidadeFiscal _
                And Not Me.SubOperacao.Devolucao) Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com Saldo de Valor 0 ou negativo e não pode ser usado para emissão de Nota Fiscal. Saldo: " & objItemNF.SaldoValorOficial & " \n"
            ElseIf pedidoSaldo.QtdeProgramadaComercializacao < 0 Then
                erroMsg &= "O Produto " & objItemNF.Produto.Nome & " esta com saldo negativo na Quantidade Programada para Comercialização e não pode ser usado para emissão de Nota Fiscal. Saldo: " & pedidoSaldo.QtdeProgramadaComercializacao & " \n"
            ElseIf objItemNF.Encargos.EncProduto.SituacaoTributariaPISCOFINS = 0 Then
                erroMsg &= "Situação Tributéria PISCOFINS não cadastrada para o Produto " & objItemNF.CodigoProduto & "-" & objItemNF.Produto.Nome & ", Operação " & Me.CodigoOperacao & "-" & Me.CodigoSubOperacao & " " & Me.SubOperacao.Descricao
            ElseIf (Not Me.SubOperacao.QuantidadeFiscal AndAlso (Math.Abs(objItemNF.SaldoValorOficial) > 0 Or Me.SubOperacao.Devolucao)) _
            OrElse Me.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE ICMS") _
            OrElse Me.SubOperacao.Descricao.ToString.Contains("COMPLEMENTACAO DE IPI") _
            OrElse Me.SubOperacao.Descricao.ToString.Contains("TRANSFERENCIA DE CREDITO DE ICMS") _
            OrElse (Me.SubOperacao.Devolucao = True AndAlso Me.SubOperacao.QuantidadeFiscal AndAlso objItemNF.SaldoValorOficial > 0 AndAlso objItemNF.ValorTotal > 0) _
            OrElse (Me.SubOperacao.EntradaSaida = eEntradaSaida.Entrada AndAlso Me.SubOperacao.Devolucao = True AndAlso objItemNF.SaldoValorOficial > 0) _
            OrElse (pedidoSaldo.QtdeProgramada = 0 AndAlso pedidoSaldo.QtdeProgramadaComercializacao > 0 AndAlso ((Me.SubOperacao.Devolucao AndAlso pedidoSaldo.QtdeComercializacaoEntregue > 0) OrElse (Me.SubOperacao.Devolucao = False AndAlso pedidoSaldo.SaldoQtdeComercializacao))) _
            OrElse (objItemNF.QuantidadeFiscal > 0 OrElse Me.SubOperacao.Classe = eClassesOperacoes.DEPOSITOS OrElse Me.SubOperacao.Classe = eClassesOperacoes.AFIXAR OrElse Me.SubOperacao.Classe = eClassesOperacoes.GLOBAL OrElse Me.SubOperacao.Classe = eClassesOperacoes.FRETES) _
            AndAlso objItemNF.Encargos.Count > 0 Then
            Else
                erroMsg &= "O item não péde ser adicionado, verifique as configuraéées de Operaéées, Operaéées X Encargos, Saldo no Extrato de Pedido e outros."
            End If
        Next

        Me.AtualizaTotais()

        Return True
    End Function

    Public Sub ConsultaDadosBancarios()

        'If Me.VencimentosNota(i).CodigoProvisao = 1 Then Exit Sub
        'Dim objCliente As [Lib].Negocio.Cliente = Me.Cliente
        'Dim ucConsultaDadosBancarios As ucConsultaDadosBancarios = DirectCast(Me.NamingContainer.FindControl("ucConsultaDadosBancarios"), ucConsultaDadosBancarios)
        'If ucConsultaDadosBancarios IsNot Nothing Then
        '    ucConsultaDadosBancarios.Limpar()
        '    ucConsultaDadosBancarios.MainUserControl = Me
        '    ucConsultaDadosBancarios.CarregaGrid(objCliente.Codigo, objCliente.CodigoEndereco)
        'End If

    End Sub

    Public Sub CarregarVencimentosDaNota(ByRef sOrigem As String, ByVal ddlCarteira As DropDownList, ByVal ddlCondicaoPagamento As DropDownList, ByVal ddlTipoDePagto As DropDownList, ByVal parameters As Dictionary(Of String, Object))

        Dim ddl As New CarregarDDL

        If Me.SubOperacao.CodigoGrupoContas.Length = 9 Then
            ddl.Carregar(ddlCarteira, CarregarDDL.Tabela.CarteiraFinanceira, "ContaClientes = '' and Adiantamento = 'N' and Classificacao = '" & IIf(Me.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "P", "R") & "' ", True)
        Else
            If Me.SubOperacao.Descricao.Contains("COMPLEMENTACAO") Then
                ddl.Carregar(ddlCarteira, CarregarDDL.Tabela.CarteiraFinanceira, "ContaClientes = '" & Me.Pedido.SubOperacao.CodigoGrupoContas & "' and Adiantamento = 'N' and Classificacao = '" & IIf(Me.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "P", "R") & "' ", True)
            Else
                ddl.Carregar(ddlCarteira, CarregarDDL.Tabela.CarteiraFinanceira, "ContaClientes = '" & Me.SubOperacao.CodigoGrupoContas & "' and Adiantamento = 'N' and Classificacao = '" & IIf(Me.SubOperacao.EntradaSaida = eEntradaSaida.Entrada, "P", "R") & "' ", True)
            End If
        End If

        ddl.Carregar(ddlCondicaoPagamento, CarregarDDL.Tabela.CondicaoDePagamento, "")

        If Me.Cliente.ApenasAVista Then ddlCondicaoPagamento.Enabled = False

        Dim Parametros As New Hashtable
        Parametros.Clear()
        Parametros.Add("listarTudo", "N")

        ddl.Carregar(ddlTipoDePagto, CarregarDDL.Tabela.TipoDePagamento, "", True, Parametros)

        If ddlCarteira.Items.Count > 1 Then ddlCarteira.SelectedIndex = 1

        Dim codPgto As Integer

        If Me.Pedido.SubOperacao.Classe = eClassesOperacoes.AFIXAR AndAlso Me.SubOperacao.Classe = eClassesOperacoes.COMPLEMENTACOES Then
            codPgto = Me.Pedido.Itens(0).Fixacoes.Where(Function(s) s.Codigo = Me.Itens(0).CodigoFixacao).FirstOrDefault.CondicaoPagamento
        Else
            If Me.Pedido.CondicaoPagamento.Codigo > 0 Then
                codPgto = Me.Pedido.CondicaoPagamento.Codigo
            End If
        End If

        With ddlCondicaoPagamento
            .SelectedIndex = .Items.IndexOf(.Items.FindByValue(codPgto))
        End With

        If (parameters IsNot Nothing AndAlso parameters.ContainsKey("Origem")) Then
            sOrigem = parameters("Origem").ToString
        End If

        If sOrigem = "NF" Then
            'SessaoRecuperaNotaFiscal()
            If Me.VencimentosNota.Count = 0 Then
                If Me.SubOperacao.Devolucao Then
                    'objNotaFiscal.VencimentosNota.DevolucaoNota()
                    ddlCondicaoPagamento.Enabled = False
                Else
                    Me.VencimentosNota.ParcelarNota(codPgto)
                End If
            End If

            If Me.VencimentosNota IsNot Nothing AndAlso Me.VencimentosNota.Count > 0 Then
                For Each v In Me.VencimentosNota
                    If ddlCarteira.SelectedIndex > 0 Then v.CodigoCarteira = ddlCarteira.SelectedValue
                Next
            End If

        End If

    End Sub

    Public Function Clone() As NotaFiscal
        Dim nota As NotaFiscal = CType(Me.MemberwiseClone(), NotaFiscal)
        Return nota
    End Function

    Public Function TotalDeNotasFiscais() As Boolean

        Dim sql As String = String.Empty

        sql = "SELECT NOTA_ID " & vbCrLf &
              "  FROM NOTASFISCAIS" & vbCrLf &
              "  WHERE Empresa_Id = '" & Me.CodigoEmpresa & "'" & vbCrLf &
              "     AND EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf &
              "     AND Pedido = " & Me.CodigoPedido

        Dim banco As New AcessaBanco
        Dim Ds As DataSet = banco.ConsultaDataSet(sql, "NF")

        If Ds.Tables(0).Rows.Count > 1 Then
            Return True
        Else
            Return False
        End If
    End Function

#End Region

End Class
