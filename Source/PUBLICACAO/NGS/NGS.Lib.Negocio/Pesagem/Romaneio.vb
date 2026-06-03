Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis


'****************************************************************************************************************************************
'**************************************************  LISTA DE ROMANEIOS  ****************************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class ListRomaneio
    Inherits List(Of Romaneio)

#Region "Variaveis"
    Private sql As String
    Private Banco As New AcessaBanco
    Private ds As DataSet
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pPedido As Integer, ByVal pOperacao As Integer, ByVal pSubOperacao As Integer)
        sql = "SELECT R.Empresa_Id, R.EndEmpresa_Id, R.Romaneio_Id, R.EntradaSaida, R.Pedido, R.Deposito, R.EndDeposito, " & vbCrLf & _
              "       isnull(R.Destino,'') as Destino,  isnull(R.EndDestino,0) as EndDestino, isnull(R.Transbordo,'') as Transbordo, isnull(R.EndTransbordo,0) as EndTransbordo, R.Produto, R.Operacao," & vbCrLf & _
              "       R.SubOperacao, R.Movimento, R.PesoBruto, R.Desconto, R.PesoLiquido, isnull(R.Observacoes,'') as Observacoes, isnull(R.Processo,'') as Processo, isnull(R.Autorizacao,0) as Autorizacao, " & vbCrLf & _
              "       isnull(R.PrimeiraPesagem,0) AS PrimeiraPesagem, isnull(R.SegundaPesagem,0) AS SegundaPesagem, " & vbCrLf & _
              "       ISNULL(NxR.Nota_Id,0) AS Nota, NxR.Cliente_Id, NxR.EndCliente_Id, NxR.EntradaSaida_Id, NxR.Serie_Id " & vbCrLf & _
              "  FROM Romaneios R " & vbCrLf & _
              "	 LEFT JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
              "	   ON NxR.Empresa_Id    = R.Empresa_id " & vbCrLf & _
              "	  AND NxR.EndEmpresa_Id = R.EndEmpresa_Id " & vbCrLf & _
              "	  AND NxR.Romaneio_Id   = R.Romaneio_Id " & vbCrLf & _
              " Where R.Pedido       = " & pPedido & vbCrLf & _
              "   and R.Operacao     = " & pOperacao & vbCrLf & _
              "   and R.SubOperacao  = " & pSubOperacao

        ds = Banco.ConsultaDataSet(sql, "Romaneio")

        Dim Nota As NotaFiscal

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Rom As New Romaneio
            Rom.CodigoEmpresa = row("Empresa_Id")
            Rom.EnderecoEmpresa = row("EndEmpresa_Id")
            Rom.Codigo = row("Romaneio_Id")
            Rom.EntradaSaida = row("EntradaSaida")
            Rom.CodigoPedido = row("Pedido")
            Rom.CodigoDeposito = row("Deposito")
            Rom.EnderecoDeposito = row("EndDeposito")
            Rom.CodigoDestino = row("Destino")
            Rom.EnderecoDestino = row("EndDestino")
            Rom.CodigoTransbordo = row("Transbordo")
            Rom.EnderecoTransbordo = row("EndTransbordo")
            Rom.CodigoProduto = row("Produto")
            Rom.CodigoOperacao = row("Operacao")
            Rom.CodigoSubOperacao = row("SubOperacao")
            Rom.Movimento = row("Movimento")
            Rom.PrimeiraPesagem = row("PrimeiraPesagem")
            Rom.SegundaPesagem = row("SegundaPesagem")
            Rom.PesoBruto = row("PesoBruto")
            Rom.Desconto = row("Desconto")
            Rom.PesoLiquido = row("PesoLiquido")
            Rom.Observacoes = row("Observacoes")
            Rom.Processo = row("Processo")
            Rom.CodigoAutorizacao = row("Autorizacao")
            Rom.TemNotaFiscal = Not (row("Nota") = 0)

            If Rom.TemNotaFiscal Then
                Nota = New NotaFiscal()
                Nota.CodigoEmpresa = row("Empresa_Id")
                Nota.EnderecoEmpresa = row("EndEmpresa_Id")
                Nota.CodigoCliente = row("Cliente_Id")
                Nota.EnderecoCliente = row("EndCliente_Id")
                Nota.Codigo = row("Nota")
                Nota.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                Nota.Serie = row("Serie_Id")

                Rom.NF = New NotaFiscal(Nota)
            End If

            If ds.Tables(0).Rows.Count = 1 Then Rom.Principal = True

            Me.Add(Rom)
        Next
    End Sub

    Public Sub New(ByVal pLaudo As Pesagem)
        sql = "SELECT R.Empresa_Id, R.EndEmpresa_Id, R.Romaneio_Id, R.EntradaSaida, " & vbCrLf & _
              "       R.Pedido, R.Deposito, R.EndDeposito, isnull(R.Destino,'') AS Destino, isnull(R.EndDestino,0) AS EndDestino, " & vbCrLf & _
              "       isnull(R.Transbordo,'') AS Transbordo, isnull(R.EndTransbordo,0) AS EndTransbordo, R.Produto, R.Operacao, R.SubOperacao, " & vbCrLf & _
              "       R.Movimento, R.PesoBruto, R.Desconto, R.PesoLiquido, isnull(R.Observacoes,'') as Observacoes, " & vbCrLf & _
              "       isnull(R.Processo,'') AS Processo, isnull(R.Autorizacao,0) AS Autorizacao, " & vbCrLf & _
              "       isnull(R.PrimeiraPesagem,0) AS PrimeiraPesagem, isnull(R.SegundaPesagem,0) AS SegundaPesagem, " & vbCrLf & _
              "       ISNULL(NxR.Nota_Id,0) AS Nota, NxR.Cliente_Id, NxR.EndCliente_Id, NxR.EntradaSaida_Id, NxR.Serie_Id " & vbCrLf & _
              "  FROM Romaneios R " & vbCrLf & _
              "	INNER JOIN RomaneiosXPesagens RxP" & vbCrLf & _
              "	   ON R.Empresa_Id    = RxP.Empresa_Id" & vbCrLf & _
              "	  AND R.EndEmpresa_Id = RxP.EndEmpresa_Id" & vbCrLf & _
              "	  AND R.Romaneio_Id   = RxP.Romaneio_Id" & vbCrLf & _
              "	 LEFT JOIN NotasFiscaisXRomaneios NxR" & vbCrLf & _
              "	   ON NxR.Empresa_Id    = R.Empresa_id" & vbCrLf & _
              "	  AND NxR.EndEmpresa_Id = R.EndEmpresa_Id" & vbCrLf & _
              "	  AND NxR.Romaneio_Id   = R.Romaneio_Id" & vbCrLf & _
              " WHERE RxP.Empresa_Id    ='" & pLaudo.CodigoEmpresa & "'" & vbCrLf & _
              "   AND RxP.EndEmpresa_Id = " & pLaudo.EnderecoEmpresa & vbCrLf & _
              "   AND RxP.Pesagem_Id    = " & pLaudo.Codigo

        ds = Banco.ConsultaDataSet(sql, "Romaneio")

        Dim Nota As NotaFiscal

        For Each row As DataRow In ds.Tables(0).Rows
            Dim Rom As New Romaneio
            Rom.CodigoEmpresa = row("Empresa_Id")
            Rom.EnderecoEmpresa = row("EndEmpresa_Id")
            Rom.Codigo = row("Romaneio_Id")
            Rom.EntradaSaida = row("EntradaSaida")
            Rom.CodigoPedido = row("Pedido")
            Rom.CodigoDeposito = row("Deposito")
            Rom.EnderecoDeposito = row("EndDeposito")
            Rom.CodigoDestino = row("Destino")
            Rom.EnderecoDestino = row("EndDestino")
            Rom.CodigoTransbordo = row("Transbordo")
            Rom.EnderecoTransbordo = row("EndTransbordo")
            Rom.CodigoProduto = row("Produto")
            Rom.CodigoOperacao = row("Operacao")
            Rom.CodigoSubOperacao = row("SubOperacao")
            Rom.Movimento = row("Movimento")
            Rom.PrimeiraPesagem = row("PrimeiraPesagem")
            Rom.SegundaPesagem = row("SegundaPesagem")
            Rom.PesoBruto = row("PesoBruto")
            Rom.Desconto = row("Desconto")
            Rom.PesoLiquido = row("PesoLiquido")
            Rom.Observacoes = row("Observacoes")
            Rom.Processo = row("Processo")
            Rom.CodigoAutorizacao = row("Autorizacao")
            Rom.TemNotaFiscal = Not (row("Nota") = 0)

            If Rom.TemNotaFiscal Then
                Nota = New NotaFiscal()
                Nota.CodigoEmpresa = row("Empresa_Id")
                Nota.EnderecoEmpresa = row("EndEmpresa_Id")
                Nota.CodigoCliente = row("Cliente_Id")
                Nota.EnderecoCliente = row("EndCliente_Id")
                Nota.Codigo = row("Nota")
                Nota.EntradaSaida = IIf(row("EntradaSaida_Id") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
                Nota.Serie = row("Serie_Id")

                Rom.NF = New NotaFiscal(Nota)
            End If

            If ds.Tables(0).Rows.Count = 1 Then Rom.Principal = True

            Me.Add(Rom)
        Next
    End Sub
#End Region

#Region "Methods"
    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each item As Romaneio In Me
            If item.IUD <> "" Then
                item.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class


'****************************************************************************************************************************************
'************************************************** CLASSE BASE ROMANEIO  ***************************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class Romaneio

#Region "Variaveis"
    Private sql As String
    Private Banco As New AcessaBanco
    Private ds As DataSet
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCodigoEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pCodigoRomaneio As Integer)

        Me.TemRateio = False

        sql = "SELECT R.Empresa_Id, R.EndEmpresa_Id, R.Romaneio_Id, R.EntradaSaida, R.Pedido, R.Deposito, R.EndDeposito, " & vbCrLf & _
              "       isnull(R.Destino,'') as Destino,  isnull(R.EndDestino,0) as EndDestino, isnull(R.Transbordo,'') as Transbordo, isnull(R.EndTransbordo,0) as EndTransbordo, R.Produto, R.Operacao," & vbCrLf & _
              "       R.SubOperacao, R.Movimento, R.PesoBruto, R.Desconto, R.PesoLiquido, isnull(R.Observacoes,'') as Observacao, isnull(R.Processo,'') as Processo, isnull(R.Autorizacao,0) as Autorizacao, " & vbCrLf & _
              "       isnull(RP.Transportador,'') as Transportador, isnull(RP.EndTransportador,0) as EndTransportador,  isnull(RP.Motorista,'') as Motorista, isnull(RP.EndMotorista,0) as EndMotorista, isnull(RP.Placa,'') as Placa, " & vbCrLf & _
              "       isnull(RP.NumeroDaNota,0) as NumeroDaNota,  isnull(RP.SerieDaNota,'') as SerieDaNota," & vbCrLf & _
              "       NXR.Cliente_id as Cliente_NF, NXR.EndCliente_Id as EndCliente_NF, isnull(NxR.Nota_Id,0) AS Nota_NF, NXR.Serie_ID as Serie_NF, NXR.EntradaSaida_id as EntradaSaida_NF," & vbCrLf & _
              "       isnull(R.PrimeiraPesagem,0) AS PrimeiraPesagem, isnull(R.SegundaPesagem,0) AS SegundaPesagem " & vbCrLf & _
              "  FROM Romaneios R" & vbCrLf & _
              "  LEFT JOIN (SELECT TOP(1) RomaneiosXPesagens.Empresa_Id, RomaneiosXPesagens.EndEmpresa_Id, RomaneiosXPesagens.Romaneio_Id," & vbCrLf & _
              "                    Pesagem.Transportador, Pesagem.EndTransportador, Pesagem.Motorista, Pesagem.EndMotorista, Pesagem.Placa,  Pesagem.NumeroDaNota,Pesagem.SerieDaNota" & vbCrLf & _
              "               FROM RomaneiosXPesagens " & vbCrLf & _
              "              INNER JOIN Pesagem " & vbCrLf & _
              "                 ON RomaneiosXPesagens.Empresa_Id    = Pesagem.Empresa_Id " & vbCrLf & _
              "                AND RomaneiosXPesagens.EndEmpresa_Id = Pesagem.EndEmpresa_Id " & vbCrLf & _
              "                AND RomaneiosXPesagens.Pesagem_Id    = Pesagem.Pesagem_Id " & vbCrLf & _
              "                AND RomaneiosXPesagens.Sequencia_Id  = Pesagem.Sequencia_Id" & vbCrLf & _
              "              Where RomaneiosXPesagens.Empresa_Id    ='" & pCodigoEmpresa & "'" & vbCrLf & _
              "                and RomaneiosXPesagens.EndEmpresa_id = " & pEndEmpresa & vbCrLf & _
              "                and RomaneiosXPesagens.Romaneio_id   = " & pCodigoRomaneio & vbCrLf & _
              "              ) RP" & vbCrLf & _
              "    ON R.Empresa_Id    = RP.Empresa_Id" & vbCrLf & _
              "   AND R.EndEmpresa_Id = RP.EndEmpresa_Id" & vbCrLf & _
              "   AND R.Romaneio_Id   = RP.Romaneio_Id" & vbCrLf & _
              "	 LEFT JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
              "	   ON NxR.Empresa_Id    = R.Empresa_id " & vbCrLf & _
              "	  AND NxR.EndEmpresa_Id = R.EndEmpresa_Id " & vbCrLf & _
              "	  AND NxR.Romaneio_Id   = R.Romaneio_Id " & vbCrLf & _
              " WHERE R.Empresa_Id    ='" & pCodigoEmpresa & "'" & vbCrLf & _
              "   AND R.EndEmpresa_id = " & pEndEmpresa & vbCrLf & _
              "   AND R.Romaneio_id   = " & pCodigoRomaneio
        ds = Banco.ConsultaDataSet(sql, "Romaneio")

        If ds.Tables(0).Rows.Count = 0 Then
            Me.CodigoEmpresa = pCodigoEmpresa
            Me.EnderecoEmpresa = pEndEmpresa
            Exit Sub
        End If

        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.CodigoEmpresa = row("Empresa_Id")
        Me.EnderecoEmpresa = row("EndEmpresa_Id")
        Me.Codigo = row("Romaneio_Id")
        Me.EntradaSaida = row("EntradaSaida")
        Me.CodigoPedido = row("Pedido")
        Me.CodigoDeposito = row("Deposito")
        Me.EnderecoDeposito = row("EndDeposito")
        Me.CodigoDestino = row("Destino")
        Me.EnderecoDestino = row("EndDestino")
        Me.CodigoTransbordo = row("Transbordo")
        Me.EnderecoTransbordo = row("EndTransbordo")
        Me.CodigoProduto = row("Produto")
        Me.CodigoOperacao = row("Operacao")
        Me.CodigoSubOperacao = row("SubOperacao")
        Me.Movimento = row("Movimento")
        Me.PrimeiraPesagem = row("PrimeiraPesagem")
        Me.SegundaPesagem = row("SegundaPesagem")
        Me.PesoBruto = row("PesoBruto")
        Me.Desconto = row("Desconto")
        Me.PesoLiquido = row("PesoLiquido")
        Me.Observacoes = row("Observacao")
        Me.Processo = row("Processo")
        If Not IsDBNull(Me.Processo) AndAlso Me.Processo.ToUpper = "RATEIO" Then
            Me.TemRateio = True
        End If
        Me.CodigoAutorizacao = row("Autorizacao")

        Me.CodigoTransportador = row("Transportador")
        Me.EnderecoTransportador = row("EndTransportador")
        Me.CodigoMotorista = row("Motorista")
        Me.EnderecoMotorista = row("EndMotorista")
        Me.Placa = row("Placa")

        If row("Nota_Nf") > 0 Then
            Dim nfconsulta As New NotaFiscal
            nfconsulta.CodigoEmpresa = row("Empresa_id")
            nfconsulta.EnderecoEmpresa = row("EndEmpresa_id")
            nfconsulta.CodigoCliente = row("Cliente_NF")
            nfconsulta.EnderecoCliente = row("EndCliente_NF")
            nfconsulta.EntradaSaida = IIf(row("EntradaSaida_NF") = "E", eEntradaSaida.Entrada, eEntradaSaida.Saida)
            nfconsulta.Numero = row("Nota_NF")
            nfconsulta.Serie = row("Serie_NF")
            Me.NF = New NotaFiscal(nfconsulta)
        End If

        Me.DescontosAnalises = New ListRomaneioXDesconto(Me)
    End Sub

    Public Sub New(ByVal pNF As NotaFiscal)
        Me.NF = pNF
        sql = "SELECT R.Empresa_Id, R.EndEmpresa_Id, R.Romaneio_Id, R.EntradaSaida, R.Pedido, R.Deposito, R.EndDeposito, " & vbCrLf & _
              "       isnull(R.Destino,'') as Destino,  isnull(R.EndDestino,0) as EndDestino," & vbCrLf & _
              "       isnull(R.Transbordo,'') as Transbordo, isnull(R.EndTransbordo,0) as EndTransbordo," & vbCrLf & _
              "       R.Produto, R.Operacao, R.SubOperacao, R.Movimento," & vbCrLf & _
              "       isnull(R.PrimeiraPesagem,0) AS PrimeiraPesagem, isnull(R.SegundaPesagem,0) AS SegundaPesagem," & vbCrLf & _
              "       R.PesoBruto, R.Desconto, R.PesoLiquido," & vbCrLf & _
              "       isnull(R.Observacoes,'') as Observacao, isnull(R.Processo,'') as Processo," & vbCrLf & _
              "       isnull(R.Autorizacao,0) as Autorizacao, " & vbCrLf & _
              "       isnull(RP.Transportador,'') as Transportador, isnull(RP.EndTransportador,0) as EndTransportador," & vbCrLf & _
              "       isnull(RP.Motorista,'') as Motorista, isnull(RP.EndMotorista,0) as EndMotorista, isnull(RP.Placa,'') as Placa, " & vbCrLf & _
              "       isnull(RP.NumeroDaNota,0) as NumeroDaNota,  isnull(RP.SerieDaNota,'') as SerieDaNota," & vbCrLf & _
              "       NXR.Cliente_id as Cliente_NF, NXR.EndCliente_Id as EndCliente_NF, isnull(NxR.Nota_Id,0) AS Nota_NF, NXR.Serie_ID as Serie_NF, NXR.EntradaSaida_id as EntradaSaida_NF" & vbCrLf & _
              "  FROM Romaneios R" & vbCrLf & _
              "  Left join (SELECT top(1) RomaneiosXPesagens.Empresa_Id, RomaneiosXPesagens.EndEmpresa_Id, RomaneiosXPesagens.Romaneio_Id," & vbCrLf & _
              "                    Pesagem.Transportador, Pesagem.EndTransportador, Pesagem.Motorista, Pesagem.EndMotorista, Pesagem.Placa,  Pesagem.NumeroDaNota,Pesagem.SerieDaNota" & vbCrLf & _
              "               FROM RomaneiosXPesagens " & vbCrLf & _
              "              INNER JOIN Pesagem " & vbCrLf & _
              "                 ON RomaneiosXPesagens.Empresa_Id    = Pesagem.Empresa_Id " & vbCrLf & _
              "                AND RomaneiosXPesagens.EndEmpresa_Id = Pesagem.EndEmpresa_Id " & vbCrLf & _
              "                AND RomaneiosXPesagens.Pesagem_Id    = Pesagem.Pesagem_Id " & vbCrLf & _
              "                AND RomaneiosXPesagens.Sequencia_Id  = Pesagem.Sequencia_Id" & vbCrLf & _
              "              Where RomaneiosXPesagens.Empresa_Id    ='" & pNF.CodigoEmpresa & "'" & vbCrLf & _
              "                and RomaneiosXPesagens.EndEmpresa_id = " & pNF.EnderecoEmpresa & vbCrLf & _
              "                and RomaneiosXPesagens.Romaneio_id   = " & pNF.CodigoRomaneio & vbCrLf & _
              "              ) RP" & vbCrLf & _
              "    on R.Empresa_Id    = RP.Empresa_Id" & vbCrLf & _
              "   and R.EndEmpresa_Id = RP.EndEmpresa_Id" & vbCrLf & _
              "   and R.Romaneio_Id   = RP.Romaneio_Id" & vbCrLf & _
              "	 LEFT JOIN NotasFiscaisXRomaneios NxR " & vbCrLf & _
              "	   ON NxR.Empresa_Id    = R.Empresa_id " & vbCrLf & _
              "	  AND NxR.EndEmpresa_Id = R.EndEmpresa_Id " & vbCrLf & _
              "	  AND NxR.Romaneio_Id   = R.Romaneio_Id " & vbCrLf & _
              " Where R.Empresa_Id    ='" & pNF.CodigoEmpresa & "'" & vbCrLf & _
              "   and R.EndEmpresa_id = " & pNF.EnderecoEmpresa & vbCrLf & _
              "   and R.Romaneio_id   = " & pNF.CodigoRomaneio
        ds = Banco.ConsultaDataSet(sql, "Romaneio")

        If ds.Tables(0).Rows.Count = 0 Then
            Me.CodigoEmpresa = pNF.CodigoEmpresa
            Me.EnderecoEmpresa = pNF.EnderecoEmpresa
            Me.EntradaSaida = pNF.SubOperacao.EntradaSaida
        Else
            Dim row As DataRow = ds.Tables(0).Rows(0)

            Me.CodigoEmpresa = row("Empresa_Id")
            Me.EnderecoEmpresa = row("EndEmpresa_Id")
            Me.Codigo = row("Romaneio_Id")
            Me.EntradaSaida = row("EntradaSaida")
            Me.CodigoPedido = row("Pedido")
            Me.CodigoDeposito = row("Deposito")
            Me.EnderecoDeposito = row("EndDeposito")
            Me.CodigoDestino = row("Destino")
            Me.EnderecoDestino = row("EndDestino")
            Me.CodigoTransbordo = row("Transbordo")
            Me.EnderecoTransbordo = row("EndTransbordo")
            Me.CodigoProduto = row("Produto")
            Me.CodigoOperacao = row("Operacao")
            Me.CodigoSubOperacao = row("SubOperacao")
            Me.Movimento = row("Movimento")
            Me.PrimeiraPesagem = row("PrimeiraPesagem")
            Me.SegundaPesagem = row("SegundaPesagem")
            Me.PesoBruto = row("PesoBruto")
            Me.Desconto = row("Desconto")
            Me.PesoLiquido = row("PesoLiquido")
            Me.Observacoes = row("Observacao")
            Me.Processo = row("Processo")

            If Not IsDBNull(Me.Processo) AndAlso Me.Processo.ToUpper = "RATEIO" Then
                Me.TemRateio = True
            End If

            Me.CodigoAutorizacao = row("Autorizacao")

            Me.CodigoTransportador = row("Transportador")
            Me.EnderecoTransportador = row("EndTransportador")
            Me.CodigoMotorista = row("Motorista")
            Me.EnderecoMotorista = row("EndMotorista")
            Me.Placa = row("Placa")
        End If
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String
    Private _CodigoEmpresa As String
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente
    Private _Codigo As Integer
    Private _EntradaSaida As String
    Private _CodigoPedido As Integer
    Private _Pedido As Pedido
    Private _CodigoDeposito As String
    Private _EnderecoDeposito As Integer
    Private _Deposito As Cliente
    Private _CodigoDestino As String
    Private _EnderecoDestino As Integer
    Private _Destino As Cliente
    Private _CodigoTransbordo As String
    Private _EnderecoTransbordo As Integer
    Private _Transbordo As Cliente
    Private _CodigoProduto As String
    Private _Produto As Produto
    Private _CodigoOperacao As Integer
    Private _Operacao As Operacao
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao
    Private _Movimento As DateTime
    Private _PrimeiraPesagem As Double
    Private _SegundaPesagem As Double
    Private _PesoBruto As Double
    Private _Desconto As Double
    Private _PesoLiquido As Double
    Private _Observacoes As String
    Private _Processo As String

    Private _CodigoTransportador As String = ""
    Private _EnderecoTransportador As Integer
    Private _Transportador As Cliente

    Private _CodigoMotorista As String = ""
    Private _EnderecoMotorista As Integer
    Private _Motorista As Cliente

    Private _Placa As String

    Private _DescontosAnalises As ListRomaneioXDesconto

    '*********** NOTA FISCAL *************
    Private _NF As NotaFiscal

    'Autorizacao de Retirada
    Private _Autorizacao As AutorizacaoDeRetirada
    Private _CodigoAutorizacao As Integer = 0

    'Romaneio x Pesagem
    Private _Pesagens As ListRomaneioXPesagem
    Private _TemRateio As Boolean
    Private _TemAgrupamento As Boolean
    Private _TemNotaFiscal As Boolean

    Private _Principal As Boolean

#End Region

#Region "Property"
    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

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

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And Not Me.CodigoEmpresa Is Nothing Then _Empresa = New Cliente(Me.CodigoEmpresa, Me.EnderecoEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
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

    Public Property EntradaSaida() As String
        Get
            Return _EntradaSaida
        End Get
        Set(ByVal value As String)
            _EntradaSaida = value
        End Set
    End Property

    Public Property CodigoPedido() As Integer
        Get
            Return _CodigoPedido
        End Get
        Set(ByVal value As Integer)
            _CodigoPedido = value
        End Set
    End Property

    Public Property Pedido() As Pedido
        Get
            If _Pedido Is Nothing And Me.CodigoPedido > 0 Then _Pedido = New Pedido(Me.CodigoEmpresa, Me.EnderecoEmpresa, Me.CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
        End Set
    End Property

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

    Public Property Deposito() As Cliente
        Get
            If _Deposito Is Nothing And Not Me.CodigoDeposito Is Nothing Then _Deposito = New Cliente(Me.CodigoDeposito, Me.EnderecoDeposito)
            Return _Deposito
        End Get
        Set(ByVal value As Cliente)
            _Deposito = value
        End Set
    End Property

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

    Public Property Destino() As Cliente
        Get
            If _Destino Is Nothing And Not Me.CodigoDestino Is Nothing Then _Destino = New Cliente(Me.CodigoDestino, Me.EnderecoDestino)
            Return _Destino
        End Get
        Set(ByVal value As Cliente)
            _Destino = value
        End Set
    End Property

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

    Public Property Transbordo() As Cliente
        Get
            If _Transbordo Is Nothing And Not Me.CodigoTransbordo Is Nothing Then _Transbordo = New Cliente(Me.CodigoTransbordo, Me.EnderecoTransbordo)
            Return _Transbordo
        End Get
        Set(ByVal value As Cliente)
            _Transbordo = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And Not Me.CodigoProduto Is Nothing Then _Produto = New Produto(Me.CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property CodigoOperacao() As Integer
        Get
            Return _CodigoOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoOperacao = value
        End Set
    End Property

    Public Property Operacao() As Operacao
        Get
            If _Operacao Is Nothing And Me.CodigoOperacao > 0 Then _Operacao = New Operacao(Me.CodigoOperacao)
            Return _Operacao
        End Get
        Set(ByVal value As Operacao)
            _Operacao = value
        End Set
    End Property

    Public Property CodigoSubOperacao() As Integer
        Get
            Return _CodigoSubOperacao
        End Get
        Set(ByVal value As Integer)
            _CodigoSubOperacao = value
        End Set
    End Property

    Public Property SubOperacao() As SubOperacao
        Get
            If _SubOperacao Is Nothing And Me.CodigoSubOperacao > 0 And Me.CodigoOperacao > 0 Then _SubOperacao = New SubOperacao(Me.CodigoOperacao, Me.CodigoSubOperacao)
            Return _SubOperacao
        End Get
        Set(ByVal value As SubOperacao)
            _SubOperacao = value
        End Set
    End Property

    Public Property Movimento() As DateTime
        Get
            Return _Movimento
        End Get
        Set(ByVal value As DateTime)
            _Movimento = value
        End Set
    End Property

    Public Property PrimeiraPesagem() As Double
        Get
            Return _PrimeiraPesagem
        End Get
        Set(ByVal value As Double)
            _PrimeiraPesagem = value
        End Set
    End Property

    Public Property SegundaPesagem() As Double
        Get
            Return _SegundaPesagem
        End Get
        Set(ByVal value As Double)
            _SegundaPesagem = value
        End Set
    End Property

    Public Property PesoBruto() As Double
        Get
            Return _PesoBruto
        End Get
        Set(ByVal value As Double)
            _PesoBruto = value
        End Set
    End Property

    Public Property Desconto() As Double
        Get
            Return _Desconto
        End Get
        Set(ByVal value As Double)
            _Desconto = value
        End Set
    End Property

    Public Property PesoLiquido() As Double
        Get
            Return _PesoLiquido
        End Get
        Set(ByVal value As Double)
            _PesoLiquido = value
        End Set
    End Property

    Public Property Observacoes() As String
        Get
            Return _Observacoes
        End Get
        Set(ByVal value As String)
            _Observacoes = value
        End Set
    End Property

    Public Property Processo() As String
        Get
            Return _Processo
        End Get
        Set(ByVal value As String)
            _Processo = value
        End Set
    End Property

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

    Public Property Transportador() As Cliente
        Get
            If _Transportador Is Nothing And Me.CodigoTransportador.Length > 0 Then _Transportador = New Cliente(Me.CodigoTransportador, Me.EnderecoTransportador)
            Return _Transportador
        End Get
        Set(ByVal value As Cliente)
            _Transportador = value
        End Set
    End Property

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

    Public Property Motorista() As Cliente
        Get
            If _Motorista Is Nothing And Me.CodigoMotorista.Length > 0 Then _Motorista = New Cliente(Me.CodigoMotorista, Me.EnderecoMotorista)
            Return _Motorista
        End Get
        Set(ByVal value As Cliente)
            _Motorista = value
        End Set
    End Property

    Public Property Placa() As String
        Get
            Return _Placa
        End Get
        Set(ByVal value As String)
            _Placa = value
        End Set
    End Property

    '************  Autorizacao de Retirada  ***********
    Public Property Autorizacao() As AutorizacaoDeRetirada
        Get
            If _Autorizacao Is Nothing And _CodigoAutorizacao > 0 Then _Autorizacao = New AutorizacaoDeRetirada(_CodigoEmpresa, _EnderecoEmpresa, Me.CodigoPedido, _CodigoAutorizacao, _SubOperacao.Classe)
            Return _Autorizacao
        End Get
        Set(ByVal value As AutorizacaoDeRetirada)
            _Autorizacao = value
        End Set
    End Property

    Public Property CodigoAutorizacao() As Integer
        Get
            Return _CodigoAutorizacao
        End Get
        Set(ByVal value As Integer)
            _CodigoAutorizacao = value
        End Set
    End Property

    '************* Analises - Descontos ***************
    Public Property DescontosAnalises() As ListRomaneioXDesconto
        Get
            If _DescontosAnalises Is Nothing Then _DescontosAnalises = New ListRomaneioXDesconto(Me)
            Return _DescontosAnalises
        End Get
        Set(ByVal value As ListRomaneioXDesconto)
            _DescontosAnalises = value
        End Set
    End Property

    '************** Pesagem - RomaneioxPesagem ********
    Public Property Pesagens() As ListRomaneioXPesagem
        Get
            If _Pesagens Is Nothing AndAlso _CodigoEmpresa IsNot Nothing AndAlso _CodigoEmpresa.Length > 0 AndAlso _Codigo > 0 Then _Pesagens = New ListRomaneioXPesagem(_CodigoEmpresa, _EnderecoEmpresa, _Codigo, 0)
            Return _Pesagens
        End Get
        Set(ByVal value As ListRomaneioXPesagem)
            _Pesagens = value
        End Set
    End Property

    Public Property TemRateio() As Boolean
        Get
            Return _TemRateio
        End Get
        Set(ByVal value As Boolean)
            _TemRateio = value
        End Set
    End Property

    Public Property TemAgrupamento() As Boolean
        Get
            Return _TemAgrupamento
        End Get
        Set(ByVal value As Boolean)
            _TemAgrupamento = value
        End Set
    End Property

    '**************  Nota Fiscal  ***********
    Public Property NF As NotaFiscal
        Get
            Return _NF
        End Get
        Set(value As NotaFiscal)
            _NF = value
        End Set
    End Property

    Public Property TemNotaFiscal() As Boolean
        Get
            Return _TemNotaFiscal
        End Get
        Set(ByVal value As Boolean)
            _TemNotaFiscal = value
        End Set
    End Property

    Public Property Principal() As Boolean
        Get
            Return _Principal
        End Get
        Set(ByVal value As Boolean)
            _Principal = value
        End Set
    End Property

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        If IUD = Nothing Then Return True
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)
        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList, Optional ByVal AtualizaNumerador As Boolean = True)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                If AtualizaNumerador Then
                    Dim N As New Numerador(_CodigoEmpresa, _EnderecoEmpresa, 110)
                    Sqls.Add(N.IncrementarNumeradorSql)
                    _Codigo = N.Sequencia + 1
                End If

                sql = "INSERT INTO Romaneios (Empresa_Id, EndEmpresa_Id, Romaneio_Id, EntradaSaida, Pedido, Deposito, EndDeposito," & vbCrLf & _
                      "  Destino, EndDestino, Transbordo, EndTransbordo, Produto, Operacao, SubOperacao, Movimento," & vbCrLf & _
                      "  PrimeiraPesagem, SegundaPesagem, PesoBruto, Desconto, PesoLiquido, Observacoes, Processo, Autorizacao)" & vbCrLf & _
                      " VALUES ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & ",'" & Me.EntradaSaida & "'," & vbCrLf & _
                      Me.CodigoPedido & ",'" & Me.CodigoDeposito & "'," & Me.EnderecoDeposito & "," & vbCrLf & _
                      "'" & Me.CodigoDestino & "', " & Me.EnderecoDestino & ",'" & Me.CodigoTransbordo & "'," & Me.EnderecoTransbordo & ",'" & Me.CodigoProduto & "'," & Me.CodigoOperacao & "," & Me.CodigoSubOperacao & ",CONVERT(DATETIME,'" & _Movimento.ToString("yyyy-MM-dd") & "', 102), " & vbCrLf & _
                      Str(Me.PrimeiraPesagem) & "," & Str(Me.SegundaPesagem) & "," & Str(Me.PesoBruto) & "," & Str(Me.Desconto) & "," & Str(Me.PesoLiquido) & ",'" & Me.Observacoes & "','" & Me.Processo & "', " & Str(Me.CodigoAutorizacao) & ")" & vbCrLf
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "Update Romaneios set " & vbCrLf & _
                      "   EntradaSaida    ='" & Me.EntradaSaida & "'" & vbCrLf & _
                      "  ,Pedido          = " & Me.CodigoPedido & vbCrLf & _
                      "  ,Deposito        ='" & Me.CodigoDeposito & "'" & vbCrLf & _
                      "  ,EndDeposito     = " & Me.EnderecoDeposito & vbCrLf & _
                      "  ,Destino         ='" & Me.CodigoDestino & "'" & vbCrLf & _
                      "  ,EndDestino      = " & Me.EnderecoDestino & vbCrLf & _
                      "  ,Transbordo      ='" & Me.CodigoTransbordo & "'" & vbCrLf & _
                      "  ,EndTransbordo   = " & Me.EnderecoTransbordo & vbCrLf & _
                      "  ,Produto         ='" & Me.CodigoProduto & "'" & vbCrLf & _
                      "  ,Operacao        = " & Me.CodigoOperacao & vbCrLf & _
                      "  ,SubOperacao     = " & Me.CodigoSubOperacao & vbCrLf & _
                      "  ,Movimento       ='" & Me.Movimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "  ,PrimeiraPesagem = " & Str(Me.PrimeiraPesagem) & vbCrLf & _
                      "  ,SegundaPesagem  = " & Str(Me.SegundaPesagem) & vbCrLf & _
                      "  ,PesoBruto       = " & Str(Me.PesoBruto) & vbCrLf & _
                      "  ,Desconto        = " & Str(Me.Desconto) & vbCrLf & _
                      "  ,PesoLiquido     = " & Str(Me.PesoLiquido) & vbCrLf & _
                      "  ,Observacoes     ='" & Me.Observacoes & "'" & vbCrLf & _
                      "  ,Processo        ='" & Me.Processo & "'" & vbCrLf & _
                      "  ,Autorizacao     = " & Str(Me.CodigoAutorizacao) & vbCrLf & _
                      " Where Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_id = " & Me.EnderecoEmpresa & vbCrLf & _
                      "   and Romaneio_Id   = " & Me.Codigo & vbCrLf
                Sqls.Add(sql)
                SalvarTabelasRelacionadasSql(Sqls)
            Case "D"
                If (Me.Processo = "Rateio" OrElse Me.Processo = "AGRUPAMENTO" OrElse Me.Pesagens.Count = 0) Then
                    SalvarTabelasRelacionadasSql(Sqls)

                    sql = " Delete Romaneios " & vbCrLf & _
                          "  Where Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                          "    and EndEmpresa_id = " & Me.EnderecoEmpresa & vbCrLf & _
                          "    and Romaneio_Id   = " & Me.Codigo & vbCrLf
                    Sqls.Add(sql)
                End If
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        DescontosAnalises.SalvarSql(Sqls)

        If Pesagens IsNot Nothing Then
            For Each P As RomaneioXPesagem In Pesagens
                P.IUD = IUD
                P.SalvarSql(Sqls)
            Next
        End If
    End Sub
#End Region

End Class