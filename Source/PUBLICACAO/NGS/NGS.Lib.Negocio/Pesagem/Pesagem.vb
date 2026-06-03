Imports Microsoft.VisualBasic
Imports System.Data
Imports System.Collections.Generic
Imports NGS.Lib.Uteis
Imports System.Web

'************************************************************************************************************************************************************
'*********************************************************** Lista da Classe Pesagem ************************************************************************
'************************************************************************************************************************************************************
<Serializable()> _
Public Class ListPesagem
    Inherits List(Of Pesagem)

#Region "Contrutor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pPesagem As Pesagem, _
                   Optional ByVal SemNota As Boolean = True, _
                   Optional ByVal DataInicial As String = "", _
                   Optional ByVal DataFinal As String = "", _
                   Optional ByVal ProcessoLaudo As String = "", _
                   Optional ByVal SegundaPesagem As Boolean = False, _
                   Optional ByVal RegistroMestre As Boolean = False)

        Dim Banco As New AcessaBanco()
        Dim Sql As String
        Dim strAndWhere As String = "WHERE"

        Try
            Sql = "SELECT P.Empresa_Id, P.EndEmpresa_Id, P.Pesagem_Id, P.Sequencia_Id," & vbCrLf & _
                  "       P.Cliente, P.EndCliente," & vbCrLf & _
                  "       P.Pedido, P.Produto, P.Situacao," & vbCrLf & _
                  "       P.Depositante, P.EndDepositante," & vbCrLf & _
                  "       P.Transportador, P.EndTransportador," & vbCrLf & _
                  "       P.Motorista, P.EndMotorista, P.Placa," & vbCrLf & _
                  "       P.Deposito, P.EndDeposito, P.TabelaDeClassificacao," & vbCrLf & _
                  "       P.Operacao, P.SubOperacao, P.EntradaSaida," & vbCrLf & _
                  "       P.PrimeiraPesagem, P.SegundaPesagem, P.BrutoBalanca, P.Liquido," & vbCrLf & _
                  "       P.ViaTransporte, P.Movimento, P.PesagemDeTerceiros, P.PesagemMecanica," & vbCrLf & _
                  "       P.NumeroDaNota, P.SerieDaNota, isnull(P.PesoFiscal,0) AS PesoFiscal," & vbCrLf & _
                  "       P.Observacoes, isnull(P.RegistroMestre,0) as RegistroMestre," & vbCrLf & _
                  "       P.EntradaPatio, P.EntradaBalanca," & vbCrLf & _
                  "	      isnull(P.SaidaBalanca,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS SaidaBalanca," & vbCrLf & _
                  "       P.UsuarioInclusao," & vbCrLf & _
                  "       isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao," & vbCrLf & _
                  "       isnull(P.UsuarioAlteracaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioAlteracaoData," & vbCrLf & _
                  "       isnull(P.UsuarioReimpressao,'') AS UsuarioReimpressao," & vbCrLf & _
                  "       isnull(P.UsuarioReimpressaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioReimpressaoData," & vbCrLf & _
                  "       isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento," & vbCrLf & _
                  "       isnull(P.UsuarioCancelamentoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioCancelamentoData," & vbCrLf & _
                  "	      case when Mov.QtdeRomaneio = 0         then 0 else 1 end TemRomaneio," & vbCrLf & _
                  "	      case when Mov.QtdeNota     = 0         then 0 else 1 end TemNota," & vbCrLf & _
                  "	      case when Mov.QtdePesagem  > 1         then 1 else 0 end TemAgrupamento," & vbCrLf & _
                  "	      case when Mov.QtdeRomaneio > 1         then 1 else 0 end TemRateio," & vbCrLf & _
                  "	      case when Mov.QtdeRomaneio <> QtdeNota then 1 else 0 end TemRomaneioSemNota" & vbCrLf & _
                  "  FROM Pesagem P" & vbCrLf & _
                  " left Join(Select RxP.Empresa_Id, RxP.EndEmpresa_Id, RxP.Pesagem_Id, RxP.sequencia_id," & vbCrLf & _
                  "                  case when isnull(agrup.QtdePesagem,0) > 0 then agrup.QtdePesagem else Rateio.QtdePesagem end as QtdePesagem," & vbCrLf & _
                  "                   count(*) QtdeRomaneio," & vbCrLf & _
                  "				   sum(case" & vbCrLf & _
                  "				         when NxR.Nota_Id is null" & vbCrLf & _
                  "						   then 0" & vbCrLf & _
                  "						   else 1" & vbCrLf & _
                  "                        end) QtdeNota" & vbCrLf & _
                  "              from Pesagem P2" & vbCrLf & _
                  "			 inner join RomaneiosXPesagens RxP" & vbCrLf & _
                  "			    on rxp.Empresa_Id    = P2.Empresa_Id" & vbCrLf & _
                  "               and rxp.EndEmpresa_Id = P2.EndEmpresa_Id" & vbCrLf & _
                  "			   and rxp.pesagem_id    = p2.Pesagem_Id" & vbCrLf & _
                  "              Left Join NotasFiscaisXRomaneios NxR" & vbCrLf & _
                  "			    on rxp.Empresa_Id    = nxr.Empresa_Id" & vbCrLf & _
                  "			   and rxp.EndEmpresa_Id = nxr.EndEmpresa_Id" & vbCrLf & _
                  "			   and rxp.Romaneio_Id   = nxr.Romaneio_Id" & vbCrLf & _
                  "              Left join (select p2.Empresa_Id, p2.EndEmpresa_Id, p2.RegistroMestre, count(*) as QtdePesagem" & vbCrLf & _
                  "			               from Pesagem p2" & vbCrLf & _
                  "						  Group by p2.Empresa_Id, p2.EndEmpresa_Id, p2.RegistroMestre" & vbCrLf & _
                  "			            ) Agrup" & vbCrLf & _
                  "                ON P2.Empresa_Id    = Agrup.Empresa_Id" & vbCrLf & _
                  "               AND P2.EndEmpresa_Id = Agrup.EndEmpresa_Id" & vbCrLf & _
                  "               AND P2.Pesagem_Id    = Agrup.RegistroMestre" & vbCrLf & _
                  "              Left join (select p2.Empresa_Id, p2.EndEmpresa_Id, p2.Pesagem_Id, count(*) as QtdePesagem" & vbCrLf & _
                  "			               from Pesagem p2" & vbCrLf & _
                  "						  Group by p2.Empresa_Id, p2.EndEmpresa_Id, p2.Pesagem_Id" & vbCrLf & _
                  "			            ) Rateio" & vbCrLf & _
                  "                ON P2.Empresa_Id    = Rateio.Empresa_Id" & vbCrLf & _
                  "               AND P2.EndEmpresa_Id = Rateio.EndEmpresa_Id" & vbCrLf & _
                  "               AND P2.Pesagem_Id    = Rateio.Pesagem_Id" & vbCrLf & _
                  "             group by RxP.Empresa_Id, RxP.EndEmpresa_Id, RxP.Pesagem_Id, RxP.Sequencia_id," & vbCrLf & _
                  "				         case when isnull(agrup.QtdePesagem,0) > 0 then agrup.QtdePesagem else Rateio.QtdePesagem end" & vbCrLf & _
                  "           ) as Mov" & vbCrLf & _
                  "    ON Mov.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
                  "   AND Mov.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                  "   AND Mov.Pesagem_Id    = case" & vbCrLf & _
                  "			                    when isnull(P.RegistroMestre,0) > 0" & vbCrLf & _
                  "							      then P.RegistroMestre" & vbCrLf & _
                  "							      else P.Pesagem_Id" & vbCrLf & _
                  "							  end" & vbCrLf & _
                  "   AND Mov.Sequencia_Id  = case" & vbCrLf & _
                  "			                    when isnull(p.RegistroMestre,0) > 0" & vbCrLf & _
                  "							      then 0" & vbCrLf & _
                  "							      else p.Sequencia_Id" & vbCrLf & _
                  "							  end" & vbCrLf


            If pPesagem.CodigoEmpresa.Length > 0 Then
                Sql &= strAndWhere & " P.Empresa_Id = '" & pPesagem.CodigoEmpresa & "' AND P.EndEmpresa_id = " & pPesagem.EnderecoEmpresa & " " & vbCrLf
                strAndWhere = " AND "
            End If

            If RegistroMestre Then
                Sql &= strAndWhere & " isnull(P.RegistroMestre,0) = " & pPesagem.Codigo & vbCrLf
                strAndWhere = " AND "
            Else
                If pPesagem.Codigo > 0 Then
                    Sql &= strAndWhere & " P.Pesagem_Id = " & pPesagem.Codigo & vbCrLf
                    strAndWhere = " AND "
                Else
                    If Not String.IsNullOrWhiteSpace(DataInicial) AndAlso Not String.IsNullOrWhiteSpace(DataFinal) Then
                        Sql &= strAndWhere & " P.Movimento BETWEEN '" & CDate(DataInicial).ToString("yyyy-MM-dd") & "' AND '" & CDate(DataFinal).ToString("yyyy-MM-dd") & "'" & vbCrLf
                    End If
                End If

                If pPesagem.CodigoPedido > 0 Then
                    Sql &= strAndWhere & " P.Pedido = " & pPesagem.CodigoPedido & " " & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoCliente.Length > 0 Then
                    Sql &= strAndWhere & " P.Cliente = '" & pPesagem.CodigoCliente & "' AND P.EndCliente = " & pPesagem.EnderecoCliente & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoDeposito.Length > 0 Then
                    Sql &= strAndWhere & " P.Deposito = '" & pPesagem.CodigoDeposito & "' AND P.EndDeposito = " & pPesagem.EnderecoDeposito & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoDepositante.Length > 0 Then
                    Sql &= strAndWhere & " P.Depositante = '" & pPesagem.CodigoDepositante & "' AND P.EndDepositante = " & pPesagem.EnderecoDepositante & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoTransportador.Length > 0 Then
                    Sql &= strAndWhere & " P.Transportador = '" & pPesagem.CodigoTransportador & "' AND P.EndTransportador = " & pPesagem.EnderecoTransportador & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoMotorista.Length > 0 Then
                    Sql &= strAndWhere & " P.Motorista = '" & pPesagem.CodigoMotorista & "' AND P.EndMotorista = " & pPesagem.EnderecoMotorista & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoSituacao > 0 Then
                    Sql &= strAndWhere & " P.Situacao = " & pPesagem.CodigoSituacao & vbCrLf
                    strAndWhere = " AND "
                End If

                If pPesagem.CodigoPlaca.Length > 0 Then
                    Sql &= strAndWhere & " P.Placa = '" & pPesagem.CodigoPlaca & "' " & vbCrLf
                    strAndWhere = " AND "
                End If

                If SegundaPesagem Then
                    Sql &= strAndWhere & " isnull(P.SegundaPesagem,0) > 0 " & vbCrLf
                    strAndWhere = " AND "
                End If
            End If

            If SemNota Then
                Sql &= strAndWhere & " Mov.QtdeNota     = 0  " & vbCrLf
                strAndWhere = " AND "
            End If

            If Not String.IsNullOrEmpty(ProcessoLaudo) Then
                If ProcessoLaudo.ToUpper = "AGRUPAMENTO" Then
                    Sql &= " AND Mov.QtdePesagem  > 1" & vbCrLf
                ElseIf ProcessoLaudo.ToUpper = "RATEIO" Then
                    Sql &= " AND Mov.QtdeRomaneio > 1" & vbCrLf
                End If
            End If

            Sql &= " AND P.Sequencia_Id = 0 " & vbCrLf

            Sql &= " ORDER BY P.Pesagem_Id ASC"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pesagem")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim P As New Pesagem
                P.SelecionaPesagem = False

                P.CodigoEmpresa = row("Empresa_Id")
                P.EnderecoEmpresa = row("EndEmpresa_Id")
                P.Codigo = row("Pesagem_Id")
                P.Sequencia = row("Sequencia_Id")
                P.CodigoPedido = row("Pedido")
                P.CodigoProduto = row("Produto")
                P.CodigoCliente = row("Cliente")
                P.EnderecoCliente = row("EndCliente")
                P.CodigoDepositante = row("Depositante")
                P.EnderecoDepositante = row("EndDepositante")
                P.CodigoTransportador = row("Transportador")
                P.EnderecoTransportador = row("EndTransportador")
                P.CodigoMotorista = row("Motorista")
                P.EnderecoMotorista = row("EndMotorista")
                P.CodigoDeposito = row("Deposito")
                P.EnderecoDeposito = row("EndDeposito")
                P.CodigoTabelaDeClassificacao = row("TabelaDeClassificacao")
                P.CodigoPlaca = row("Placa")
                P.CodigoSituacao = row("Situacao")
                P.CodigoOperacao = row("Operacao")
                P.CodigoSubOperacao = row("SubOperacao")
                P.EntradaSaida = row("EntradaSaida")
                P.PrimeiraPesagem = row("PrimeiraPesagem")
                P.SegundaPesagem = row("SegundaPesagem")
                P.BrutoBalanca = row("BrutoBalanca")
                P.Desconto = row("BrutoBalanca") - row("Liquido")
                P.Liquido = row("Liquido")
                P.EntradaPatio = row("EntradaPatio")
                P.EntradaBalanca = row("EntradaBalanca")
                If Not IsDBNull(row("SaidaBalanca")) Then P.SaidaBalanca = row("SaidaBalanca")
                P.CodigoViaDeTransporte = row("ViaTransporte")
                P.Movimento = row("Movimento")
                P.PesagemDeTerceiros = row("PesagemDeTerceiros")
                P.PesagemMecanica = row("PesagemMecanica")
                P.NumeroDaNota = row("NumeroDaNota")
                P.SerieDaNota = row("SerieDaNota")
                P.PesoFiscal = row("PesoFiscal")
                P.UsuarioInclusao = row("UsuarioInclusao")
                P.UsuarioAlteracao = row("UsuarioAlteracao")
                P.DataAlteracao = DateTime.Now
                P.UsuarioReimpressao = row("UsuarioReimpressao")
                P.DataReimpressao = DateTime.Now
                P.UsuarioCancelamento = row("UsuarioCancelamento")
                P.DataCancelamento = row("UsuarioCancelamentoData")
                P.Observacoes = row("Observacoes").ToString()

                P.RegistroMestre = row("RegistroMestre")

                'If IsDBNull(row("Autorizacao")) Then
                '    P.CodigoAutorizacao = 0
                'Else
                '    P.CodigoAutorizacao = CInt(row("Autorizacao").ToString())
                'End If

                P.CodigoAutorizacao = 0

                P.TemRomaneio = row("TemRomaneio")
                P.TemNota = row("TemNota")
                P.TemAgrupamento = row("TemAgrupamento")
                P.TemRateio = row("TemRateio")
                P.TemRomaneioSemNota = row("TemRomaneioSemNota")
                Me.Add(P)
            Next
        Catch ex As Exception

        Finally
            Banco = Nothing
        End Try
    End Sub

#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList

        Sqls.Clear()
        SalvarSql(Sqls)

        Return Banco.GravaBanco(Sqls)
    End Function

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each P As Pesagem In Me
            If P.IUD <> "" Then
                P.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

'************************************************************************************************************************************************************
'*************************************************************** Classe Base Pesagem ************************************************************************
'************************************************************************************************************************************************************
<Serializable()> _
Public Class Pesagem

#Region "Construtores"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pEmpresa As String, ByVal pEndEmpresa As Integer, ByVal pPesagem As Integer, Optional ByVal pSequencia As Integer = 0)
        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT P.Empresa_Id, P.EndEmpresa_Id, P.Pesagem_Id, P.Sequencia_Id," & vbCrLf & _
                  "       P.Cliente, P.EndCliente," & vbCrLf & _
                  "       P.Pedido, P.Produto, P.Situacao," & vbCrLf & _
                  "       P.Depositante, P.EndDepositante," & vbCrLf & _
                  "       P.Transportador, P.EndTransportador," & vbCrLf & _
                  "       P.Motorista, P.EndMotorista, P.Placa," & vbCrLf & _
                  "       P.Deposito, P.EndDeposito, P.TabelaDeClassificacao," & vbCrLf & _
                  "       P.Operacao, P.SubOperacao, P.EntradaSaida," & vbCrLf & _
                  "       P.PrimeiraPesagem, P.SegundaPesagem, P.BrutoBalanca, P.Liquido," & vbCrLf & _
                  "       P.ViaTransporte, P.Movimento, P.PesagemDeTerceiros, P.PesagemMecanica," & vbCrLf & _
                  "       P.NumeroDaNota, P.SerieDaNota, isnull(P.PesoFiscal,0) AS PesoFiscal," & vbCrLf & _
                  "       P.Observacoes, isnull(P.RegistroMestre,0) as RegistroMestre," & vbCrLf & _
                  "       P.EntradaPatio, P.EntradaBalanca," & vbCrLf & _
                  "	      isnull(P.SaidaBalanca,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS SaidaBalanca," & vbCrLf & _
                  "       P.UsuarioInclusao," & vbCrLf & _
                  "       isnull(P.UsuarioAlteracao,'') AS UsuarioAlteracao," & vbCrLf & _
                  "       isnull(P.UsuarioAlteracaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioAlteracaoData," & vbCrLf & _
                  "       isnull(P.UsuarioReimpressao,'') AS UsuarioReimpressao," & vbCrLf & _
                  "       isnull(P.UsuarioReimpressaoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioReimpressaoData," & vbCrLf & _
                  "       isnull(P.UsuarioCancelamento,'') AS UsuarioCancelamento," & vbCrLf & _
                  "       isnull(P.UsuarioCancelamentoData,CONVERT(varchar, CURRENT_TIMESTAMP, 112)) AS UsuarioCancelamentoData," & vbCrLf & _
                  "	      case when isnull(QtdeRomaneio,0) = 0            then 0 else 1 end TemRomaneio," & vbCrLf & _
                  "	      case when isnull(QtdeNota,0)     = 0            then 0 else 1 end TemNota," & vbCrLf & _
                  "	      case when isnull(QtdePesagem,0)  > 1            then 1 else 0 end TemAgrupamento," & vbCrLf & _
                  "	      case when isnull(QtdeRomaneio,0) > 1            then 1 else 0 end TemRateio," & vbCrLf & _
                  "	      case when isnull(QtdeRomaneio,0) > 0 and isnull(QtdeRomaneio,0) <> isnull(QtdeNota,0) then 1 else 0 end TemRomaneioSemNota" & vbCrLf & _
                  "  FROM Pesagem P" & vbCrLf & _
                  "  left Join(Select P2.Empresa_Id, P2.EndEmpresa_Id, P2.Pesagem_Id, P2.sequencia_id," & vbCrLf & _
                  "                   case when isnull(agrup.QtdePesagem,0) > 0 then agrup.QtdePesagem else Rateio.QtdePesagem end as QtdePesagem," & vbCrLf & _
                  "                   count(*) QtdeRomaneio," & vbCrLf & _
                  "				      sum(case" & vbCrLf & _
                  "				            when NxR.Nota_Id is null" & vbCrLf & _
                  "						      then 0" & vbCrLf & _
                  "						      else 1" & vbCrLf & _
                  "                       end) QtdeNota" & vbCrLf & _
                  "             from Pesagem P2" & vbCrLf & _
                  "			   inner join RomaneiosXPesagens RxP" & vbCrLf & _
                  "			      on rxp.Empresa_Id    = P2.Empresa_Id" & vbCrLf & _
                  "              and rxp.EndEmpresa_Id = P2.EndEmpresa_Id" & vbCrLf & _
                  "			     and rxp.pesagem_id    = p2.Pesagem_Id" & vbCrLf & _
                  "             Left Join NotasFiscaisXRomaneios NxR" & vbCrLf & _
                  "			      on rxp.Empresa_Id    = nxr.Empresa_Id" & vbCrLf & _
                  "			     and rxp.EndEmpresa_Id = nxr.EndEmpresa_Id" & vbCrLf & _
                  "			     and rxp.Romaneio_Id   = nxr.Romaneio_Id" & vbCrLf & _
                  "             Left join (select p2.Empresa_Id, p2.EndEmpresa_Id, p2.RegistroMestre, count(*) as QtdePesagem" & vbCrLf & _
                  "			                 from Pesagem p2" & vbCrLf & _
                  "						    Group by p2.Empresa_Id, p2.EndEmpresa_Id, p2.RegistroMestre" & vbCrLf & _
                  "			              ) Agrup" & vbCrLf & _
                  "                ON P2.Empresa_Id    = Agrup.Empresa_Id" & vbCrLf & _
                  "               AND P2.EndEmpresa_Id = Agrup.EndEmpresa_Id" & vbCrLf & _
                  "               AND P2.Pesagem_Id    = Agrup.RegistroMestre" & vbCrLf & _
                  "              Left join (select p2.Empresa_Id, p2.EndEmpresa_Id, p2.Pesagem_Id, count(*) as QtdePesagem" & vbCrLf & _
                  "			                  from Pesagem p2" & vbCrLf & _
                  "						     Group by p2.Empresa_Id, p2.EndEmpresa_Id, p2.Pesagem_Id" & vbCrLf & _
                  "			               ) Rateio" & vbCrLf & _
                  "                ON P2.Empresa_Id    = Rateio.Empresa_Id" & vbCrLf & _
                  "               AND P2.EndEmpresa_Id = Rateio.EndEmpresa_Id" & vbCrLf & _
                  "               AND P2.Pesagem_Id    = Rateio.Pesagem_Id" & vbCrLf & _
                  "             group by P2.Empresa_Id, P2.EndEmpresa_Id, P2.Pesagem_Id, P2.Sequencia_id," & vbCrLf & _
                  "				         case when isnull(agrup.QtdePesagem,0) > 0 then agrup.QtdePesagem else Rateio.QtdePesagem end" & vbCrLf & _
                  "           ) as Mov" & vbCrLf & _
                  "    ON Mov.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
                  "   AND Mov.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
                  "   AND Mov.Pesagem_Id    = case" & vbCrLf & _
                  "			                    when isnull(P.RegistroMestre,0) > 0" & vbCrLf & _
                  "							      then P.RegistroMestre" & vbCrLf & _
                  "							      else P.Pesagem_Id" & vbCrLf & _
                  "							  end" & vbCrLf & _
                  "   AND Mov.Sequencia_Id  = case" & vbCrLf & _
                  "			                    when isnull(p.RegistroMestre,0) > 0" & vbCrLf & _
                  "							      then 0" & vbCrLf & _
                  "							      else p.Sequencia_Id" & vbCrLf & _
                  "							  end" & vbCrLf & _
                  " WHERE P.Empresa_Id    ='" & pEmpresa & "'" & vbCrLf & _
                  "   AND P.EndEmpresa_Id = " & pEndEmpresa & vbCrLf & _
                  "   AND P.Pesagem_Id    = " & pPesagem & vbCrLf & _
                  "   and p.Sequencia_Id  = " & pSequencia & vbCrLf

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Pesagem")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)
            Me.SelecionaPesagem = False

            Me.CodigoEmpresa = row("Empresa_Id")
            Me.EnderecoEmpresa = row("EndEmpresa_Id")
            Me.Codigo = row("Pesagem_Id")
            Me.Sequencia = row("Sequencia_Id")
            Me.SequenciaBanco = ds.Tables(0).Rows.Count - 1
            Me.CodigoPedido = row("Pedido")
            Me.CodigoProduto = row("Produto")
            Me.CodigoCliente = row("Cliente")
            Me.EnderecoCliente = row("EndCliente")
            Me.CodigoDepositante = row("Depositante").ToString
            Me.EnderecoDepositante = row("EndDepositante")
            Me.CodigoTransportador = row("Transportador")
            Me.EnderecoTransportador = row("EndTransportador")
            Me.CodigoMotorista = row("Motorista")
            Me.EnderecoMotorista = row("EndMotorista")
            Me.CodigoDeposito = row("Deposito")
            Me.EnderecoDeposito = row("EndDeposito")
            Me.CodigoTabelaDeClassificacao = row("TabelaDeClassificacao")
            Me.CodigoPlaca = row("Placa")
            Me.CodigoSituacao = row("Situacao")
            Me.CodigoOperacao = row("Operacao")
            Me.CodigoSubOperacao = row("SubOperacao")
            Me.EntradaSaida = IIf(IsDBNull(row("EntradaSaida")), "", row("EntradaSaida"))
            Me.PrimeiraPesagem = row("PrimeiraPesagem")
            Me.SegundaPesagem = row("SegundaPesagem")
            Me.BrutoBalanca = row("BrutoBalanca")
            Me.Desconto = row("BrutoBalanca") - row("Liquido")
            Me.Liquido = row("Liquido")
            Me.EntradaPatio = row("EntradaPatio")
            Me.EntradaBalanca = row("EntradaBalanca")
            Me.SaidaBalanca = row("SaidaBalanca")
            Me.CodigoViaDeTransporte = row("ViaTransporte")
            Me.Movimento = row("Movimento")
            Me.PesagemDeTerceiros = row("PesagemDeTerceiros")
            Me.PesagemMecanica = row("PesagemMecanica")
            Me.NumeroDaNota = row("NumeroDaNota")
            Me.SerieDaNota = row("SerieDaNota")
            Me.PesoFiscal = row("PesoFiscal")
            Me.UsuarioInclusao = row("UsuarioInclusao")
            Me.UsuarioAlteracao = row("UsuarioAlteracao")
            Me.DataAlteracao = row("UsuarioAlteracaoData")
            Me.UsuarioReimpressao = row("UsuarioReimpressao")
            Me.DataReimpressao = row("UsuarioReimpressaoData")
            Me.UsuarioCancelamento = row("UsuarioCancelamento")
            Me.DataCancelamento = row("UsuarioCancelamentoData")
            Me.Observacoes = row("Observacoes")
            Me.RegistroMestre = row("RegistroMestre")

            Me.TemRomaneio = row("TemRomaneio")
            Me.TemNota = row("TemNota")
            Me.TemAgrupamento = row("TemAgrupamento")
            Me.TemRateio = row("TemRateio")
            Me.TemRomaneioSemNota = row("TemRomaneioSemNota")
        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Public Erro As Exception
    Private _SelecionaPesagem As Boolean = False

    Private _IUD As String

    Private _CodigoEmpresa As String = ""
    Private _EnderecoEmpresa As Integer
    Private _Empresa As Cliente

    Private _Codigo As Integer
    Private _Sequencia As Integer
    Private _SequenciaBanco As Integer

    Private _CodigoPedido As Integer
    Private _Pedido As Pedido

    Private _CodigoProduto As String = ""
    Private _Produto As Produto

    Private _CodigoCliente As String = ""
    Private _EnderecoCliente As Integer
    Private _Consolidado As Boolean
    Private _Cliente As Cliente

    Private _CodigoDepositante As String = ""
    Private _EnderecoDepositante As Integer
    Private _Depositante As Cliente

    Private _CodigoTransportador As String = ""
    Private _EnderecoTransportador As Integer
    Private _Transportador As Cliente

    Private _CodigoMotorista As String = ""
    Private _EnderecoMotorista As Integer
    Private _Motorista As Cliente

    Private _CodigoDeposito As String = ""
    Private _EnderecoDeposito As Integer
    Private _Deposito As Cliente

    Private _CodigoPlaca As String = ""
    Private _Placa As Placa

    Private _CodigoSituacao As Integer
    Private _Situacao As Situacao

    Private _CodigoOperacao As Integer
    Private _Operacao As Operacao
    Private _CodigoSubOperacao As Integer
    Private _SubOperacao As SubOperacao

    Private _EntradaSaida As String = ""

    Private _PrimeiraPesagem As Decimal
    Private _SegundaPesagem As Decimal
    Private _BrutoBalanca As Decimal
    Private _Desconto As Decimal
    Private _Liquido As Decimal

    Private _EntradaPatio As DateTime
    Private _EntradaBalanca As DateTime
    Private _SaidaBalanca As DateTime

    Private _CodigoViaDeTransporte As Integer
    Private _ViaDeTransporte As ViaDeTransporte

    Private _Movimento As DateTime = Date.Today

    Private _PesagemDeTerceiros As Decimal
    Private _PesagemMecanica As Decimal

    Private _NumeroDaNota As Integer
    Private _SerieDaNota As String = ""
    Private _PesoFiscal As Decimal

    Private _Observacoes As String = ""
    Private _RegistroMestre As Integer

    Private _CodigoTabelaDeClassificacao As Integer
    Private _TabelaDeClassificacao As TabelaDeClassificacao

    Public CarregandoAnalises As Boolean = False
    Private _Analises As ListPesagemXAnalises

    Private _CriarRomaneio As Boolean = False
    Private _Romaneios As ListRomaneio
    Private _Processo As String = ""

    Private _CodigoAutorizacao As Integer
    Private _AutorizacaoDeRetirada As AutorizacaoDeRetirada

    '*********** Atributos **************
    Private _TemRomaneio As Boolean
    Private _TemNota As Boolean
    Private _TemRomaneioSemNota As Boolean
    Private _TemAgrupamento As Boolean
    Private _TemRateio As Boolean


    '****** Controle de Usuario **********
    Private _UsuarioInclusao As String
    Private _UsuarioAlteracao As String
    Private _UsuarioAlteracaoData As DateTime?
    Private _UsuarioReimpressao As String
    Private _UsuarioReimpressaoData As DateTime?
    Private _UsuarioCancelamento As String
    Private _UsuarioCancelamentoData As DateTime?
#End Region

#Region "Propriedades"
    Public Property SelecionaPesagem() As Boolean
        Get
            Return _SelecionaPesagem
        End Get
        Set(ByVal value As Boolean)
            _SelecionaPesagem = value
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

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property EnderecoEmpresa() As Integer
        Get
            Return _EnderecoEmpresa
        End Get
        Set(ByVal value As Integer)
            _EnderecoEmpresa = value
            Me.Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Trim.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EnderecoEmpresa)
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

    Public Property Sequencia() As Integer
        Get
            Return _Sequencia
        End Get
        Set(ByVal value As Integer)
            _Sequencia = value
        End Set
    End Property

    Public Property RegistroMestre() As Integer
        Get
            Return _RegistroMestre
        End Get
        Set(ByVal value As Integer)
            _RegistroMestre = value
        End Set
    End Property

    Public Property SequenciaBanco() As Integer
        Get
            Return _SequenciaBanco
        End Get
        Set(ByVal value As Integer)
            _SequenciaBanco = value
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
            If _Pedido Is Nothing And Me.CodigoPedido > 0 Then _Pedido = New Pedido(_CodigoEmpresa, _EnderecoEmpresa, _CodigoPedido)
            Return _Pedido
        End Get
        Set(ByVal value As Pedido)
            _Pedido = value
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
            If _Produto Is Nothing Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
        End Set
    End Property

    Public Property EnderecoCliente() As Integer
        Get
            Return _EnderecoCliente
        End Get
        Set(ByVal value As Integer)
            _EnderecoCliente = value
        End Set
    End Property

    Public Property Consolidado() As Boolean
        Get
            Return _Consolidado
        End Get
        Set(ByVal value As Boolean)
            _Consolidado = value
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Trim.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EnderecoCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property

    Public Property CodigoDepositante() As String
        Get
            Return _CodigoDepositante
        End Get
        Set(ByVal value As String)
            _CodigoDepositante = value
        End Set
    End Property

    Public Property EnderecoDepositante() As Integer
        Get
            Return _EnderecoDepositante
        End Get
        Set(ByVal value As Integer)
            _EnderecoDepositante = value
        End Set
    End Property

    Public Property Depositante() As Cliente
        Get
            If _Depositante Is Nothing And _CodigoDepositante.Trim.Length > 0 Then _Depositante = New Cliente(_CodigoDepositante, _EnderecoDepositante)
            Return _Depositante
        End Get
        Set(ByVal value As Cliente)
            _Depositante = value
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
            If _Transportador Is Nothing And _CodigoTransportador.Trim.Length > 0 Then _Transportador = New Cliente(_CodigoTransportador, _EnderecoTransportador)
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
            If _Motorista Is Nothing And _CodigoMotorista.Trim.Length > 0 Then _Motorista = New Cliente(_CodigoMotorista, _EnderecoMotorista)
            Return _Motorista
        End Get
        Set(ByVal value As Cliente)
            _Motorista = value
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
            If _Deposito Is Nothing And _CodigoDeposito.Trim.Length > 0 Then _Deposito = New Cliente(_CodigoDeposito, _EnderecoDeposito)
            Return _Deposito
        End Get
        Set(ByVal value As Cliente)
            _Deposito = value
        End Set
    End Property

    Public Property CodigoPlaca() As String
        Get
            Return _CodigoPlaca
        End Get
        Set(ByVal value As String)
            _CodigoPlaca = value
        End Set
    End Property

    Public Property Placa() As Placa
        Get
            If _Placa Is Nothing And _CodigoPlaca.Trim.Length > 0 Then _Placa = New Placa(_CodigoPlaca)
            Return _Placa
        End Get
        Set(ByVal value As Placa)
            _Placa = value
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

    Public Property Situacao() As Situacao
        Get
            If _Situacao Is Nothing And _CodigoSituacao > 0 Then _Situacao = New Situacao(_CodigoSituacao)
            Return _Situacao
        End Get
        Set(ByVal value As Situacao)
            _Situacao = value
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
            If _Operacao Is Nothing And _CodigoOperacao > 0 Then _Operacao = New Operacao(_CodigoOperacao)
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
            If _SubOperacao Is Nothing And _CodigoSubOperacao > 0 And _CodigoOperacao > 0 Then _SubOperacao = New SubOperacao(_CodigoOperacao, _CodigoSubOperacao)
            Return _SubOperacao
        End Get
        Set(ByVal value As SubOperacao)
            _SubOperacao = value
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

    '*********   Pesagem *************** 
    Public Property PrimeiraPesagem() As Decimal
        Get
            Return _PrimeiraPesagem
        End Get
        Set(ByVal value As Decimal)
            _PrimeiraPesagem = value
        End Set
    End Property

    Public Property SegundaPesagem() As Decimal
        Get
            Return _SegundaPesagem
        End Get
        Set(ByVal value As Decimal)
            _SegundaPesagem = value
        End Set
    End Property

    Public Property BrutoBalanca() As Decimal
        Get
            Return _BrutoBalanca
        End Get
        Set(ByVal value As Decimal)
            _BrutoBalanca = value
        End Set
    End Property

    Public Property Desconto() As Decimal
        Get
            Return _Desconto
        End Get
        Set(ByVal value As Decimal)
            _Desconto = value
        End Set
    End Property

    Public Property Liquido() As Decimal
        Get
            Return _Liquido
        End Get
        Set(ByVal value As Decimal)
            _Liquido = value
        End Set
    End Property

    Public Property EntradaPatio() As DateTime
        Get
            Return _EntradaPatio
        End Get
        Set(ByVal value As DateTime)
            _EntradaPatio = value
        End Set
    End Property

    Public Property EntradaBalanca() As DateTime
        Get
            Return _EntradaBalanca
        End Get
        Set(ByVal value As DateTime)
            _EntradaBalanca = value
        End Set
    End Property

    Public Property SaidaBalanca() As DateTime
        Get
            Return _SaidaBalanca
        End Get
        Set(ByVal value As DateTime)
            _SaidaBalanca = value
        End Set
    End Property

    Public Property CodigoViaDeTransporte() As Integer
        Get
            Return _CodigoViaDeTransporte
        End Get
        Set(ByVal value As Integer)
            _CodigoViaDeTransporte = value
        End Set
    End Property

    Public Property ViaDeTransporte() As ViaDeTransporte
        Get
            If _ViaDeTransporte Is Nothing And _CodigoViaDeTransporte > 0 Then _ViaDeTransporte = New ViaDeTransporte(_CodigoViaDeTransporte)
            Return _ViaDeTransporte
        End Get
        Set(ByVal value As ViaDeTransporte)
            _ViaDeTransporte = value
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

    Public Property PesagemDeTerceiros() As Decimal
        Get
            Return _PesagemDeTerceiros
        End Get
        Set(ByVal value As Decimal)
            _PesagemDeTerceiros = value
        End Set
    End Property

    Public Property PesagemMecanica() As Decimal
        Get
            Return _PesagemMecanica
        End Get
        Set(ByVal value As Decimal)
            _PesagemMecanica = value
        End Set
    End Property

    '**************************************************
    '**********  INFORMACOES DA NOTA ******************
    '**************************************************
    Public Property NumeroDaNota() As Integer
        Get
            Return _NumeroDaNota
        End Get
        Set(ByVal value As Integer)
            _NumeroDaNota = value
        End Set
    End Property

    Public Property SerieDaNota() As String
        Get
            Return _SerieDaNota
        End Get
        Set(ByVal value As String)
            _SerieDaNota = value
        End Set
    End Property

    Public Property PesoFiscal() As Decimal
        Get
            Return _PesoFiscal
        End Get
        Set(ByVal value As Decimal)
            _PesoFiscal = value
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

    '**************************************************
    '****************  ANALISES  **********************
    '**************************************************
    Public Property CodigoTabelaDeClassificacao() As Integer
        Get
            Return _CodigoTabelaDeClassificacao
        End Get
        Set(ByVal value As Integer)
            _CodigoTabelaDeClassificacao = value
            _TabelaDeClassificacao = Nothing
            'If Analises IsNot Nothing AndAlso Analises.Count > 0 Then Analises.CalcularDescontos()
        End Set
    End Property

    Public Property TabelaDeClassificacao() As TabelaDeClassificacao
        Get
            If _TabelaDeClassificacao Is Nothing And _CodigoTabelaDeClassificacao > 0 Then _TabelaDeClassificacao = New TabelaDeClassificacao(_CodigoTabelaDeClassificacao)
            Return _TabelaDeClassificacao
        End Get
        Set(ByVal value As TabelaDeClassificacao)
            _TabelaDeClassificacao = value
            _Analises = Nothing
        End Set
    End Property

    Public Property Analises() As ListPesagemXAnalises
        Get
            If _Analises Is Nothing And _CodigoTabelaDeClassificacao > 0 Then _Analises = New ListPesagemXAnalises(Me)
            Return _Analises
        End Get
        Set(ByVal value As ListPesagemXAnalises)
            _Analises = value
        End Set
    End Property

    '**************************************************
    '**************************************************

    Public Property Processo() As String
        Get
            Return _Processo
        End Get
        Set(ByVal value As String)
            _Processo = value
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

    Public Property AutorizacaoDeRetirada() As AutorizacaoDeRetirada
        Get
            If _AutorizacaoDeRetirada Is Nothing And _CodigoAutorizacao > 0 Then _AutorizacaoDeRetirada = New AutorizacaoDeRetirada(_CodigoEmpresa, _EnderecoEmpresa, Me.CodigoPedido, _CodigoAutorizacao, _SubOperacao.Classe)
            Return _AutorizacaoDeRetirada
        End Get
        Set(ByVal value As AutorizacaoDeRetirada)
            _AutorizacaoDeRetirada = value
        End Set
    End Property


    '****************  Romaneio  ***********************
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
            Return 1
        End Get
        Set(ByVal value As Integer)

        End Set
    End Property


    Public Property Romaneios() As ListRomaneio
        Get
            If _Romaneios Is Nothing Then _Romaneios = New ListRomaneio(Me)
            Return _Romaneios
        End Get
        Set(ByVal value As ListRomaneio)
            _Romaneios = value
        End Set
    End Property

    '****************  Atributos  ***********************
    Public Property TemRomaneio() As Boolean
        Get
            Return _TemRomaneio
        End Get
        Set(ByVal value As Boolean)
            _TemRomaneio = value
        End Set
    End Property

    Public Property TemNota() As Boolean
        Get
            Return _TemNota
        End Get
        Set(ByVal value As Boolean)
            _TemNota = value
        End Set
    End Property

    Public Property TemRomaneioSemNota() As Boolean
        Get
            Return _TemRomaneioSemNota
        End Get
        Set(ByVal value As Boolean)
            _TemRomaneioSemNota = value
        End Set
    End Property

    Public Property TemAgrupamento As Boolean
        Get
            Return _TemAgrupamento
        End Get
        Set(value As Boolean)
            _TemAgrupamento = value
        End Set
    End Property

    Public Property TemRateio As Boolean
        Get
            Return _TemRateio
        End Get
        Set(value As Boolean)
            _TemRateio = value
        End Set
    End Property

    '********** Controle de Usuarios  *************
    Public Property UsuarioInclusao() As String
        Get
            Return _UsuarioInclusao
        End Get
        Set(ByVal value As String)
            _UsuarioInclusao = value
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

    Public Property DataAlteracao() As DateTime?
        Get
            Return _UsuarioAlteracaoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioAlteracaoData = value
        End Set
    End Property

    Public Property UsuarioReimpressao() As String
        Get
            Return _UsuarioReimpressao
        End Get
        Set(ByVal value As String)
            _UsuarioReimpressao = value
        End Set
    End Property

    Public Property DataReimpressao() As DateTime?
        Get
            Return _UsuarioReimpressaoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioReimpressaoData = value
        End Set
    End Property

    Public Property UsuarioCancelamento() As String
        Get
            Return _UsuarioCancelamento
        End Get
        Set(ByVal value As String)
            _UsuarioCancelamento = value
        End Set
    End Property

    Public Property DataCancelamento() As DateTime?
        Get
            Return _UsuarioCancelamentoData
        End Get
        Set(ByVal value As DateTime?)
            _UsuarioCancelamentoData = value
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

        If Banco.GravaBanco(Sqls) Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Dim usa As DateTime = DateTime.Now
        Dim urd As DateTime = DateTime.Now

        Select Case Me.IUD
            Case "I"
                If Me.Codigo.Equals(0) Then
                    Dim N As New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 101)
                    Me.Codigo = N.Sequencia + 1
                    Sqls.Add(N.IncrementarNumeradorSql)
                End If

                sql = "Insert Into Pesagem " & vbCrLf & _
                      " (Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id, Pedido, Produto, " & vbCrLf & _
                      " Cliente, EndCliente, Depositante, EndDepositante, Transportador, EndTransportador, " & vbCrLf & _
                      " Motorista, EndMotorista, Deposito, EndDeposito, TabelaDeClassificacao, Placa, Situacao, " & vbCrLf & _
                      " Operacao, SubOperacao, EntradaSaida, PrimeiraPesagem, SegundaPesagem, " & vbCrLf & _
                      " BrutoBalanca, Liquido, EntradaPatio, EntradaBalanca, SaidaBalanca, " & vbCrLf & _
                      " ViaTransporte, Movimento, PesagemDeTerceiros, PesagemMecanica, " & vbCrLf & _
                      " NumeroDaNota, SerieDaNota, PesoFiscal, UsuarioInclusao, UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf & _
                      " UsuarioReimpressao, UsuarioReimpressaoData, UsuarioCancelamento, UsuarioCancelamentoData, " & vbCrLf & _
                      " Observacoes, RegistroMestre) " & vbCrLf & _
                      " Values ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & "," & Me.Sequencia & "," & Me.CodigoPedido & ",'" & Me.CodigoProduto & "'," & vbCrLf & _
                      "'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.CodigoDepositante & "'," & Me.EnderecoDepositante & ",'" & Me.CodigoTransportador & "'," & Me.EnderecoTransportador & "," & vbCrLf & _
                      "'" & Me.CodigoMotorista & "', " & Me.EnderecoMotorista & ",'" & Me.CodigoDeposito & "'," & Me.EnderecoDeposito & "," & Me.CodigoTabelaDeClassificacao & ",'" & Me.CodigoPlaca & "'," & Me.CodigoSituacao & "," & vbCrLf & _
                      "" & Me.CodigoOperacao & "," & Me.CodigoSubOperacao & "," & IIf(String.IsNullOrWhiteSpace(Me.EntradaSaida), "NULL", "'" & Me.EntradaSaida & "'") & "," & Str(Me.PrimeiraPesagem) & "," & Str(Me.SegundaPesagem) & "," & vbCrLf & _
                      "" & Str(Me.BrutoBalanca) & "," & Str(Me.Liquido) & ",'" & Me.EntradaPatio.ToString("yyyy-MM-dd HH:mm:ss") & "','" & Me.EntradaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "', " & IIf(Me.SegundaPesagem > 0, "'" & Me.SaidaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & ", " & vbCrLf & _
                      "" & Me.CodigoViaDeTransporte & ",'" & Me.Movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'," & Str(Me.PesagemDeTerceiros) & "," & Str(Me.PesagemMecanica) & "," & vbCrLf & _
                      "" & Me.NumeroDaNota & ",'" & Me.SerieDaNota & "'," & Str(Me.PesoFiscal) & ",'" & Me.UsuarioInclusao & "'," & IIf(String.IsNullOrWhiteSpace(UsuarioAlteracao), "NULL", "'" & UsuarioAlteracao & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf & _
                      "" & IIf(String.IsNullOrWhiteSpace(UsuarioReimpressao), "NULL", "'" & UsuarioReimpressao & "'") & ",'" & urd.ToString("yyyy-MM-dd HH:mm:ss") & "'," & IIf(String.IsNullOrWhiteSpace(UsuarioCancelamento), "NULL", "'" & UsuarioCancelamento & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf & _
                      "'" & Me.Observacoes & "'," & Me.RegistroMestre & ")"

                Sqls.Add(sql)
                '***********************************************************************
                '********* Procedimento para Salvar as tabelas relacionadas   **********
                '***********************************************************************
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = "Update Pesagem set " & vbCrLf & _
                      "     Pedido                  = " & Me.CodigoPedido & vbCrLf & _
                      "    ,Produto                 ='" & Me.CodigoProduto & "'" & vbCrLf & _
                      "    ,Cliente                 ='" & Me.CodigoCliente & "'" & vbCrLf & _
                      "    ,EndCliente              = " & Me.EnderecoCliente & vbCrLf & _
                      "    ,Depositante             ='" & Me.CodigoDepositante & "'" & vbCrLf & _
                      "    ,EndDepositante          = " & Me.EnderecoDepositante & vbCrLf & _
                      "    ,Transportador           ='" & Me.CodigoTransportador & "'" & vbCrLf & _
                      "    ,EndTransportador        = " & Me.EnderecoTransportador & vbCrLf & _
                      "    ,Motorista               ='" & Me.CodigoMotorista & "'" & vbCrLf & _
                      "    ,EndMotorista            = " & Me.EnderecoMotorista & vbCrLf & _
                      "    ,Deposito                ='" & Me.CodigoDeposito & "'" & vbCrLf & _
                      "    ,EndDeposito             = " & Me.EnderecoDeposito & vbCrLf & _
                      "    ,TabelaDeClassificacao   = " & Me.CodigoTabelaDeClassificacao & vbCrLf & _
                      "    ,Placa                   ='" & Me.CodigoPlaca & "'" & vbCrLf & _
                      "    ,Situacao                = " & Me.CodigoSituacao & vbCrLf & _
                      "    ,Operacao                = " & Me.CodigoOperacao & vbCrLf & _
                      "    ,SubOperacao             = " & Me.CodigoSubOperacao & vbCrLf & _
                      "    ,EntradaSaida            ='" & Me.EntradaSaida & "'" & vbCrLf & _
                      "    ,PrimeiraPesagem         = " & Me.PrimeiraPesagem & vbCrLf & _
                      "    ,SegundaPesagem          = " & Me.SegundaPesagem & vbCrLf & _
                      "    ,BrutoBalanca            = " & Me.BrutoBalanca & vbCrLf & _
                      "    ,Liquido                 = " & Me.Liquido & vbCrLf & _
                      "    ,SaidaBalanca            = " & IIf(Me.SegundaPesagem > 0, "'" & Me.SaidaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & vbCrLf & _
                      "    ,ViaTransporte           = " & Me.CodigoViaDeTransporte & vbCrLf & _
                      "    ,Movimento               ='" & Me.Movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                      "    ,PesagemDeTerceiros      = " & Me.PesagemDeTerceiros & vbCrLf & _
                      "    ,PesagemMecanica         = " & Me.PesagemMecanica & vbCrLf & _
                      "    ,NumeroDaNota            = " & Me.NumeroDaNota & vbCrLf & _
                      "    ,SerieDaNota             ='" & Me.SerieDaNota & "'" & vbCrLf & _
                      "    ,PesoFiscal              = " & Me.PesoFiscal & vbCrLf & _
                      "    ,UsuarioInclusao         ='" & Me.UsuarioInclusao & "'" & vbCrLf & _
                      "    ,UsuarioAlteracao        ='" & Me.UsuarioAlteracao & "'" & vbCrLf & _
                      "    ,UsuarioAlteracaoData    = '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                      "    ,UsuarioReimpressao      ='" & Me.UsuarioReimpressao & "'" & vbCrLf & _
                      "    ,UsuarioReimpressaoData  ='" & String.Format("{0:yyyy-MM-dd HH:mm:ss}", Me.DataReimpressao) & "'" & vbCrLf & _
                      "    ,UsuarioCancelamento     ='" & _UsuarioCancelamento & "'" & vbCrLf & _
                      "    ,UsuarioCancelamentoData ='" & String.Format("{0:yyyy-MM-dd HH:mm:ss}", Me.DataCancelamento) & "'" & vbCrLf & _
                      "    ,Observacoes             ='" & Me.Observacoes & "'" & vbCrLf & _
                      "    ,RegistroMestre          = " & Me.RegistroMestre & vbCrLf & _
                      "	Where Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                      "   and Pesagem_Id    = " & Me.Codigo & vbCrLf & _
                      "   and Sequencia_Id  = " & Me.Sequencia

                SalvarTabelasRelacionadasSql(Sqls)
                Sqls.Add(sql)
            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)

                sql = " Update Pesagem Set" & vbCrLf & _
                      "     Situacao                = 2" & vbCrLf & _
                      "    ,Observacoes             ='" & Me.Observacoes & "'" & vbCrLf & _
                      "    ,RegistroMestre          = " & Me.RegistroMestre & vbCrLf & _
                      "    ,UsuarioCancelamento     = '" & Me.UsuarioCancelamento & "'" & vbCrLf & _
                      "    ,UsuarioCancelamentoData = '" & CDate(Me.DataCancelamento).ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf & _
                      "	Where Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                      "   and EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
                      "   and Pesagem_Id    = " & Me.Codigo & vbCrLf & _
                      "   and Sequencia_Id  = " & Me.Sequencia

                Sqls.Add(sql)
            Case "S" 'Gravar Sequência das Alterações
                sql = "Insert Into Pesagem " & vbCrLf & _
                      " (Empresa_Id, EndEmpresa_Id, Pesagem_Id, Sequencia_Id, Pedido, Produto, " & vbCrLf & _
                      " Cliente, EndCliente, Depositante, EndDepositante, Transportador, EndTransportador, " & vbCrLf & _
                      " Motorista, EndMotorista, Deposito, EndDeposito, TabelaDeClassificacao, Placa, Situacao, " & vbCrLf & _
                      " Operacao, SubOperacao, EntradaSaida, PrimeiraPesagem, SegundaPesagem, " & vbCrLf & _
                      " BrutoBalanca, Liquido, EntradaPatio, EntradaBalanca, SaidaBalanca, " & vbCrLf & _
                      " ViaTransporte, Movimento, PesagemDeTerceiros, PesagemMecanica, " & vbCrLf & _
                      " NumeroDaNota, SerieDaNota, PesoFiscal, UsuarioInclusao, UsuarioAlteracao, UsuarioAlteracaoData, " & vbCrLf & _
                      " UsuarioReimpressao, UsuarioReimpressaoData, UsuarioCancelamento, UsuarioCancelamentoData, " & vbCrLf & _
                      " Observacoes, RegistroMestre) " & vbCrLf & _
                      " Values ('" & Me.CodigoEmpresa & "'," & Me.EnderecoEmpresa & "," & Me.Codigo & "," & Me.Sequencia & "," & Me.CodigoPedido & ",'" & Me.CodigoProduto & "'," & vbCrLf & _
                      "'" & Me.CodigoCliente & "'," & Me.EnderecoCliente & ",'" & Me.CodigoDepositante & "'," & Me.EnderecoDepositante & ",'" & Me.CodigoTransportador & "'," & Me.EnderecoTransportador & "," & vbCrLf & _
                      "'" & Me.CodigoMotorista & "', " & Me.EnderecoMotorista & ",'" & Me.CodigoDeposito & "'," & Me.EnderecoDeposito & "," & Me.CodigoTabelaDeClassificacao & ",'" & Me.CodigoPlaca & "'," & Me.CodigoSituacao & "," & vbCrLf & _
                      "" & Me.CodigoOperacao & "," & Me.CodigoSubOperacao & "," & IIf(String.IsNullOrWhiteSpace(Me.EntradaSaida), "NULL", "'" & Me.EntradaSaida & "'") & "," & Str(Me.PrimeiraPesagem) & "," & Str(Me.SegundaPesagem) & "," & vbCrLf & _
                      "" & Str(Me.BrutoBalanca) & "," & Str(Me.Liquido) & "," & IIf(Me.SegundaPesagem = 0, "'" & Me.EntradaPatio.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & ", " & IIf(Me.SegundaPesagem = 0, "'" & Me.EntradaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & ", " & IIf(Me.SegundaPesagem > 0, "'" & Me.SaidaBalanca.ToString("yyyy-MM-dd HH:mm:ss") & "'", "NULL") & ", " & vbCrLf & _
                      "" & Me.CodigoViaDeTransporte & ",'" & Me.Movimento.ToString("yyyy-MM-dd HH:mm:ss") & "'," & Str(Me.PesagemDeTerceiros) & "," & Str(Me.PesagemMecanica) & "," & vbCrLf & _
                      "" & Me.NumeroDaNota & ",'" & Me.SerieDaNota & "'," & Str(Me.PesoFiscal) & ",'" & Me.UsuarioInclusao & "'," & IIf(String.IsNullOrWhiteSpace(UsuarioAlteracao), "NULL", "'" & UsuarioAlteracao & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf & _
                      "" & IIf(String.IsNullOrWhiteSpace(UsuarioReimpressao), "NULL", "'" & UsuarioReimpressao & "'") & ",'" & urd.ToString("yyyy-MM-dd HH:mm:ss") & "'," & IIf(String.IsNullOrWhiteSpace(UsuarioCancelamento), "NULL", "'" & UsuarioCancelamento & "'") & ", '" & usa.ToString("yyyy-MM-dd HH:mm:ss") & "', " & vbCrLf & _
                      "'" & Me.Observacoes & "'," & Me.RegistroMestre & ")"

                Sqls.Add(sql)

                Analises.Parent.IUD = "I"
                Analises.SalvarSql(Sqls)
        End Select
    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        Analises.SalvarSql(Sqls)

        If Romaneios.Count > 0 Then
            If Me.Romaneios.Count = 1 AndAlso Not Me.IUD.Contains("D") Then
                Me.Romaneios(0).IUD = "U"
                Me.Romaneios(0).CodigoEmpresa = Me.CodigoEmpresa
                Me.Romaneios(0).EnderecoEmpresa = Me.EnderecoEmpresa
                Me.Romaneios(0).EntradaSaida = Me.EntradaSaida
                Me.Romaneios(0).CodigoPedido = Me.CodigoPedido

                If Me.CodigoDeposito.Length > 0 Then
                    Me.Romaneios(0).CodigoDeposito = Me.CodigoDeposito
                    Me.Romaneios(0).EnderecoDeposito = Me.EnderecoDeposito
                Else
                    Me.Romaneios(0).CodigoDeposito = Me.CodigoEmpresa
                    Me.Romaneios(0).EnderecoDeposito = Me.EnderecoEmpresa
                End If

                If Me.EntradaSaida = "E" Then
                    Me.Romaneios(0).CodigoDestino = Me.CodigoEmpresa
                    Me.Romaneios(0).EnderecoDestino = Me.EnderecoEmpresa
                Else
                    Me.Romaneios(0).CodigoDestino = Me.CodigoCliente
                    Me.Romaneios(0).EnderecoDestino = Me.EnderecoCliente
                End If

                Me.Romaneios(0).CodigoTransbordo = ""
                Me.Romaneios(0).EnderecoTransbordo = 0

                Me.Romaneios(0).CodigoProduto = Me.CodigoProduto
                Me.Romaneios(0).CodigoOperacao = Me.CodigoOperacao
                Me.Romaneios(0).CodigoSubOperacao = Me.CodigoSubOperacao
                Me.Romaneios(0).Movimento = Me.Movimento

                Me.Romaneios(0).PesoBruto = Me.BrutoBalanca
                Me.Romaneios(0).Desconto = Me.BrutoBalanca - Me.Liquido
                Me.Romaneios(0).PesoLiquido = Me.Liquido

                Me.Romaneios(0).Observacoes = Me.Observacoes
                Me.Romaneios(0).Processo = Me.Processo

                Me.Romaneios(0).CodigoAutorizacao = Me.CodigoAutorizacao

                Me.Romaneios(0).Pesagens = New ListRomaneioXPesagem()
                Dim RomaneioXPesagem As New Negocio.RomaneioXPesagem(Me.Romaneios(0), Me)
                Me.Romaneios(0).Pesagens.Add(RomaneioXPesagem)

                Me.Romaneios(0).DescontosAnalises.Clear()
                For Each row As PesagemXAnalises In Me.Analises
                    Dim RxP As New RomaneioXDesconto(Me.Romaneios(0))
                    RxP.CodigoAnalise = row.CodigoAnalise
                    RxP.Desconto = row.Desconto
                    RxP.Percentual = row.Percentual
                    RxP.Indice = row.Indice
                    Me.Romaneios(0).DescontosAnalises.Add(RxP)
                Next
                Me.Romaneios(0).SalvarSql(Sqls, False)
            Else
                For Each ron As [Lib].Negocio.Romaneio In Romaneios
                    ron.SalvarSql(Sqls, False)
                Next
            End If
        ElseIf (SegundaPesagem > 0 Or Processo = "NOTAFISCAL") And Not Me.IUD.Contains("D") Then
            Dim N As New Numerador(Me.CodigoEmpresa, Me.EnderecoEmpresa, 110)
            Sqls.Add(N.IncrementarNumeradorSql)

            Dim MyRomaneio = New Negocio.Romaneio()
            MyRomaneio.Codigo = N.Sequencia + 1

            MyRomaneio.IUD = "I"
            MyRomaneio.CodigoEmpresa = Me.CodigoEmpresa
            MyRomaneio.EnderecoEmpresa = Me.EnderecoEmpresa

            MyRomaneio.EntradaSaida = Me.EntradaSaida
            MyRomaneio.CodigoPedido = Me.CodigoPedido

            If Me.CodigoDeposito.Length > 0 Then
                MyRomaneio.CodigoDeposito = Me.CodigoDeposito
                MyRomaneio.EnderecoDeposito = Me.EnderecoDeposito
            Else
                MyRomaneio.CodigoDeposito = Me.CodigoEmpresa
                MyRomaneio.EnderecoDeposito = Me.EnderecoEmpresa
            End If

            If Me.EntradaSaida = "E" Then
                MyRomaneio.CodigoDestino = Me.CodigoEmpresa
                MyRomaneio.EnderecoDestino = Me.EnderecoEmpresa
            Else
                MyRomaneio.CodigoDestino = Me.CodigoCliente
                MyRomaneio.EnderecoDestino = Me.EnderecoCliente
            End If

            MyRomaneio.CodigoTransbordo = ""
            MyRomaneio.EnderecoTransbordo = 0

            MyRomaneio.CodigoProduto = Me.CodigoProduto
            MyRomaneio.CodigoOperacao = Me.CodigoOperacao
            MyRomaneio.CodigoSubOperacao = Me.CodigoSubOperacao
            MyRomaneio.Movimento = Me.Movimento

            MyRomaneio.PesoBruto = Me.BrutoBalanca
            MyRomaneio.Desconto = Me.BrutoBalanca - Me.Liquido
            MyRomaneio.PesoLiquido = Me.Liquido

            MyRomaneio.Observacoes = Me.Observacoes
            MyRomaneio.Processo = Me.Processo

            MyRomaneio.CodigoAutorizacao = Me.CodigoAutorizacao
            MyRomaneio.Pesagens = New ListRomaneioXPesagem()

            Dim RomaneioXPesagem As New Negocio.RomaneioXPesagem(MyRomaneio, Me)
            MyRomaneio.Pesagens.Add(RomaneioXPesagem)

            MyRomaneio.DescontosAnalises.Clear()
            For Each row As PesagemXAnalises In Me.Analises
                Dim RxP As New RomaneioXDesconto(MyRomaneio)
                RxP.CodigoAnalise = row.CodigoAnalise
                RxP.Desconto = row.Desconto
                RxP.Percentual = row.Percentual
                RxP.Indice = row.Indice
                MyRomaneio.DescontosAnalises.Add(RxP)
            Next

            MyRomaneio.SalvarSql(Sqls, False)
        End If
    End Sub

    Public Sub DesfazerRateio(ByRef Sqls As ArrayList)
        For Each xRomaneio As Romaneio In Me.Romaneios
            xRomaneio.IUD = "D"
            xRomaneio.SalvarSql(Sqls, False)
        Next

        Dim SqlN As String = "exec sp_Numerador '" & CodigoEmpresa & "'," & EnderecoEmpresa & ",110"
        Dim dsN As New DataSet
        Dim objBanco As New AcessaBanco()
        dsN = objBanco.ConsultaDataSet(SqlN, "Numerador")

        Dim pRomaneio As New Romaneio()
        pRomaneio.Codigo = dsN.Tables(0).Rows(0).Item(0)

        pRomaneio.IUD = "I"
        pRomaneio.CodigoEmpresa = CodigoEmpresa
        pRomaneio.EnderecoEmpresa = EnderecoEmpresa

        pRomaneio.EntradaSaida = EntradaSaida
        pRomaneio.CodigoPedido = CodigoPedido

        If CodigoDeposito.Length > 0 Then
            pRomaneio.CodigoDeposito = CodigoDeposito
            pRomaneio.EnderecoDeposito = EnderecoDeposito
        Else
            pRomaneio.CodigoDeposito = CodigoEmpresa
            pRomaneio.EnderecoDeposito = EnderecoEmpresa
        End If

        If EntradaSaida = "E" Then
            pRomaneio.CodigoDestino = CodigoEmpresa
            pRomaneio.EnderecoDestino = EnderecoEmpresa
        Else
            pRomaneio.CodigoDestino = CodigoCliente
            pRomaneio.EnderecoDestino = EnderecoCliente
        End If

        pRomaneio.CodigoTransbordo = ""
        pRomaneio.EnderecoTransbordo = 0

        pRomaneio.CodigoProduto = CodigoProduto
        pRomaneio.CodigoOperacao = CodigoOperacao
        pRomaneio.CodigoSubOperacao = CodigoSubOperacao
        pRomaneio.Movimento = Movimento

        pRomaneio.PesoBruto = BrutoBalanca
        pRomaneio.Desconto = BrutoBalanca - Liquido
        pRomaneio.PesoLiquido = Liquido

        pRomaneio.Observacoes = Observacoes
        pRomaneio.Processo = ""

        pRomaneio.CodigoAutorizacao = CodigoAutorizacao

        pRomaneio.DescontosAnalises = New ListRomaneioXDesconto(pRomaneio)
        pRomaneio.DescontosAnalises.Clear()

        For Each row As PesagemXAnalises In Analises
            Dim RxD As New RomaneioXDesconto(pRomaneio)
            RxD.CodigoAnalise = row.CodigoAnalise
            RxD.Desconto = row.Desconto
            RxD.Percentual = row.Percentual
            RxD.Indice = row.Indice
            pRomaneio.DescontosAnalises.Add(RxD)
        Next

        pRomaneio.SalvarSql(Sqls, False)

        Dim rxp As New RomaneioXPesagem()
        rxp.IUD = "I"
        rxp.CodigoEmpresa = pRomaneio.CodigoEmpresa
        rxp.EndEmpresa = pRomaneio.EnderecoEmpresa
        rxp.CodigoRomaneio = pRomaneio.Codigo
        rxp.CodigoPesagem = Codigo
        rxp.Sequencia = 0

        rxp.SalvarSql(Sqls)
    End Sub

    Public Sub SalvarSqlAgrupamento(ByRef Sqls As ArrayList, ByVal Pesagens As ListPesagem)
        SalvarSql(Sqls)

        For Each Ps As Pesagem In Pesagens
            If Ps.SelecionaPesagem Then
                'Ps.IUD = "D"
                'Ps.CodigoSituacao = 2
                'Ps.RegistroMestre = Me.Codigo
                'If Ps.Observacoes.Length > 0 Then
                '    Ps.Observacoes = Ps.Observacoes & " - LAUDO AGRUPADO"
                'Else
                '    Ps.Observacoes = "LAUDO AGRUPADO"
                'End If

                'Ps.SalvarSql(Sqls)

                'For Each ro As Romaneio In Ps.Romaneios
                '    ro.IUD = "D"
                '    ro.SalvarSql(Sqls)
                'Next

                Dim Sql As String = " Update Pesagem Set" & vbCrLf &
                                      "     Situacao                = 2" & vbCrLf

                If Ps.Observacoes.Length > 0 Then
                    Sql &= "    ,Observacoes             ='" & Ps.Observacoes & " - LAUDO AGRUPADO'" & vbCrLf
                Else
                    Sql &= "    ,Observacoes             ='LAUDO AGRUPADO'" & vbCrLf
                End If

                Sql &= "    ,RegistroMestre          = " & Me.Codigo & vbCrLf &
                        "    ,UsuarioCancelamento     = '" & Me.UsuarioInclusao & "'" & vbCrLf &
                        "    ,UsuarioCancelamentoData = '" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                        "	Where Empresa_Id    ='" & Ps.CodigoEmpresa & "'" & vbCrLf &
                        "   and EndEmpresa_Id = " & Ps.EnderecoEmpresa & vbCrLf &
                        "   and Pesagem_Id    = " & Ps.Codigo & vbCrLf &
                        "   and Sequencia_Id  = " & Ps.Sequencia

                Sqls.Add(Sql)

            End If
        Next
    End Sub

    Public Sub DesfazerAgrupamento(ByRef Sqls As ArrayList, ByRef listPesagem As [Lib].Negocio.ListPesagem)
        Dim N As New Numerador(CodigoEmpresa, EnderecoEmpresa, 110)
        Dim objBanco As New AcessaBanco()

        For Each p As [Lib].Negocio.Pesagem In listPesagem
            If p.Observacoes.Length > 14 Then
                p.Observacoes = p.Observacoes.Replace(" - LAUDO AGRUPADO", "")
            Else
                p.Observacoes = p.Observacoes.Replace("LAUDO AGRUPADO", "")
            End If

            p.UsuarioAlteracao = HttpContext.Current.Session("ssNomeUsuario")
            p.DataAlteracao = DateTime.Now

            Dim Sql As String = " Update Pesagem Set" & vbCrLf &
                                "     Situacao                = 1 " & vbCrLf &
                                "    ,Observacoes             ='" & p.Observacoes & "'" & vbCrLf &
                                "    ,RegistroMestre          = 0 " & vbCrLf &
                                "    ,UsuarioAlteracao     = '" & p.UsuarioAlteracao & "'" & vbCrLf &
                                "    ,UsuarioAlteracaoData = '" & CDate(p.DataAlteracao).ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                                "    ,UsuarioCancelamento  = ''" & vbCrLf &
                                "    ,UsuarioCancelamentoData = '" & Now.ToString("yyyy-MM-dd HH:mm:ss") & "'" & vbCrLf &
                                "	Where Empresa_Id    ='" & p.CodigoEmpresa & "'" & vbCrLf &
                                "   and EndEmpresa_Id = " & p.EnderecoEmpresa & vbCrLf &
                                "   and Pesagem_Id    = " & p.Codigo & vbCrLf &
                                "   and Sequencia_Id  = " & p.Sequencia
            Sqls.Add(Sql)

            For Each ro In p.Romaneios
                Sql = " Update Romaneios Set" & vbCrLf & _
                        "     Observacoes ='" & p.Observacoes & "'" & vbCrLf & _
                        "	Where Empresa_Id    ='" & ro.CodigoEmpresa & "'" & vbCrLf & _
                        "     and EndEmpresa_Id = " & ro.EnderecoEmpresa & vbCrLf & _
                        "     and Romaneio_Id   = " & ro.Codigo

                Sqls.Add(Sql)
            Next

            'Dim pRomaneio As New Romaneio()

            'Dim SqlN As String = "exec sp_Numerador '" & p.CodigoEmpresa & "'," & p.EnderecoEmpresa & ",110"
            'Dim dsN As New DataSet
            'dsN = objBanco.ConsultaDataSet(SqlN, "Numerador")

            'pRomaneio.Codigo = dsN.Tables(0).Rows(0).Item(0)

            'pRomaneio.IUD = "I"
            'pRomaneio.CodigoEmpresa = p.CodigoEmpresa
            'pRomaneio.EnderecoEmpresa = p.EnderecoEmpresa

            'pRomaneio.EntradaSaida = p.EntradaSaida
            'pRomaneio.CodigoPedido = p.CodigoPedido

            'If p.CodigoDeposito.Length > 0 Then
            '    pRomaneio.CodigoDeposito = p.CodigoDeposito
            '    pRomaneio.EnderecoDeposito = p.EnderecoDeposito
            'Else
            '    pRomaneio.CodigoDeposito = p.CodigoEmpresa
            '    pRomaneio.EnderecoDeposito = p.EnderecoEmpresa
            'End If

            'If p.EntradaSaida = "E" Then
            '    pRomaneio.CodigoDestino = p.CodigoEmpresa
            '    pRomaneio.EnderecoDestino = p.EnderecoEmpresa
            'Else
            '    pRomaneio.CodigoDestino = p.CodigoCliente
            '    pRomaneio.EnderecoDestino = p.EnderecoCliente
            'End If

            'pRomaneio.CodigoTransbordo = ""
            'pRomaneio.EnderecoTransbordo = 0

            'pRomaneio.CodigoProduto = p.CodigoProduto
            'pRomaneio.CodigoOperacao = p.CodigoOperacao
            'pRomaneio.CodigoSubOperacao = p.CodigoSubOperacao
            'pRomaneio.Movimento = p.Movimento

            'pRomaneio.PesoBruto = p.BrutoBalanca
            'pRomaneio.Desconto = p.BrutoBalanca - p.Liquido
            'pRomaneio.PesoLiquido = p.Liquido

            'pRomaneio.Observacoes = p.Observacoes
            'pRomaneio.Processo = ""

            'pRomaneio.CodigoAutorizacao = p.CodigoAutorizacao

            'pRomaneio.DescontosAnalises = New ListRomaneioXDesconto(pRomaneio)

            'For Each row As PesagemXAnalises In p.Analises
            '    Dim RxD As New RomaneioXDesconto(pRomaneio)
            '    RxD.CodigoAnalise = row.CodigoAnalise
            '    RxD.Desconto = row.Desconto
            '    RxD.Percentual = row.Percentual
            '    RxD.Indice = row.Indice
            '    pRomaneio.DescontosAnalises.Add(RxD)
            'Next

            'pRomaneio.SalvarSql(Sqls, False)

            'Dim rxp As New RomaneioXPesagem()
            'rxp.IUD = "I"
            'rxp.CodigoEmpresa = pRomaneio.CodigoEmpresa
            'rxp.EndEmpresa = pRomaneio.EnderecoEmpresa
            'rxp.CodigoRomaneio = pRomaneio.Codigo
            'rxp.CodigoPesagem = p.Codigo
            'rxp.Sequencia = 0

            'rxp.SalvarSql(Sqls)
        Next
    End Sub

    Function SaldoDePedidos(ByRef Laudo As Negocio.Pesagem, Optional ByVal AnoCorrente As Boolean = True, Optional ByVal Safra As String = "") As DataSet
        Dim ds As New DataSet
        Dim Banco As New AcessaBanco
        Dim sql As String = String.Empty

        sql = "Select P.Empresa_Id AS CodigoEmpresa, " & vbCrLf & _
              "       P.EndEmpresa_Id AS EnderecoEmpresa, " & vbCrLf & _
              "       P.Pedido_Id AS Pedido, " & vbCrLf & _
              "       P.DataPedido, " & vbCrLf & _
              "       P.DataEntrega, " & vbCrLf & _
              "       P.Operacao, " & vbCrLf & _
              "       P.SubOperacao, " & vbCrLf & _
              "       P.Safra, " & vbCrLf & _
              "       sbPed.Produto_Id as Produto, " & vbCrLf & _
              "       prd.Nome as NomeDoProduto, " & vbCrLf & _
              "       sOP.Classe, " & vbCrLf & _
              "       isnull(sbPed.QtdNormal,0) as QtdNormal, " & vbCrLf & _
              "       case when isnull(sbPed.Total,0) = 0 or isnull(sbPed.QtdNormal,0) = 0 " & vbCrLf & _
              "         then 0.00" & vbCrLf & _
              "         else (isnull(sbPed.Total,0) / isnull(sbPed.QtdNormal,0))" & vbCrLf & _
              "       end as Unitario," & vbCrLf & _
              "       case " & vbCrLf & _
              "       when prd.ControlarRomaneio = 0 " & vbCrLf & _
              "             then isnull(sbnf.EntregueNota,0) " & vbCrLf & _
              "             else isnull(sbro.EntregueNota,0) " & vbCrLf & _
              "       end as EntregueNota, " & vbCrLf & _
              "       case " & vbCrLf & _
              "       when prd.ControlarRomaneio = 0 " & vbCrLf & _
              "             then isnull(sbnf.DevolucaoNota,0) " & vbCrLf & _
              "             else isnull(sbro.DevolucaoNota,0) " & vbCrLf & _
              "       end as DevolucaoNota " & vbCrLf & _
              "  Into #Consulta " & vbCrLf & _
              "  from Pedidos P " & vbCrLf & _
              " inner Join (Select Pxi.Empresa_Id, Pxi.EndEmpresa_Id, Pxi.Pedido_Id, Pxi.Produto_Id, " & vbCrLf & _
              "			           sum(case " & vbCrLf & _
              "					         When Pxi.TipoDeLancamento = 'E' " & vbCrLf & _
              "					           then Pxi.Quantidade * - 1 " & vbCrLf & _
              "					           else Pxi.Quantidade " & vbCrLf & _
              "				             end) as QtdNormal, " & vbCrLf & _
              "                    sum(case" & vbCrLf & _
              "		           			 When Pxi.TipoDeLancamento = 'E'" & vbCrLf & _
              "		           			   then Pxi.TotalOficial * - 1" & vbCrLf & _
              "		           			   else Pxi.TotalOficial" & vbCrLf & _
              "		           		   end) as Total" & vbCrLf & _
              "		          from PedidoxItemXLancamento Pxi " & vbCrLf & _
              "		         inner Join Pedidos ped " & vbCrLf & _
              "		            on Ped.Empresa_Id    = Pxi.Empresa_Id " & vbCrLf & _
              "		           and Ped.EndEmpresa_Id = Pxi.EndEmpresa_Id " & vbCrLf & _
              "		           and Ped.Pedido_Id     = Pxi.Pedido_Id " & vbCrLf & _
              "		         Where ped.Situacao = 1 " & vbCrLf & _
              "		           and isnull(ped.FiscalAberto,1) = 1 " & vbCrLf & _
              "		           and isnull(ped.PedidoBloqueado,0) = 0 " & vbCrLf & _
              "		         Group By Pxi.Empresa_Id, Pxi.EndEmpresa_Id, Pxi.Pedido_Id, Pxi.Produto_Id" & vbCrLf & _
              "             ) as sbPed " & vbCrLf & _
              "		on sbPed.Empresa_Id = P.Empresa_Id " & vbCrLf & _
              "	   and sbPed.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
              "	   and sbPed.Pedido_Id = P.Pedido_Id " & vbCrLf & _
              " inner Join Produtos prd " & vbCrLf & _
              "    on prd.produto_id =  sbPed.Produto_id " & vbCrLf & _
              " INNER JOIN SubOperacoes sOP " & vbCrLf & _
              "    ON P.Operacao    = sOP.Operacao_Id " & vbCrLf & _
              "   AND P.SubOperacao = sOP.SubOperacoes_Id " & vbCrLf & _
              " left join (SELECT R.Empresa_Id, " & vbCrLf & _
              "                   R.EndEmpresa_Id, " & vbCrLf & _
              "                   R.Pedido, " & vbCrLf & _
              "                   R.Produto, " & vbCrLf & _
              "                   sum(case " & vbCrLf & _
              "                           when SO.Devolucao = 'N' " & vbCrLf & _
              "                               then R.PesoLiquido " & vbCrLf & _
              "                               else 0 " & vbCrLf & _
              "                       end) AS EntregueNota, " & vbCrLf & _
              "                       sum(case " & vbCrLf & _
              "                              when SO.Devolucao = 'S' " & vbCrLf & _
              "                                then R.PesoLiquido " & vbCrLf & _
              "                                else 0 " & vbCrLf & _
              "                           end) AS DevolucaoNota " & vbCrLf & _
              "              FROM Romaneios R " & vbCrLf & _
              "             INNER JOIN SubOperacoes SO " & vbCrLf & _
              "                ON R.Operacao    = SO.Operacao_Id " & vbCrLf & _
              "               AND R.SubOperacao = SO.SubOperacoes_Id " & vbCrLf & _
              "               AND SO.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "') " & vbCrLf & _
              "              LEFT JOIN RomaneiosXPesagens RxP" & vbCrLf & _
              "	               ON Rxp.Empresa_Id    = R.Empresa_Id" & vbCrLf & _
              "		          AND RxP.EndEmpresa_Id = R.EndEmpresa_Id" & vbCrLf & _
              "		          AND RxP.Romaneio_Id   = R.Romaneio_Id" & vbCrLf & _
              "	             LEFT JOIN Pesagem P" & vbCrLf & _
              "	               ON RxP.Empresa_Id    = P.Empresa_Id" & vbCrLf & _
              "		          AND RxP.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
              " 		      AND RxP.Pesagem_Id    = P.Pesagem_Id" & vbCrLf & _
              "              LEFT JOIN NotasFiscaisXRomaneios NFxR" & vbCrLf & _
              "                ON R.Empresa_Id    = NFxR.Empresa_Id" & vbCrLf & _
              "               AND R.EndEmpresa_Id = NFxR.EndEmpresa_Id" & vbCrLf & _
              "               AND R.Romaneio_id   = NFxR.Romaneio_Id" & vbCrLf & _
              "              LEFT JOIN NotasFiscais NF" & vbCrLf & _
              "                ON NFxR.Empresa_Id      = NF.Empresa_Id " & vbCrLf & _
              "		          AND NFxR.EndEmpresa_Id   = NF.EndEmpresa_Id " & vbCrLf & _
              "		          AND NFxR.Cliente_Id      = NF.Cliente_Id " & vbCrLf & _
              "		          AND NFxR.EndCliente_Id   = NF.EndCliente_Id " & vbCrLf & _
              "		          AND NFxR.EntradaSaida_Id = NF.EntradaSaida_Id " & vbCrLf & _
              "		          AND NFxR.Serie_Id        = NF.Serie_Id " & vbCrLf & _
              "		          AND NFxR.Nota_Id         = NF.Nota_Id" & vbCrLf & _
              "             WHERE ISNULL(P.Situacao, NF.Situacao) IN (1, 4, 7) " & vbCrLf & _
              "             GROUP BY R.Empresa_Id, " & vbCrLf & _
              "                      R.EndEmpresa_Id, " & vbCrLf & _
              "                      R.Pedido, " & vbCrLf & _
              "                      R.Produto " & vbCrLf & _
              "             ) sbro " & vbCrLf & _
              "    on sbro.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
              "   and sbro.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
              "   and sbro.Pedido        = P.Pedido_Id " & vbCrLf & _
              "   and sbro.Produto       = sbPed.Produto_id " & vbCrLf & _
              " left join (SELECT nXi.Empresa_Id, " & vbCrLf & _
              "        nXi.EndEmpresa_Id, " & vbCrLf & _
              "         nXi.Pedido, " & vbCrLf & _
              "         nXi.Produto_Id,  " & vbCrLf & _
              "         sum(case " & vbCrLf & _
              "                when SO.Devolucao = 'N' " & vbCrLf & _
              "                  then nXi.QuantidadeFiscal " & vbCrLf & _
              "                  else 0 " & vbCrLf & _
              "             end) AS EntregueNota, " & vbCrLf & _
              "         sum(case " & vbCrLf & _
              "                when SO.Devolucao = 'S' " & vbCrLf & _
              "                  then nXi.QuantidadeFiscal " & vbCrLf & _
              "                  else 0 " & vbCrLf & _
              "             end) AS DevolucaoNota " & vbCrLf & _
              "    FROM NotasFiscaisXItens nXi " & vbCrLf & _
              "    INNER JOIN SubOperacoes SO " & vbCrLf & _
              "      ON nXi.Operacao    = SO.Operacao_Id " & vbCrLf & _
              "     AND nXI.SubOperacao = SO.SubOperacoes_Id " & vbCrLf & _
              "     And SO.Classe not in ('" & eClassesOperacoes.GLOBAL.ToString & "','" & eClassesOperacoes.COMPLEMENTACOES.ToString & "') " & vbCrLf & _
              "   group by nXi.Empresa_Id, " & vbCrLf & _
              "nXi.EndEmpresa_Id, " & vbCrLf & _
              "nXi.Pedido, " & vbCrLf & _
              "nXi.Produto_Id " & vbCrLf & _
              ") sbnf " & vbCrLf & _
              "on sbnf.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
              "and sbnf.EndEmpresa_Id = P.EndEmpresa_Id " & vbCrLf & _
              "and sbnf.Pedido        = P.Pedido_Id " & vbCrLf & _
              "and sbnf.Produto_Id    = sbPed.Produto_id   " & vbCrLf & _
              " where P.Situacao = 1" & vbCrLf & _
              "   and P.Empresa_Id    ='" & Laudo.CodigoEmpresa & "'" & vbCrLf & _
              "   and P.EndEmpresa_Id = " & Laudo.EnderecoEmpresa & vbCrLf
        If Laudo.Consolidado Then
            If Laudo.CodigoCliente.Length.Equals(11) Then
                sql &= "  AND (P.Cliente = '" & Laudo.CodigoCliente & "') " & vbCrLf
            Else
                sql &= "  AND (P.Cliente like '" & Left(Laudo.CodigoCliente, 8) & "%') " & vbCrLf
            End If
        Else
            If Laudo.CodigoCliente.Length > 0 Then
                sql &= "  AND (P.Cliente = '" & Laudo.CodigoCliente & "') " & vbCrLf
                sql &= "  AND (P.EndCliente = " & Laudo.EnderecoCliente & ") " & vbCrLf
            End If
        End If

        If Laudo.CodigoPedido > 0 Then
            sql &= "  AND (P.Pedido_Id = " & Laudo.CodigoPedido & ") " & vbCrLf
        ElseIf Safra.Length > 0 Then
            sql &= "  AND (P.Safra = '" & Safra & "') " & vbCrLf
        ElseIf (AnoCorrente) Then
            sql &= " AND YEAR(P.DataPedido) = " & Date.Now.Year & vbCrLf
        End If

        sql &= "  AND (isnull(P.FiscalAberto,1) = 1) " & vbCrLf & _
               "  AND (isnull(P.PedidoBloqueado,0) = 0) " & vbCrLf & _
               "  AND (prd.Agrupar = 'N' or prd.ControlarPesagem = 1) " & vbCrLf

        sql &= "  group by P.Empresa_Id, " & vbCrLf & _
               "          P.EndEmpresa_Id, " & vbCrLf & _
               "          P.Pedido_Id, " & vbCrLf & _
               "		  P.DataPedido, " & vbCrLf & _
               "		  P.DataEntrega, " & vbCrLf & _
               "		  P.Operacao, " & vbCrLf & _
               "		  P.SubOperacao, " & vbCrLf & _
               "		  P.Safra, " & vbCrLf & _
               "          sbPed.Produto_Id, " & vbCrLf & _
               "          prd.ControlarRomaneio, " & vbCrLf & _
               "          prd.Nome, " & vbCrLf & _
               "          sOP.Classe, " & vbCrLf & _
               "          isnull(sbPed.QtdNormal,0), " & vbCrLf & _
               "          isnull(sbPed.Total,0)," & vbCrLf & _
               "          isnull(sbnf.EntregueNota,0), " & vbCrLf & _
               "          isnull(sbnf.DevolucaoNota,0), " & vbCrLf & _
               "          isnull(sbro.EntregueNota,0), " & vbCrLf & _
               "          isnull(sbro.DevolucaoNota, 0) " & vbCrLf

        sql &= " select C.CodigoEmpresa, " & vbCrLf & _
               "        C.EnderecoEmpresa, " & vbCrLf & _
               "        C.Pedido, " & vbCrLf & _
               "        C.DataPedido, " & vbCrLf & _
               "        C.DataEntrega, " & vbCrLf & _
               "        C.Operacao, " & vbCrLf & _
               "        C.SubOperacao, " & vbCrLf & _
               "        SO.Descricao as DescricaoOperacao, " & vbCrLf & _
               "        C.Safra, " & vbCrLf & _
               "        C.Produto, " & vbCrLf & _
               "        C.NomeDoProduto, " & vbCrLf & _
               "        C.QtdNormal As Contratada, " & vbCrLf & _
               "        C.Unitario," & vbCrLf & _
               "        case " & vbCrLf & _
               "            when so.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' " & vbCrLf & _
               "				then C.EntregueNota " & vbCrLf & _
               "				else C.EntregueNota - C.DevolucaoNota " & vbCrLf & _
               "        end as Entregue, " & vbCrLf & _
               "        C.DevolucaoNota AS Devolucao, " & vbCrLf & _
               "        case " & vbCrLf & _
               "            when so.classe = '" & eClassesOperacoes.DEPOSITOS.ToString & "' " & vbCrLf & _
               "				then " & vbCrLf & _
               "                    case " & vbCrLf & _
               "	                    when C.QtdNormal > 0 " & vbCrLf & _
               "		                then C.QtdNormal - C.EntregueNota	" & vbCrLf & _
               "		                else C.EntregueNota " & vbCrLf & _
               "                    end " & vbCrLf & _
               "				else C.QtdNormal - (C.EntregueNota - C.DevolucaoNota) " & vbCrLf & _
               "        end as Saldo, " & vbCrLf & _
               "        (C.EntregueNota - C.DevolucaoNota) as SaldoDevolucao " & vbCrLf & _
               "  from #Consulta C " & vbCrLf & _
               " Inner Join SubOperacoes SO " & vbCrLf & _
               "    ON C.Operacao    = SO.Operacao_Id " & vbCrLf & _
               "   AND C.SubOperacao = SO.SubOperacoes_Id " & vbCrLf

        If Laudo.CodigoPedido = 0 AndAlso Safra.Length = 0 AndAlso Not AnoCorrente Then
            sql &= " where " & vbCrLf & _
                   "    case " & vbCrLf & _
                   "        when SO.Classe IN('" & eClassesOperacoes.AFIXAR.ToString & "','" & eClassesOperacoes.DEPOSITOS.ToString & "') " & vbCrLf & _
                   "            then " & vbCrLf & _
                   "                case " & vbCrLf & _
                   "	                when C.QtdNormal > 0 " & vbCrLf & _
                   "		            then C.QtdNormal - C.DevolucaoNota	" & vbCrLf & _
                   "		            else C.EntregueNota " & vbCrLf & _
                   "                end " & vbCrLf & _
                   "            else " & vbCrLf & _
                   "                C.QtdNormal - (C.EntregueNota - C.DevolucaoNota) " & vbCrLf & _
                   "        end > 0  " & vbCrLf
        End If

        ds = Banco.ConsultaDataSet(sql, "Laudo")
        Return ds

    End Function

    Private Sub AtualizaAtributos()
        Dim Banco As New AcessaBanco()
        Dim Sql As String
        Sql = "SELECT case when Mov.QtdeRomaneio = 0         then 0 else 1 end TemRomaneio," & vbCrLf & _
              "	      case when Mov.QtdeNota     = 0         then 0 else 1 end TemNota, " & vbCrLf & _
              "	      case when Mov.QtdePesagem  > 1         then 1 else 0 end TemAgrupamento," & vbCrLf & _
              "	      case when Mov.QtdeRomaneio > 1         then 1 else 0 end TemRateio," & vbCrLf & _
              "	      case when Mov.QtdeRomaneio <> QtdeNota then 1 else 0 end TemRomaneioSemNota" & vbCrLf & _
              "  FROM Pesagem P" & vbCrLf & _
              " left Join(Select RxP.Empresa_Id, RxP.EndEmpresa_Id, RxP.Pesagem_Id, RxP.sequencia_id," & vbCrLf & _
              "                  case when isnull(agrup.QtdePesagem,0) > 0 then agrup.QtdePesagem else Rateio.QtdePesagem end as QtdePesagem," & vbCrLf & _
              "                  count(*) QtdeRomaneio," & vbCrLf & _
              "				     sum(case when NxR.Nota_Id is null then 0 else 1 end) QtdeNota" & vbCrLf & _
              "             from Pesagem P2" & vbCrLf & _
              "			    inner join RomaneiosXPesagens RxP" & vbCrLf & _
              "			       on rxp.Empresa_Id    = P2.Empresa_Id" & vbCrLf & _
              "               and rxp.EndEmpresa_Id = P2.EndEmpresa_Id" & vbCrLf & _
              "               and rxp.pesagem_id    = p2.Pesagem_Id" & vbCrLf & _
              "              Left Join NotasFiscaisXRomaneios NxR" & vbCrLf & _
              "			       on rxp.Empresa_Id    = nxr.Empresa_Id" & vbCrLf & _
              "			      and rxp.EndEmpresa_Id = nxr.EndEmpresa_Id" & vbCrLf & _
              "			      and rxp.Romaneio_Id   = nxr.Romaneio_Id" & vbCrLf & _
              "              Left join (select p2.Empresa_Id, p2.EndEmpresa_Id, p2.RegistroMestre, count(*) as QtdePesagem" & vbCrLf & _
              "			                  from Pesagem p2" & vbCrLf & _
              "						     Group by p2.Empresa_Id, p2.EndEmpresa_Id, p2.RegistroMestre" & vbCrLf & _
              "			               ) Agrup" & vbCrLf & _
              "                ON P2.Empresa_Id    = Agrup.Empresa_Id" & vbCrLf & _
              "               AND P2.EndEmpresa_Id = Agrup.EndEmpresa_Id" & vbCrLf & _
              "               AND P2.Pesagem_Id    = Agrup.RegistroMestre" & vbCrLf & _
              "              Left join (select p2.Empresa_Id, p2.EndEmpresa_Id, p2.Pesagem_Id, count(*) as QtdePesagem" & vbCrLf & _
              "			                  from Pesagem p2" & vbCrLf & _
              "						     Group by p2.Empresa_Id, p2.EndEmpresa_Id, p2.Pesagem_Id" & vbCrLf & _
              "			               ) Rateio" & vbCrLf & _
              "                ON P2.Empresa_Id    = Rateio.Empresa_Id" & vbCrLf & _
              "               AND P2.EndEmpresa_Id = Rateio.EndEmpresa_Id" & vbCrLf & _
              "               AND P2.Pesagem_Id    = Rateio.Pesagem_Id" & vbCrLf & _
              "             group by RxP.Empresa_Id, RxP.EndEmpresa_Id, RxP.Pesagem_Id, RxP.Sequencia_id," & vbCrLf & _
              "				         case when isnull(agrup.QtdePesagem,0) > 0 then agrup.QtdePesagem else Rateio.QtdePesagem end" & vbCrLf & _
              "           ) as Mov" & vbCrLf & _
              "    ON Mov.Empresa_Id    = P.Empresa_Id " & vbCrLf & _
              "   AND Mov.EndEmpresa_Id = P.EndEmpresa_Id" & vbCrLf & _
              "   AND Mov.Pesagem_Id    = case when isnull(P.RegistroMestre,0) > 0 then P.RegistroMestre else P.Pesagem_Id   end" & vbCrLf & _
              "   AND Mov.Sequencia_Id  = case when isnull(p.RegistroMestre,0) > 0 then 0                else p.Sequencia_Id end" & vbCrLf & _
              " WHERE P.Empresa_Id    ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
              "   AND P.EndEmpresa_Id = " & Me.EnderecoEmpresa & vbCrLf & _
              "   AND P.Pesagem_Id    = " & Me.Codigo & vbCrLf & _
              "   and p.Sequencia_Id  = " & Me.Sequencia

        Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "Consulta")
        Dim row As DataRow = ds.Tables(0).Rows(0)

        Me.TemRomaneio = row("TemRomaneio")
        Me.TemNota = row("TemNota")
        Me.TemAgrupamento = row("TemAgrupamento")
        Me.TemRateio = row("TemRateio")
        Me.TemRomaneioSemNota = row("TemRomaneioSemNota")
    End Sub
#End Region

End Class