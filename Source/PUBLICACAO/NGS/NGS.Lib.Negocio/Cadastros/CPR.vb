Imports Microsoft.VisualBasic
Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class ListCPR
    Inherits List(Of CPR)

#Region "Contrutor"
    Public Sub New(Optional ByVal pCPR As CPR = Nothing, Optional ByVal pWhere As Hashtable = Nothing)
        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String = ""

            Sql = "SELECT C.Cartorio_Id, C.EndCartorio_Id, Cartorio.Nome + ' / ' + Cartorio.cidade + '-' + Cartorio.estado as CartorioCidadeUF, " & vbCrLf & _
                  "       C.Cliente, C.EndCliente, C.CPR_Id, C.Situacao, C.DataEmissao, C.Registro, C.DataRegistro, C.DataVencimento, C.Empresa, " & vbCrLf & _
                  "       C.EndEmpresa, C.Safra, C.Produto, P.Nome as NomeProduto, C.Quantidade, C.Produtividade, C.Endossado, C.EndEndossado, " & vbCrLf & _
                  "       C.Observacao, C.DevedorSolidarioCPR, C.DevedorSolidarioCartorio, C.DevedorSolidarioEndCartorio, " & vbCrLf & _
                  "       ISNULL(C2.CPR_Id, '') CautelaCPR, ISNULL(C2.Cartorio_Id, '') CautelaCartorio, ISNULL(C2.EndCartorio_Id,0) CautelaEndCartorio " & vbCrLf & _
                  "  FROM CPR C" & vbCrLf & _
                  " Inner Join Clientes Cartorio" & vbCrLf & _
                  "    on Cartorio.Cliente_Id  = C.Cartorio_Id" & vbCrLf & _
                  "   and Cartorio.Endereco_Id = C.EndCartorio_Id" & vbCrLf & _
                  " Inner Join Produtos P" & vbCrLf & _
                  "    on P.Produto_id = C.Produto" & vbCrLf & _
                  "  LEFT JOIN CPR C2" & vbCrLf & _
                  "    ON C2.CPR_Id = C.DevedorSolidarioCPR " & vbCrLf & _
                  "   AND C2.Cartorio_Id = C.DevedorSolidarioCartorio " & vbCrLf & _
                  "   AND C2.EndCartorio_Id = C.DevedorSolidarioEndCartorio " & vbCrLf & _
                  " WHERE 1 = 1 "

            If Not pCPR Is Nothing Then
                If pCPR.CodigoEmpresa.Length > 0 Then
                    Sql &= "   and C.Empresa    ='" & pCPR.CodigoEmpresa & "'" & vbCrLf & _
                           "   and C.EndEmpresa = " & pCPR.EndEmpresa
                End If

                If pCPR.CodigoCartorio.Length > 0 Then
                    Sql &= "   and C.Cartorio_Id    ='" & pCPR.CodigoCartorio & "'" & vbCrLf & _
                           "   and C.EndCartorio_Id = " & pCPR.EndCartorio
                End If


                If pCPR.CodigoCliente.Length > 0 Then
                    Sql &= "   and C.Cliente    ='" & pCPR.CodigoCliente & "'" & vbCrLf & _
                           "   and C.EndCliente = " & pCPR.EndCliente
                End If

                If pCPR.CodigoCPR.Length > 0 Then
                    Sql &= "   and C.CPR_Id    ='" & pCPR.CodigoCPR & "'" & vbCrLf
                End If

                If pCPR.CodigoSafra.Length > 0 Then
                    Sql &= "   and C.Safra    ='" & pCPR.CodigoSafra & "'" & vbCrLf
                End If

                If pCPR.CodigoProduto.Length > 0 Then
                    Sql &= "   and C.Produto    ='" & pCPR.CodigoProduto & "'" & vbCrLf
                ElseIf pWhere.ContainsKey("GrupoProduto") Then
                    Sql &= "   and P.Grupo      ='" & pWhere("GrupoProduto") & "'" & vbCrLf
                End If

                If pCPR.CodigoEndossado.Length > 0 Then
                    Sql &= "   and C.Endossado    ='" & pCPR.CodigoEndossado & "'" & vbCrLf & _
                           "   and C.EndEndossado = " & pCPR.EndEndossado
                End If

                If pCPR.Observacao.Length > 0 Then
                    Sql &= "   and C.Observacao like '%" & pCPR.Observacao & "%'" & vbCrLf
                End If

                If pCPR.CodigoSituacao >= 0 Then
                    Sql &= "   and C.Situacao =" & pCPR.CodigoSituacao & vbCrLf
                End If

                If pWhere.ContainsKey("Emissao") Then
                    Dim emi As String() = pWhere("Emissao").ToString.Split("|")
                    Sql &= "   and C.DataEmissao between '" & emi(0) & "' and '" & emi(1) & "'" & vbCrLf
                End If

                If pWhere.ContainsKey("Vencimento") Then
                    Dim venc As String() = pWhere("Vencimento").ToString.Split("|")
                    Sql &= "   and C.DataVencimento between '" & venc(0) & "' and '" & venc(1) & "'" & vbCrLf
                End If

                If pCPR.Registro > 0 Then
                    Sql &= "   and C.Registro = " & pCPR.Registro
                End If

                If pCPR.Quantidade > 0 Then
                    Sql &= "   and C.Quantidade = " & Str(pCPR.Quantidade)
                End If

                If pCPR.Produtividade > 0 Then
                    Sql &= "   and C.Produtividade = " & Str(pCPR.Produtividade)
                End If

                If pWhere.ContainsKey("Proprietario") Then
                    Dim Prop As String() = pWhere("Proprietario").ToString.Split("-")
                    Sql &= "   and exists(Select 1 from CPRxCliente CC where CC.Cartorio_id = C.Cartorio_Id and CC.EndCartorio_id = C.EndCartorio_id and CC.CPR_Id = C.CPR_Id and " & IIf(pWhere("Consolidar"), "left(CC.Cliente_Id,8) = '" & Prop(0).Substring(0, 8) & "'", "CC.Cliente_id ='" & Prop(0) & "' and CC.EndCliente_id =" & Prop(1)) & ")"
                End If

                If pWhere.ContainsKey("Fazenda") Then
                    Dim Faz As String() = pWhere("Fazenda").ToString.Split("-")
                    Sql &= "   and exists(Select 1 from CPRxFazenda CF where CF.Cartorio_id = C.Cartorio_Id and CF.EndCartorio_id = C.EndCartorio_id and CF.CPR_Id = C.CPR_Id and CF.Fazenda_id = '" & Faz(0) & "' and EndFazenda_Id =" & Faz(1) & ")"
                End If

                If pWhere.ContainsKey("Matricula") Then
                    Sql &= "   and exists(Select 1 from CPRxMatricula CM where CM.Cartorio_id = C.Cartorio_Id and CM.EndCartorio_id = C.EndCartorio_id and CM.CPR_Id = C.CPR_Id and CM.Matricula_id in (" & pWhere("Matricula") & "))"
                End If


                If pWhere.ContainsKey("Grau") Then
                    Sql &= "   and exists(Select 1 from CPRxGrau CG where CG.Cartorio_id = C.Cartorio_Id and CG.EndCartorio_id = C.EndCartorio_id and CG.CPR_Id = C.CPR_Id and CG.Cliente = C.Empresa and CG.EndCliente = C.EndEmpresa and CG.Grau_Id =" & pWhere("Grau") & ")"
                End If

                If pWhere.ContainsKey("ExcessaoCodigoCPR") And pWhere.ContainsKey("ExcessaoCodigoCartorio") And pWhere.ContainsKey("ExcessaoEndCartorio") Then
                    Sql &= "   and not (C.Cpr_id =  '" & pWhere("ExcessaoCodigoCPR") & "'" & vbCrLf
                    Sql &= "   and C.Cartorio_id =  '" & pWhere("ExcessaoCodigoCartorio") & "'" & vbCrLf
                    Sql &= "   and C.EndCartorio_id =  " & pWhere("ExcessaoEndCartorio") & ")"

                End If



            End If

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPR")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim cp As New CPR()

                cp.CodigoCartorio = row("Cartorio_Id")
                cp.EndCartorio = row("EndCartorio_Id")
                cp.CartorioCidadeUF = row("CartorioCidadeUF")
                cp.CodigoCliente = row("Cliente")
                cp.EndCliente = row("EndCliente")
                cp.CodigoCPR = row("CPR_Id")
                cp.CodigoSituacao = row("Situacao")
                cp.DataEmissao = row("DataEmissao")
                cp.Registro = row("Registro")
                cp.DataRegistro = row("DataRegistro")
                cp.DataVencimento = row("DataVencimento")
                cp.CodigoEmpresa = row("Empresa")
                cp.EndEmpresa = row("EndEmpresa")
                cp.CodigoSafra = row("Safra")
                cp.CodigoProduto = row("Produto")
                cp.NomeProduto = row("NomeProduto")
                cp.Quantidade = row("Quantidade")
                cp.Produtividade = row("Produtividade")
                cp.CodigoEndossado = row("Endossado")
                cp.EndEndossado = row("EndEndossado")
                cp.Observacao = row("Observacao")
                cp.DevedorSolidarioCodigoCPR = row("DevedorSolidarioCPR")
                cp.DevedorSolidarioCodigoCartorio = row("DevedorSolidarioCartorio")
                cp.DevedorSolidarioEndCartorio = row("DevedorSolidarioEndCartorio")
                cp.DevedorSolidarioEndCartorio = row("DevedorSolidarioEndCartorio")
                cp.CautelaCodigoCPR = row("CautelaCPR")
                cp.CautelaCodigoCartorio = row("CautelaCartorio")
                cp.CautelaEndCartorio = row("CautelaEndCartorio")
                Me.Add(cp)
            Next

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Public Erro As Exception
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
        For Each cp As CPR In Me
            If cp.IUD <> "" Then
                cp.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

<Serializable()> _
Public Class CPR
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal pCartorio As String, ByVal pEndCartorio As Integer, ByVal pCPR As String)
        Dim Banco As New AcessaBanco()

        Try
            Dim Sql As String

            Sql = "SELECT C.Cartorio_Id, C.EndCartorio_Id, Cartorio.Nome + ' / ' + Cartorio.cidade + '-' + Cartorio.estado as CartorioCidadeUF, C.CPR_Id, " & vbCrLf & _
                  "       C.Situacao, C.DataEmissao, C.Registro, C.DataRegistro, C.DataVencimento, C.Empresa, " & vbCrLf &
                  "       C.EndEmpresa, C.Safra, C.Produto, P.Nome as NomeProduto, C.Quantidade, C.Produtividade, C.Endossado, C.EndEndossado, C.Observacao," & vbCrLf & _
                  "       C.Cliente, C.EndCliente, C.DevedorSolidarioCPR, C.DevedorSolidarioCartorio, C.DevedorSolidarioCartorio, " & vbCrLf & _
                  "       ISNULL(C2.CPR_Id, '') CautelaCPR, ISNULL(C2.Cartorio_Id, '') CautelaCartorio, ISNULL(C2.EndCartorio_Id,0) CautelaEndCartorio " & vbCrLf & _
                  "  FROM CPR C" & vbCrLf & _
                  " Inner Join Clientes Cartorio" & vbCrLf & _
                  "    on Cartorio.Cliente_Id  = C.Cartorio_Id" & vbCrLf & _
                  "   and Cartorio.Endereco_Id = C.EndCartorio_Id" & vbCrLf & _
                  " Inner Join Produtos P" & vbCrLf & _
                  "    on P.Produto_id = C.Produto" & vbCrLf & _
                  "  LEFT JOIN CPR C2" & vbCrLf & _
                  "    ON C2.CPR_Id = C.DevedorSolidarioCPR " & vbCrLf & _
                  "   AND C2.Cartorio_Id = C.DevedorSolidarioCartorio " & vbCrLf & _
                  "   AND C2.EndCartorio_Id = C.DevedorSolidarioEndCartorio " & vbCrLf & _
                  " WHERE C.Cartorio_Id    ='" & pCartorio & "'" & vbCrLf & _
                  "   AND C.EndCartorio_Id = " & pEndCartorio & vbCrLf & _
                  "   AND C.CPR_Id         ='" & pCPR & "'"

            Dim ds As DataSet = Banco.ConsultaDataSet(Sql, "CPR")
            If ds.Tables(0).Rows.Count = 0 Then Exit Sub
            Dim row As DataRow = ds.Tables(0).Rows(0)

            _CodigoCartorio = row("Cartorio_Id")
            _EndCartorio = row("EndCartorio_Id")
            _CartorioCidadeUF = row("CartorioCidadeUF")
            _CodigoCPR = row("CPR_Id")
            _CodigoSituacao = row("Situacao")
            _DataEmissao = row("DataEmissao")
            _Registro = row("Registro")
            _DataRegistro = row("DataRegistro")
            _DataVencimento = row("DataVencimento")
            _CodigoEmpresa = row("Empresa")
            _EndEmpresa = row("EndEmpresa")
            _CodigoSafra = row("Safra")
            _CodigoProduto = row("Produto")
            _NomeProduto = row("NomeProduto")
            _Quantidade = row("Quantidade")
            _Produtividade = row("Produtividade")
            _CodigoEndossado = row("Endossado")
            _EndEndossado = row("EndEndossado")
            _CodigoCliente = row("Cliente")
            _EndCliente = row("EndCliente")
            _DevedorSolidarioCodigoCPR = row("DevedorSolidarioCPR")
            _DevedorSolidarioCodigoCartorio = row("DevedorSolidarioCartorio")
            _DevedorSolidarioEndCartorio = row("DevedorSolidarioEndCartorio")
            _CautelaCodigoCPR = row("CautelaCPR")
            _CautelaCodigoCartorio = row("CautelaCartorio")
            _CautelaEndCartorio = row("CautelaEndCartorio")

        Catch ex As Exception
            Me.Erro = ex
        Finally
            Banco = Nothing
        End Try
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String

    Private _CodigoCartorio As String = ""
    Private _EndCartorio As Integer
    Private _Cartorio As Cliente
    Private _CartorioCidadeUF As String = ""

    Private _CodigoCliente As String = ""
    Private _EndCliente As Integer
    Private _Cliente As Cliente


    Private _CodigoCPR As String = ""

    Private _CodigoSituacao As Integer = -1
    Private _Situacao As Situacao

    Private _DataEmissao As Date
    Private _Registro As Integer
    Private _DataRegistro As Date
    Private _DataVencimento As Date

    Private _CodigoEmpresa As String = ""
    Private _EndEmpresa As Integer
    Private _Empresa As Cliente

    Private _CodigoSafra As String = ""
    Private _Safra As Safra

    Private _CodigoProduto As String = ""
    Private _NomeProduto As String = ""
    Private _Produto As Produto

    Private _Quantidade As Decimal
    Private _Produtividade As Decimal

    Private _CodigoEndossado As String = ""
    Private _EndEndossado As Integer
    Private _Endossado As Cliente

    Private _Observacao As String = ""

    Private _DevedorSolidarioCodigoCPR As String = String.Empty
    Private _DevedorSolidarioCPRCautela As CPR
    Private _DevedorSolidarioCodigoCartorio As String = String.Empty
    Private _DevedorSolidarioCartorio As Cliente
    Private _DevedorSolidarioEndCartorio As Integer
    Private _CautelaCodigoCPR As String = String.Empty
    Private _CautelaCPR As CPR
    Private _CautelaCodigoCartorio As String = String.Empty
    Private _CautelaEndCartorio As Integer

    Private _Fazendas As ListCPRxFazenda
    Private _Avalistas As ListCPRxAvalista
    Private _Graus As ListCPRxGrau
    Private _Liquidacoes As ListCPRxLiquidacao

#End Region

#Region "Property"
    Public Erro As Exception

    Public Property IUD() As String
        Get
            Return _IUD
        End Get
        Set(ByVal value As String)
            _IUD = value
        End Set
    End Property

    Public Property CodigoCartorio() As String
        Get
            Return _CodigoCartorio
        End Get
        Set(ByVal value As String)
            _CodigoCartorio = value
            _Cartorio = Nothing
        End Set
    End Property

    Public Property EndCartorio() As Integer
        Get
            Return _EndCartorio
        End Get
        Set(ByVal value As Integer)
            _EndCartorio = value
            _Cartorio = Nothing
        End Set
    End Property

    Public Property Cartorio() As Cliente
        Get
            If _Cartorio Is Nothing And _CodigoCartorio.Length > 0 Then _Cartorio = New Cliente(_CodigoCartorio, _EndCartorio)
            Return _Cartorio
        End Get
        Set(ByVal value As Cliente)
            _Cartorio = value
        End Set
    End Property

    Public Property CartorioCidadeUF() As String
        Get
            If _CartorioCidadeUF.Length = 0 And CodigoCartorio.Length > 0 Then _CartorioCidadeUF = Cartorio.Nome + " / " + Cartorio.Cidade + "-" + Cartorio.CodigoEstado
            Return _CartorioCidadeUF
        End Get
        Set(ByVal value As String)
            _CartorioCidadeUF = value
        End Set
    End Property

    Public Property CodigoCliente() As String
        Get
            Return _CodigoCliente
        End Get
        Set(ByVal value As String)
            _CodigoCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property EndCliente() As Integer
        Get
            Return _EndCliente
        End Get
        Set(ByVal value As Integer)
            _EndCliente = value
            _Cliente = Nothing
        End Set
    End Property

    Public Property Cliente() As Cliente
        Get
            If _Cliente Is Nothing And _CodigoCliente.Length > 0 Then _Cliente = New Cliente(_CodigoCliente, _EndCliente)
            Return _Cliente
        End Get
        Set(ByVal value As Cliente)
            _Cliente = value
        End Set
    End Property


    Public Property CodigoCPR() As String
        Get
            Return _CodigoCPR
        End Get
        Set(ByVal value As String)
            _CodigoCPR = value
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

    Public ReadOnly Property DescricaoSituacao() As String
        Get
            Select Case _CodigoSituacao
                Case 0 : Return "LIQUIDADA"
                Case 1 : Return "ABERTA"
                Case Else : Return ""
            End Select
        End Get
    End Property

    Public Property DataEmissao() As Date
        Get
            Return _DataEmissao
        End Get
        Set(ByVal value As Date)
            _DataEmissao = value
        End Set
    End Property

    Public Property Registro() As Integer
        Get
            Return _Registro
        End Get
        Set(ByVal value As Integer)
            _Registro = value
        End Set
    End Property

    Public Property DataRegistro() As Date
        Get
            Return _DataRegistro
        End Get
        Set(ByVal value As Date)
            _DataRegistro = value
        End Set
    End Property

    Public Property DataVencimento() As Date
        Get
            Return _DataVencimento
        End Get
        Set(ByVal value As Date)
            _DataVencimento = value
        End Set
    End Property

    Public Property CodigoEmpresa() As String
        Get
            Return _CodigoEmpresa
        End Get
        Set(ByVal value As String)
            _CodigoEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property EndEmpresa() As Integer
        Get
            Return _EndEmpresa
        End Get
        Set(ByVal value As Integer)
            _EndEmpresa = value
            _Empresa = Nothing
        End Set
    End Property

    Public Property Empresa() As Cliente
        Get
            If _Empresa Is Nothing And _CodigoEmpresa.Length > 0 Then _Empresa = New Cliente(_CodigoEmpresa, _EndEmpresa)
            Return _Empresa
        End Get
        Set(ByVal value As Cliente)
            _Empresa = value
        End Set
    End Property

    Public Property CodigoSafra() As String
        Get
            Return _CodigoSafra
        End Get
        Set(ByVal value As String)
            _CodigoSafra = value
            _Safra = Nothing
        End Set
    End Property

    Public Property Safra() As Safra
        Get
            If _Safra Is Nothing And _CodigoSafra.Length > 0 Then _Safra = New Safra(_CodigoSafra)
            Return _Safra
        End Get
        Set(ByVal value As Safra)
            _Safra = value
        End Set
    End Property

    Public Property CodigoProduto() As String
        Get
            Return _CodigoProduto
        End Get
        Set(ByVal value As String)
            _CodigoProduto = value
            _Produto = Nothing
        End Set
    End Property

    Public Property Produto() As Produto
        Get
            If _Produto Is Nothing And _CodigoProduto.Length > 0 Then _Produto = New Produto(_CodigoProduto)
            Return _Produto
        End Get
        Set(ByVal value As Produto)
            _Produto = value
        End Set
    End Property

    Public Property NomeProduto() As String
        Get
            If _NomeProduto.Length = 0 And CodigoProduto.Length > 0 Then _NomeProduto = Produto.Nome
            Return _NomeProduto
        End Get
        Set(ByVal value As String)
            _NomeProduto = value
        End Set
    End Property


    Public Property Quantidade() As Decimal
        Get
            Return _Quantidade
        End Get
        Set(ByVal value As Decimal)
            _Quantidade = value
        End Set
    End Property

    Public Property Produtividade() As Decimal
        Get
            Return _Produtividade
        End Get
        Set(ByVal value As Decimal)
            _Produtividade = value
        End Set
    End Property

    Public Property CodigoEndossado() As String
        Get
            Return _CodigoEndossado
        End Get
        Set(ByVal value As String)
            _CodigoEndossado = value
            _Endossado = Nothing
        End Set
    End Property

    Public Property EndEndossado() As Integer
        Get
            Return _EndEndossado
        End Get
        Set(ByVal value As Integer)
            _EndEndossado = value
        End Set
    End Property

    Public Property Endossado() As Cliente
        Get
            If _Endossado Is Nothing And _CodigoEndossado.Length > 0 Then _Endossado = New Cliente(_CodigoEndossado, _EndEndossado)
            Return _Endossado
        End Get
        Set(ByVal value As Cliente)
            _Endossado = value
        End Set
    End Property

    Public Property Observacao() As String
        Get
            Return _Observacao
        End Get
        Set(ByVal value As String)
            _Observacao = value
        End Set
    End Property

    Public Property DevedorSolidarioCodigoCPR() As String
        Get
            Return _DevedorSolidarioCodigoCPR
        End Get
        Set(ByVal value As String)
            _DevedorSolidarioCodigoCPR = value
        End Set
    End Property

    Public Property DevedorSolidarioCPRCautela() As CPR
        Get
            If _DevedorSolidarioCPRCautela Is Nothing And DevedorSolidarioCodigoCartorio.Length > 0 And DevedorSolidarioEndCartorio >= 0 And Me.DevedorSolidarioCodigoCPR.Length > 0 Then
                _DevedorSolidarioCPRCautela = New CPR(_DevedorSolidarioCodigoCartorio, _DevedorSolidarioEndCartorio, _DevedorSolidarioCodigoCPR)
            End If
            Return _DevedorSolidarioCPRCautela
        End Get
        Set(ByVal value As CPR)
            _DevedorSolidarioCPRCautela = value
        End Set
    End Property


    Public Property DevedorSolidarioCodigoCartorio() As String
        Get
            Return _DevedorSolidarioCodigoCartorio
        End Get
        Set(ByVal value As String)
            _DevedorSolidarioCodigoCartorio = value
            _DevedorSolidarioCartorio = Nothing
        End Set
    End Property

    Public Property DevedorSolidarioCartorio() As Cliente
        Get
            If _DevedorSolidarioCartorio Is Nothing And _DevedorSolidarioCodigoCartorio.Length > 0 Then _DevedorSolidarioCartorio = New Cliente(_DevedorSolidarioCodigoCartorio, _DevedorSolidarioEndCartorio)
            Return _DevedorSolidarioCartorio
        End Get
        Set(ByVal value As Cliente)
            _DevedorSolidarioCartorio = value
        End Set
    End Property

    Public Property DevedorSolidarioEndCartorio() As Integer
        Get
            Return _DevedorSolidarioEndCartorio
        End Get
        Set(ByVal value As Integer)
            _DevedorSolidarioEndCartorio = value
            _DevedorSolidarioCartorio = Nothing
        End Set
    End Property

    Public Property CautelaCodigoCPR() As String
        Get
            Return _CautelaCodigoCPR
        End Get
        Set(ByVal value As String)
            _CautelaCodigoCPR = value
        End Set
    End Property

    Public Property CautelaCPR() As CPR
        Get
            If _CautelaCPR Is Nothing And CautelaCodigoCartorio.Length > 0 And CautelaEndCartorio >= 0 And Me.CautelaCodigoCPR.Length > 0 Then
                _CautelaCPR = New CPR(_CautelaCodigoCartorio, _CautelaEndCartorio, _CautelaCodigoCPR)
            End If
            Return _CautelaCPR
        End Get
        Set(ByVal value As CPR)
            _CautelaCPR = value
        End Set
    End Property



    Public Property CautelaCodigoCartorio() As String
        Get
            Return _CautelaCodigoCartorio
        End Get
        Set(ByVal value As String)
            _CautelaCodigoCartorio = value
        End Set
    End Property



    Public Property CautelaEndCartorio As Integer
        Get
            Return _CautelaEndCartorio
        End Get
        Set(ByVal value As Integer)
            _CautelaEndCartorio = value
        End Set
    End Property


    Public Property Fazendas() As ListCPRxFazenda
        Get
            If _Fazendas Is Nothing Then _Fazendas = New ListCPRxFazenda(Me)
            Return _Fazendas
        End Get
        Set(ByVal value As ListCPRxFazenda)
            _Fazendas = value
        End Set
    End Property

    Public Property Avalistas() As ListCPRxAvalista
        Get
            If _Avalistas Is Nothing Then _Avalistas = New ListCPRxAvalista(Me)
            Return _Avalistas
        End Get
        Set(ByVal value As ListCPRxAvalista)
            _Avalistas = value
        End Set
    End Property

    Public Property Graus() As ListCPRxGrau
        Get
            If _Graus Is Nothing Then _Graus = New ListCPRxGrau(Me)
            Return _Graus
        End Get
        Set(ByVal value As ListCPRxGrau)
            _Graus = value
        End Set
    End Property

    Public Property Liquidacoes() As ListCPRxLiquidacao
        Get
            If _Liquidacoes Is Nothing Then _Liquidacoes = New ListCPRxLiquidacao(Me)
            Return _Liquidacoes
        End Get
        Set(ByVal value As ListCPRxLiquidacao)
            _Liquidacoes = value
        End Set
    End Property



    Public ReadOnly Property Area() As Decimal
        Get
            Dim x As Decimal
            For Each CxF As CPRxFazenda In Me.Fazendas
                x += CxF.Area
            Next
            Return x
        End Get
    End Property


    Public ReadOnly Property ProdutividadeCalculada() As Decimal
        Get
            If Quantidade > 0 And Area > 0 Then
                Return Math.Round(Quantidade / Area / 60, 2)
            Else
                Return 0
            End If
        End Get
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

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = " Insert Into CPR(Cartorio_Id, EndCartorio_Id, Cliente, EndCliente, CPR_Id, Situacao, DataEmissao, Registro, DataRegistro, DataVencimento," & vbCrLf & _
                      "                 Empresa, EndEmpresa, Safra, Produto, Quantidade, Produtividade, Endossado, EndEndossado, Observacao, DevedorSolidarioCPR,  " & vbCrLf & _
                      "                 DevedorSolidarioCartorio, DevedorSolidarioEndCartorio)" & vbCrLf & _
                      " Values ('" & Me.CodigoCartorio & "'," & Me.EndCartorio & ",'" & Me.CodigoCliente & "'," & Me.EndCliente & ", '" & Me.CodigoCPR & "'," & Me.CodigoSituacao & "," & vbCrLf & _
                      " '" & Me.DataEmissao.ToString("yyyy-MM-dd") & "'," & Me.Registro & ",'" & Me.DataRegistro.ToString("yyyy-MM-dd") & "','" & Me.DataVencimento.ToString("yyyy-MM-dd") & "'," & vbCrLf & _
                      "'" & Me.CodigoEmpresa & "'," & Me.EndEmpresa & ",'" & Me.CodigoSafra & "','" & Me.CodigoProduto & "'," & Str(Me.Quantidade) & "," & Str(Me.Produtividade) & ",'" & Me.CodigoEndossado & "'," & Me.EndEndossado & ",'" & _Observacao & "'," & vbCrLf & _
                      "'" & Me.DevedorSolidarioCodigoCPR & "', '" & Me.DevedorSolidarioCodigoCartorio & "', " & Me.DevedorSolidarioEndCartorio & ")"

                Sqls.Add(sql)
                '***********************************************************************
                '********* Procedimento para Salvar as tabelas relacionadas   **********
                '***********************************************************************
                SalvarTabelasRelacionadasSql(Sqls)
            Case "U"
                sql = " Update CPR set " & vbCrLf & _
                      "    Situacao       = " & Me.CodigoSituacao & vbCrLf & _
                      "   ,DataEmissao    ='" & Me.DataEmissao.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Registro       = " & Me.Registro & vbCrLf & _
                      "   ,DataRegistro   ='" & Me.DataRegistro.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,DataVencimento ='" & Me.DataVencimento.ToString("yyyy-MM-dd") & "'" & vbCrLf & _
                      "   ,Empresa        ='" & Me.CodigoEmpresa & "'" & vbCrLf & _
                      "   ,EndEmpresa     = " & Me.EndEmpresa & vbCrLf & _
                      "   ,Cliente        ='" & Me.CodigoCliente & "'" & vbCrLf & _
                      "   ,EndCliente     = " & Me.EndCliente & vbCrLf & _
                      "   ,Safra          ='" & Me.CodigoSafra & "'" & vbCrLf & _
                      "   ,Produto        ='" & Me.CodigoProduto & "'" & vbCrLf & _
                      "   ,Quantidade     = " & Str(Me.Quantidade) & vbCrLf & _
                      "   ,Produtividade  = " & Str(Me.Produtividade) & vbCrLf & _
                      "   ,Endossado      = '" & CodigoEndossado & "'" & vbCrLf & _
                      "   ,EndEndossado   = " & Me.EndEndossado & vbCrLf & _
                      "   ,Observacao     ='" & Me.Observacao & "'" & vbCrLf & _
                      "   ,DevedorSolidarioCPR ='" & Me.DevedorSolidarioCodigoCPR & "'" & vbCrLf & _
                      "   ,DevedorSolidarioCartorio ='" & Me.DevedorSolidarioCodigoCartorio & "'" & vbCrLf & _
                      "   ,DevedorSolidarioEndCartorio =" & Me.DevedorSolidarioEndCartorio & "" & vbCrLf & _
                      "	Where Cartorio_Id    ='" & Me.CodigoCartorio & "'" & vbCrLf & _
                      "   and EndCartorio_Id = " & Me.EndCartorio & vbCrLf & _
                      "   and CPR_Id         ='" & Me.CodigoCPR & "'"
                Sqls.Add(sql)

                'Ao se liquidar uma CPR, é verificada se a mesma têm outra CPR como Cautela e caso tenha esta também 
                'terá sua situação alterada para liquidada.
                If Not String.IsNullOrWhiteSpace(Me.CautelaCodigoCPR) AndAlso Me.CodigoSituacao = 0 Then
                    sql = " UPDATE CPR " & vbCrLf & _
                          "    SET Situacao       = " & Me.CodigoSituacao & vbCrLf & _
                          "	 WHERE Cartorio_Id    ='" & Me.CautelaCodigoCartorio & "'" & vbCrLf & _
                          "    AND EndCartorio_Id = " & Me.CautelaEndCartorio & vbCrLf & _
                          "    AND CPR_Id         ='" & Me.CautelaCodigoCPR & "'"
                    Sqls.Add(sql)
                End If

                SalvarTabelasRelacionadasSql(Sqls)

            Case "D"
                SalvarTabelasRelacionadasSql(Sqls)
                sql = " Delete CPR" & vbCrLf & _
                      "	 Where Cartorio_Id    ='" & Me.CodigoCartorio & "'" & vbCrLf & _
                      "    and EndCartorio_Id = " & Me.EndCartorio & vbCrLf & _
                      "    and CPR_Id         ='" & Me.CodigoCPR & "'"
                Sqls.Add(sql)
        End Select

    End Sub

    Private Sub SalvarTabelasRelacionadasSql(ByRef Sqls As ArrayList)
        If Not Fazendas Is Nothing Then Fazendas.SalvarSql(Sqls)
        If Not Avalistas Is Nothing Then Avalistas.SalvarSql(Sqls)
        If Not Graus Is Nothing Then Graus.SalvarSql(Sqls)
        If Not Liquidacoes Is Nothing Then Liquidacoes.SalvarSql(Sqls)
    End Sub
#End Region

End Class