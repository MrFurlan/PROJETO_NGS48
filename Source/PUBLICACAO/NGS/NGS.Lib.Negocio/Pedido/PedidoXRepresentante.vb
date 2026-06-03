Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'*****************************************************************************************************************
'************************************** List Representante do Pedido   *******************************************
'*****************************************************************************************************************
Public Class ListPedidoXRepresentante
    Inherits List(Of PedidoXRepresentante)

#Region "Construtor"

    Public Sub New()

    End Sub

    Public Sub New(ByVal Parent As Pedido)
        _Pedido = Parent
        If _Pedido.Codigo = 0 Then Exit Sub

        Dim objBanco As New AcessaBanco()

        Try
            Dim sql As String
            sql = "SELECT Representante_Id, EndRepresentante_Id, CONVERT(DECIMAL(18,10), Percentual) AS Percentual, ValorComissao, Principal, isnull(PercentualFixo,case when valorcomissao > 0 then 1 else 0 end) as PercentualFixo" & vbCrLf &
                  "  FROM Comissoes " & vbCrLf &
                  " WHERE Empresa_Id    ='" & Me._Pedido.CodigoEmpresa & "'" & vbCrLf &
                  "   AND EndEmpresa_Id = " & Me._Pedido.EnderecoEmpresa & vbCrLf &
                  "   AND Pedido_Id     = " & Me._Pedido.Codigo & vbCrLf

            Dim ds As DataSet = objBanco.ConsultaDataSet(sql, "Comissoes")

            For Each row As DataRow In ds.Tables(0).Rows
                Dim rep As New PedidoXRepresentante(Me._Pedido)

                rep.CodigoRepresentante = row("Representante_Id")
                rep.CodigoEnderecoRepresentante = row("EndRepresentante_Id")
                rep.Percentual = row("Percentual")
                rep.ValorComissao = row("ValorComissao")
                rep.Principal = row("Principal").Equals("S")
                rep.PercentualFixo = row("PercentualFixo")

                Me.Add(rep)
            Next

        Catch ex As Exception

        Finally
            objBanco = Nothing
        End Try
    End Sub
#End Region

#Region "Fieds"
    Private _Pedido As Pedido
#End Region

#Region "Property"
    Public ReadOnly Property Pedido As Pedido
        Get
            Return _Pedido
        End Get
    End Property
#End Region

#Region "Methods"
    Public Sub RecalcularComissoesFixas()
        For Each rep In Me
            If rep.PercentualFixo Then
                rep.IUD = "U"
                rep.ValorComissao = Math.Round(Me.Pedido.Itens.TotalOficial * rep.Percentual / 100, 2)
            End If
        Next
    End Sub

    Public Sub SalvarSql(ByVal Sqls As ArrayList)
        For Each rep As PedidoXRepresentante In Me
            If Pedido.IUD = "U" AndAlso rep.IUD = "I" Then
                rep.IUD = "D"
                rep.SalvarSql(Sqls)
                rep.IUD = "I"
                rep.SalvarSql(Sqls)
            ElseIf Pedido.IUD = "I" Or Pedido.IUD = "D" Or Pedido.IUD = "C" Then
                rep.IUD = Pedido.IUD
                rep.SalvarSql(Sqls)
            ElseIf Pedido.IUD = "U" Then
                If rep.IUD = "D" Then
                    rep.IUD = "D"
                    rep.SalvarSql(Sqls)
                End If
                'rep.IUD = "I"
                'rep.SalvarSql(Sqls)
            End If
        Next
    End Sub
#End Region

End Class

'*****************************************************************************************************************
'**************************************   Representantes do Pedido    ********************************************
'*****************************************************************************************************************
Public Class PedidoXRepresentante

#Region "Construtor"
    Public Sub New(ByVal pPedido As Pedido)
        _Pedido = pPedido
    End Sub
#End Region

#Region "Fields"
    Private _IUD As String = ""
    Private _Pedido As Pedido
    Private _CodigoRepresentante As String = ""
    Private _CodigoEnderecoRepresentante As Integer
    Private _Representante As Cliente
    Private _Percentual As Decimal
    Private _ValorComissao As Decimal
    Private _Principal As Boolean
    Private _PercentualFixo As Boolean
    Private _Comissoes As ListPedidoXRepresentantesxTabelaDeComissao
#End Region

#Region "Property"
    Public Property IUD As String
        Get
            Return _IUD
        End Get
        Set(value As String)
            _IUD = value
        End Set
    End Property

    Public ReadOnly Property Pedido() As Pedido
        Get
            Return _Pedido
        End Get
    End Property

    Public Property CodigoRepresentante() As String
        Get
            Return _CodigoRepresentante
        End Get
        Set(ByVal value As String)
            _CodigoRepresentante = value
        End Set
    End Property

    Public Property CodigoEnderecoRepresentante() As Integer
        Get
            Return _CodigoEnderecoRepresentante
        End Get
        Set(ByVal value As Integer)
            _CodigoEnderecoRepresentante = value
        End Set
    End Property

    Public Property Representante() As Cliente
        Get
            If _Representante Is Nothing And Not CodigoRepresentante Is Nothing Then _Representante = New Cliente(Me.CodigoRepresentante, Me.CodigoEnderecoRepresentante)
            Return _Representante
        End Get
        Set(ByVal value As Cliente)
            _Representante = value
        End Set
    End Property

    Public Property Percentual() As Decimal
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Decimal)
            _Percentual = value
        End Set
    End Property

    Public Property ValorComissao() As Decimal
        Get
            Return _ValorComissao
        End Get
        Set(ByVal value As Decimal)
            _ValorComissao = value
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

    Public Property PercentualFixo As Boolean
        Get
            Return _PercentualFixo
        End Get
        Set(value As Boolean)
            _PercentualFixo = value
        End Set
    End Property

    Public Property Comissoes() As ListPedidoXRepresentantesxTabelaDeComissao
        Get
            If _Comissoes Is Nothing Then _Comissoes = New ListPedidoXRepresentantesxTabelaDeComissao(Me, False)
            Return _Comissoes
        End Get
        Set(ByVal value As ListPedidoXRepresentantesxTabelaDeComissao)
            _Comissoes = value
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

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim sql As String
        Select Case Me.IUD
            Case "I"
                sql = "INSERT INTO Comissoes " & vbCrLf &
                      "       (Empresa_Id, EndEmpresa_Id, Pedido_Id," & vbCrLf &
                      "        Representante_Id, EndRepresentante_Id," & vbCrLf &
                      "        ValorComissao, Principal," & vbCrLf &
                      "        Percentual, PercentualFixo) " & vbCrLf &
                      "VALUES('" & Me.Pedido.CodigoEmpresa & "', " & Me.Pedido.EnderecoEmpresa & ", " & Me.Pedido.Codigo & "," & vbCrLf &
                      "'" & Me.CodigoRepresentante & "', " & Me.CodigoEnderecoRepresentante & "," & vbCrLf &
                      Str(Me.ValorComissao) & "," & "'" & IIf(Me.Principal, "S", "N") & "'," & vbCrLf &
                      Str(Me.Percentual) & ", " & IIf(Me.PercentualFixo, 1, 0) & ")"
                Sqls.Add(sql)
                'Me.Comissoes.SalvarSql(Sqls)

                'Adicionando Represenante a Empresa caso não exista - Furlan 27/11/2024
                'Utilizado para integração com a OnSoft
                If Left(Me.Pedido.CodigoEmpresa, 8) = "40938762" Then
                    'Empresa X Representante
                    sql = "DECLARE @RESULTADO INT" & vbCrLf &
                          "set @RESULTADO = (select count(*) from ClienteXRepresentante " & vbCrLf &
                          "					 where Cliente_Id          = '" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                          "					   and EndCliente_Id       = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                          "					   and Representante_Id    = '" & Me.CodigoRepresentante & "'" & vbCrLf &
                          "					   and EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & ");" & vbCrLf &
                          "IF (@RESULTADO = 0) -- INSERE EMPRESAxREPRESENTANTE" & vbCrLf &
                          "	   BEGIN" & vbCrLf &
                          "		   INSERT INTO ClienteXRepresentante (Cliente_Id, EndCliente_Id, Representante_Id, EndRepresentante_Id, Principal)" & vbCrLf &
                          "		   VALUES ('" & Me.Pedido.CodigoEmpresa & "'," & Me.Pedido.EnderecoEmpresa & ",'" & Me.CodigoRepresentante & "'," & Me.CodigoEnderecoRepresentante & ",0);" & vbCrLf &
                          "	   END" & vbCrLf &
                          "" & vbCrLf &
                          "" & vbCrLf
                    'Cliente X Representante
                    sql &= "set @RESULTADO = (select count(*) from ClienteXRepresentante " & vbCrLf &
                          "					 where Cliente_Id          = '" & Me.Pedido.CodigoCliente & "'" & vbCrLf &
                          "					   and EndCliente_Id       = " & Me.Pedido.EnderecoCliente & vbCrLf &
                          "					   and Representante_Id    = '" & Me.CodigoRepresentante & "'" & vbCrLf &
                          "					   and EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & ");" & vbCrLf &
                          "IF (@RESULTADO = 0) -- INSERE CLIENTExREPRESENTANTE" & vbCrLf &
                          "	   BEGIN" & vbCrLf &
                          "		   INSERT INTO ClienteXRepresentante (Cliente_Id, EndCliente_Id, Representante_Id, EndRepresentante_Id, Principal)" & vbCrLf &
                          "		   VALUES ('" & Me.Pedido.CodigoCliente & "'," & Me.Pedido.EnderecoCliente & ",'" & Me.CodigoRepresentante & "'," & Me.CodigoEnderecoRepresentante & ",0);" & vbCrLf &
                          "	   END"

                    Sqls.Add(sql)
                End If

            Case "U"
                sql = "Update Comissoes Set " & vbCrLf &
                      "    ValorComissao  = " & Str(Me.ValorComissao) & vbCrLf &
                      "   ,Principal      ='" & IIf(Me.Principal, "S", "N") & "'" & vbCrLf &
                      "   ,Percentual     = " & Str(Me.Percentual) & vbCrLf &
                      "   ,PercentualFixo = " & IIf(Me.PercentualFixo, 1, 0) & vbCrLf &
                      " WHERE Empresa_Id          ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                      "   AND EndEmpresa_Id       = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                      "   AND Pedido_Id           = " & Me.Pedido.Codigo & vbCrLf &
                      "   AND Representante_Id    ='" & Me.CodigoRepresentante & "'" & vbCrLf &
                      "   AND EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & vbCrLf
                Sqls.Add(sql)
                'Me.Comissoes.SalvarSql(Sqls)

                'Adicionando Represenante a Empresa caso não exista - Furlan 27/11/2024
                'Utilizado para integração com a OnSoft
                If Left(Me.Pedido.CodigoEmpresa, 8) = "40938762" Then
                    'Empresa X Representante
                    sql = "DECLARE @RESULTADO INT" & vbCrLf &
                          "set @RESULTADO = (select count(*) from ClienteXRepresentante " & vbCrLf &
                          "					 where Cliente_Id          = '" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                          "					   and EndCliente_Id       = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                          "					   and Representante_Id    = '" & Me.CodigoRepresentante & "'" & vbCrLf &
                          "					   and EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & ");" & vbCrLf &
                          "IF (@RESULTADO = 0) -- INSERE EMPRESAxREPRESENTANTE" & vbCrLf &
                          "	   BEGIN" & vbCrLf &
                          "		   INSERT INTO ClienteXRepresentante (Cliente_Id, EndCliente_Id, Representante_Id, EndRepresentante_Id, Principal)" & vbCrLf &
                          "		   VALUES ('" & Me.Pedido.CodigoEmpresa & "'," & Me.Pedido.EnderecoEmpresa & ",'" & Me.CodigoRepresentante & "'," & Me.CodigoEnderecoRepresentante & ",0);" & vbCrLf &
                          "	   END" & vbCrLf &
                          "" & vbCrLf &
                          "" & vbCrLf
                    'Cliente X Representante
                    sql &= "set @RESULTADO = (select count(*) from ClienteXRepresentante " & vbCrLf &
                          "					 where Cliente_Id          = '" & Me.Pedido.CodigoCliente & "'" & vbCrLf &
                          "					   and EndCliente_Id       = " & Me.Pedido.EnderecoCliente & vbCrLf &
                          "					   and Representante_Id    = '" & Me.CodigoRepresentante & "'" & vbCrLf &
                          "					   and EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & ");" & vbCrLf &
                          "IF (@RESULTADO = 0) -- INSERE CLIENTExREPRESENTANTE" & vbCrLf &
                          "	   BEGIN" & vbCrLf &
                          "		   INSERT INTO ClienteXRepresentante (Cliente_Id, EndCliente_Id, Representante_Id, EndRepresentante_Id, Principal)" & vbCrLf &
                          "		   VALUES ('" & Me.Pedido.CodigoCliente & "'," & Me.Pedido.EnderecoCliente & ",'" & Me.CodigoRepresentante & "'," & Me.CodigoEnderecoRepresentante & ",0);" & vbCrLf &
                          "	   END"

                    Sqls.Add(sql)
                End If

            Case "D"
                Me.Comissoes.SalvarSql(Sqls)
                sql = "DELETE Comissoes " & vbCrLf &
                      " WHERE Empresa_Id          ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                      "   AND EndEmpresa_Id       = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                      "   AND Pedido_Id           = " & Me.Pedido.Codigo & vbCrLf &
                      "   AND Representante_Id    ='" & Me.CodigoRepresentante & "'" & vbCrLf &
                      "   AND EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & vbCrLf
                Sqls.Add(sql)
            Case "C"
                'Me.Comissoes.SalvarSql(Sqls)
                sql = "DELETE Comissoes " & vbCrLf &
                      " WHERE Empresa_Id          ='" & Me.Pedido.CodigoEmpresa & "'" & vbCrLf &
                      "   AND EndEmpresa_Id       = " & Me.Pedido.EnderecoEmpresa & vbCrLf &
                      "   AND Pedido_Id           = " & Me.Pedido.Codigo & vbCrLf &
                      "   AND Representante_Id    ='" & Me.CodigoRepresentante & "'" & vbCrLf &
                      "   AND EndRepresentante_Id = " & Me.CodigoEnderecoRepresentante & vbCrLf
                Sqls.Add(sql)
            Case Else
                'Me.Comissoes.SalvarSql(Sqls)
        End Select
    End Sub
#End Region

End Class

