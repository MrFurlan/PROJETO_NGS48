Imports System.Collections.Generic
Imports System.Data
Imports NGS.Lib.Uteis
Imports System.Configuration


'***********************************************************************************************************************************************
'************************************* Lista dos Titulos da Fatura de Frete ********************************************************************
'***********************************************************************************************************************************************
<Serializable()> _
Public Class ListFaturaDeFretexTitulo
    Inherits List(Of FaturaDeFretexTitulo)
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(pFaturaDeFrete As FaturaDeFrete)
        _FaturaDeFrete = pFaturaDeFrete
        Dim sql As String
        sql = "SELECT Titulo_Id" & vbCrLf & _
              "  FROM FaturaDeFreteXTitulo" & vbCrLf & _
              " Where Empresa_Id        ='" & pFaturaDeFrete.CodigoEmpresa & "'" & vbCrLf & _
              "   and EndEmpresa_Id     = " & pFaturaDeFrete.EnderecoEmpresa & vbCrLf & _
              "   and Conveniado_Id     ='" & pFaturaDeFrete.CodigoConveniado & "'" & vbCrLf & _
              "   and EndConveniado_Id  = " & pFaturaDeFrete.EnderecoConveniado & vbCrLf & _
              "   and Fatura_Id         = " & pFaturaDeFrete.CodigoFatura
        Dim ds As DataSet
        Dim Banco As New AcessaBanco
        ds = Banco.ConsultaDataSet(sql, "Titulos")

        For Each row In ds.Tables(0).Rows
            Dim fxt As New FaturaDeFretexTitulo(pFaturaDeFrete)
            fxt.CodigoTitulo = row("Titulo_id")
            Me.Add(fxt)
        Next
    End Sub
#End Region

#Region "Fields"
    Private _FaturaDeFrete As FaturaDeFrete
#End Region

#Region "Property"
    Public ReadOnly Property FaturaDeFrete As FaturaDeFrete
        Get
            Return _FaturaDeFrete
        End Get
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim sqls As New ArrayList

        sqls.Clear()
        Me.SalvarSql(sqls)

        Return sqls.Count = 0 OrElse Banco.GravaBanco(sqls)
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        For Each row In Me
            If FaturaDeFrete.IUD = "D" OrElse FaturaDeFrete.IUD = "I" Then row.IUD = FaturaDeFrete.IUD
            row.SalvarSql(Sqls)
        Next
    End Sub
#End Region

End Class


'***********************************************************************************************************************************************
'************************************* Classe Base Titulo da Fatura de Frete *******************************************************************
'***********************************************************************************************************************************************
Public Class FaturaDeFretexTitulo
    Implements IBaseEntity

#Region "Construtor"
    Public Sub New(pFaturaDeFrete As FaturaDeFrete)
        _FaturaDeFrete = pFaturaDeFrete
    End Sub
#End Region

#Region "fields"
    Private _IUD As String = ""
    Private _FaturaDeFrete As FaturaDeFrete
    Private _CodigoTitulo As Integer
    Private _Titulo As Titulo
    Private _TituloNovo As Novo.TituloNovo
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

    Public ReadOnly Property FaturaDeFrete As FaturaDeFrete
        Get
            Return _FaturaDeFrete
        End Get
    End Property

    Public Property CodigoTitulo As Integer
        Get
            Return _CodigoTitulo
        End Get
        Set(value As Integer)
            _CodigoTitulo = value
            _Titulo = Nothing
        End Set
    End Property

    Public Property Titulo As Titulo
        Get
            If _Titulo Is Nothing And Me.CodigoTitulo > 0 Then _Titulo = New Titulo(Me.CodigoTitulo)
            Return _Titulo
        End Get
        Set(value As Titulo)
            _Titulo = value
        End Set
    End Property

    Public Property TituloNovo As Novo.TituloNovo
        Get
            If _TituloNovo Is Nothing And Me.CodigoTitulo > 0 Then _TituloNovo = New Novo.TituloNovo(Me.CodigoTitulo)
            Return _TituloNovo
        End Get
        Set(value As Novo.TituloNovo)
            _TituloNovo = value
        End Set
    End Property
#End Region

#Region "Methods"
    Public Function Salvar() As Boolean
        Dim Banco As New AcessaBanco
        Dim Sqls As New ArrayList
        Sqls.Clear()
        SalvarSql(Sqls)
        If Banco.GravaBanco(Sqls) Then
            Me.IUD = Nothing
            Return True
        Else
            Return False
        End If
    End Function

    Public Sub SalvarSql(ByRef Sqls As ArrayList)
        Dim strSQL As String

        Select Case Me.IUD
            Case "I"

                Dim CodTit As Integer
                If Me.CodigoTitulo <= 0 Then
                    If FaturaDeFrete.FinanceiroNovo Then
                        TituloNovo.IUD = "I"
                        TituloNovo.SalvarSql(Sqls)
                        CodTit = TituloNovo.Codigo
                    Else
                        Titulo.IUD = "I"
                        Titulo.SalvarSql(Sqls)
                        CodTit = Titulo.Codigo
                    End If
                Else
                    CodTit = Me.CodigoTitulo
                End If

                strSQL = "INSERT INTO FaturaDeFreteXTitulo " & vbCrLf & _
                         "       (Empresa_Id, EndEmpresa_Id, Conveniado_Id, EndConveniado_Id, Fatura_Id, Titulo_Id)" & vbCrLf & _
                         "VALUES ('" & FaturaDeFrete.CodigoEmpresa & "'," & FaturaDeFrete.EnderecoEmpresa & ",'" & FaturaDeFrete.CodigoConveniado & "'," & FaturaDeFrete.EnderecoConveniado & "," & FaturaDeFrete.CodigoFatura & "," & CodTit & "); "
                Sqls.Add(strSQL)
            Case "D"
                strSQL = "DELETE FaturaDeFreteXTitulo" & vbCrLf & _
                         " WHERE Empresa_Id       ='" & FaturaDeFrete.CodigoEmpresa & "'" & vbCrLf & _
                         "   AND EndEmpresa_Id    = " & FaturaDeFrete.EnderecoEmpresa & vbCrLf & _
                         "	 AND Conveniado_Id    ='" & FaturaDeFrete.CodigoConveniado & "'" & vbCrLf & _
                         "	 AND EndConveniado_Id = " & FaturaDeFrete.EnderecoConveniado & vbCrLf & _
                         "	 AND Fatura_Id        = " & FaturaDeFrete.CodigoFatura & vbCrLf & _
                         "   AND Titulo_Id        = " & Me.CodigoTitulo & ";" & vbCrLf
                Sqls.Add(strSQL)

                If FaturaDeFrete.FinanceiroNovo Then
                    TituloNovo.IUD = "D"
                    TituloNovo.SalvarSql(Sqls)
                Else
                    Titulo.IUD = "C"
                    Titulo.SalvarSql(Sqls)
                End If
            Case ""
                If FaturaDeFrete.FinanceiroNovo Then
                    If TituloNovo.IUD <> "" Then TituloNovo.SalvarSql(Sqls)
                Else
                    If Titulo.IUD <> "" Then Titulo.SalvarSql(Sqls)
                End If

        End Select
    End Sub



#End Region

End Class