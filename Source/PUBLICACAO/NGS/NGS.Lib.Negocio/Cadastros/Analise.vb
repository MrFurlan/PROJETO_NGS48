Imports Microsoft.VisualBasic
Imports System.Data
Imports NGS.Lib.Uteis

'****************************************************************************************************************************************
'********************************************  LISTA DE PESAGEM X ANALISES   ************************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class ListAnalise
    Inherits List(Of Analise)

    Public Sub New()
        Dim sql As String
        sql = "SELECT Analise_Id, Descricao, IndiceMinimo, IndiceMaximo, isnull(opcao,'') as Opcao " & _
              "  FROM Analises" & _
              " Order By Analise_Id"

        Dim Banco As New AcessaBanco
        Dim ds As DataSet = Banco.ConsultaDataSet(sql, "Analises")

        For Each row As DataRow In ds.Tables(0).Rows
            Dim obj As New Analise()
            obj.Codigo = row("Analise_Id")
            obj.Descricao = row("Descricao")
            obj.IndiceMinimo = row("IndiceMinimo")
            obj.IndiceMaximo = row("IndiceMaximo")
            obj.Opcao = row("Opcao")
            Me.Add(obj)
        Next
    End Sub
End Class

'****************************************************************************************************************************************
'*******************************************  CLASSE BASE DE PESAGEM X ANALISES   *******************************************************
'****************************************************************************************************************************************
<Serializable()> _
Public Class Analise
    Implements IBaseEntity

#Region "Fields"
    Private _IUD As String
    Private _Codigo As Integer
    Private _Descricao As String
    Private _IndiceMinimo As Double
    Private _IndiceMaximo As Double
    Private _Opcao As String
#End Region

#Region "Construtor"
    Public Sub New()

    End Sub

    Public Sub New(ByVal Codigo As Integer)
        Dim objBanco As New AcessaBanco()

        Try
            Dim strSQL As String
            strSQL = "SELECT Analise_Id, Descricao, IndiceMinimo, IndiceMaximo, isnull(Opcao,'') as Opcao" & _
                     "  FROM Analises " & _
                     " WHERE Analise_Id = " & Codigo.ToString()

            Dim dsAnalises As DataSet = objBanco.ConsultaDataSet(strSQL, "Analises")

            For Each drAnalise As DataRow In dsAnalises.Tables(0).Rows
                With Me
                    .Codigo = drAnalise("Analise_Id")
                    .Descricao = drAnalise("Descricao")
                    .IndiceMinimo = drAnalise("IndiceMinimo")
                    .IndiceMaximo = drAnalise("IndiceMaximo")
                    .Opcao = drAnalise("Opcao")
                End With
            Next
        Catch ex As Exception
        Finally
            objBanco = Nothing
        End Try
    End Sub
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

    Public Property Codigo() As Integer
        Get
            Return _Codigo
        End Get
        Set(ByVal value As Integer)
            _Codigo = value
        End Set
    End Property

    Public Property Descricao() As String
        Get
            Return _Descricao
        End Get
        Set(ByVal value As String)
            _Descricao = value
        End Set
    End Property

    Public Property IndiceMinimo() As Double
        Get
            Return _IndiceMinimo
        End Get
        Set(ByVal value As Double)
            _IndiceMinimo = value
        End Set
    End Property

    Public Property IndiceMaximo() As Double
        Get
            Return _IndiceMaximo
        End Get
        Set(ByVal value As Double)
            _IndiceMaximo = value
        End Set
    End Property

    Public Property Opcao As String
        Get
            Return _Opcao
        End Get
        Set(value As String)
            _Opcao = value
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
        Dim Sql As String = ""
        Select Case Me.IUD
            Case "I"
                Sql = " INSERT INTO Analises (Analise_Id, Descricao, IndiceMinimo, IndiceMaximo) " & vbCrLf & _
                      " VALUES (" & _Codigo & "," & _Descricao & ",'" & _IndiceMinimo & "','" & _IndiceMaximo & "')"
                Sqls.Add(Sql)
            Case "U"
                Sql = " UPDATE Analises SET" & vbCrLf & _
                      "    Analise_Id  = " & _Codigo & "," & vbCrLf & _
                      "    Descricao     = '" & _Descricao & "'" & vbCrLf & _
                      "    IndiceMinimo     = '" & _IndiceMinimo & "'" & vbCrLf & _
                      "    IndiceMaximo     = '" & _IndiceMaximo & "'" & vbCrLf & _
                      "  WHERE Analise_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
            Case "D"
                Sql = " DELETE Analises" & vbCrLf & _
                      "  WHERE Analise_Id = " & _Codigo & vbCrLf
                Sqls.Add(Sql)
        End Select
    End Sub
#End Region

End Class


