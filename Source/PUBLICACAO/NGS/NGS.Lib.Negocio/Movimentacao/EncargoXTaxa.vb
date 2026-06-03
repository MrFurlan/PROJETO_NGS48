Imports System.Data
Imports Microsoft.VisualBasic
Imports NGS.Lib.Uteis

<Serializable()> _
Public Class EncargoXTaxa

    Private _Estado As String
    Private _Encargo As String
    Private _Produto As String
    Private _Data As DateTime
    Private _Percentual As Double
    Private _SimplesNacional As Double

    Public Erro As Exception

    Public Property Estado() As String
        Get
            Return _Estado
        End Get
        Set(ByVal value As String)
            _Estado = value
        End Set
    End Property

    Public Property Encargo() As String
        Get
            Return _Encargo
        End Get
        Set(ByVal value As String)
            _Encargo = value
        End Set
    End Property

    Public Property Produto() As String
        Get
            Return _Produto
        End Get
        Set(ByVal value As String)
            _Produto = value
        End Set
    End Property

    Public Property Data() As DateTime
        Get
            Return _Data
        End Get
        Set(ByVal value As DateTime)
            _Data = value
        End Set
    End Property

    Public Property Percentual() As Double
        Get
            Return _Percentual
        End Get
        Set(ByVal value As Double)
            _Percentual = value
        End Set
    End Property

    Public Property SimplesNacional() As Double
        Get
            Return _SimplesNacional
        End Get
        Set(ByVal value As Double)
            _SimplesNacional = value
        End Set
    End Property

    Public Function SelecionarVigente(ByVal Estado As String, ByVal Encargo As String, ByVal database As Date, Optional ByVal Produto As String = "")
        Dim objBanco As New AcessaBanco

        Try
            Dim strSQL As String = "SELECT Data_Id, Percentual, SimplesNacional " &
                                   "FROM EncargosXTaxas ET " &
                                   "WHERE Data_Id = (SELECT MAX(Data_Id) " &
                                                    "FROM EncargosXTaxas ET1 " &
                                                    "WHERE ET1.Estado_Id = ET.Estado_Id " &
                                                    "AND ET1.Encargo_Id = ET.Encargo_Id "

            If Produto.Length > 0 Then strSQL &= "AND ET1.Produto_Id = ET.Produto_Id "

            strSQL &= "AND ET1.Data_Id <= '" & database.ToSqlDate & "')" &
                                   "AND Estado_Id = '" & Estado & "' " &
                                   "AND Encargo_Id = '" & Encargo & "'"

            If Produto.Length > 0 Then strSQL &= "AND Produto_Id = '" & Produto & "'"

            Dim dsVigente As DataSet = objBanco.ConsultaDataSet(strSQL, "EncargosXTaxas")

            If dsVigente.Tables(0).Rows.Count > 0 Then
                Dim drVigente As DataRow = dsVigente.Tables(0).Rows(0)

                Me.Estado = Estado
                Me.Encargo = Encargo
                Me.Produto = Produto
                Me.Data = Convert.ToDateTime(drVigente("Data_Id"))
                Me.Percentual = Convert.ToDouble(drVigente("Percentual"))
                Me.SimplesNacional = Convert.ToDouble(drVigente("SimplesNacional"))
            End If

            Return True
        Catch ex As Exception
            Me.Erro = ex
            Return False
        Finally
            objBanco = Nothing
        End Try
    End Function

End Class