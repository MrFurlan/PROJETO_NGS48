Imports NGS.Lib.Negocio
Imports NGS.Lib.Uteis

Public Class MDTwoKeys(Of k1, k2, v)
    Private Property List1() As List(Of k1)
        Get
            Return m_List1
        End Get
        Set(value As List(Of k1))
            m_List1 = value
        End Set
    End Property
    Private m_List1 As List(Of k1)
    Private Property List2() As List(Of k2)
        Get
            Return m_List2
        End Get
        Set(value As List(Of k2))
            m_List2 = value
        End Set
    End Property
    Private m_List2 As List(Of k2)
    Private Property List3() As Dictionary(Of Int32, v)
        Get
            Return m_List3
        End Get
        Set(value As Dictionary(Of Int32, v))
            m_List3 = value
        End Set
    End Property
    Private m_List3 As Dictionary(Of Int32, v)

#Region "constructor"
    Public Sub New()
        List1 = New List(Of k1)()
        List2 = New List(Of k2)()
        List3 = New Dictionary(Of Int32, v)()
    End Sub
#End Region

#Region "public properties"
    Default Public Property Item(key1 As k1, key2 As k2) As v
        Get
            Dim index = IndexOf(key1, key2)
            Return If(index = -1, Nothing, List3(index))
        End Get
        Set(value As v)
            Remove(key1, key2)
            Add(key1, key2, value)
        End Set
    End Property

    Public ReadOnly Property Keys() As List(Of KeysPair)
        Get
            Dim lst = New List(Of KeysPair)()
            For x As Integer = 0 To List1.Count - 1
                lst.Add(New MDTwoKeys(Of k1, k2, v).KeysPair() With { _
                 .Key1 = List1(x), _
                 .Key2 = List2(x) _
                })
            Next
            Return lst
        End Get
    End Property

    Public ReadOnly Property Values() As Dictionary(Of Int32, v).ValueCollection
        Get
            Return List3.Values
        End Get
    End Property
#End Region

#Region "public methods"
    Public Function ContainsKeys(key1 As k1, key2 As k2) As [Boolean]
        Return IndexOf(key1, key2) <> -1
    End Function

    Public Function ContainsValue(value As v) As [Boolean]
        Return List3.ContainsValue(value)
    End Function

    Public Sub Remove(key1 As k1, key2 As k2)
        Dim index = IndexOf(key1, key2)
        If index <> -1 Then
            RemoveAt(index)
        End If
    End Sub

    Public Sub RemoveAt(index As Int32)
        List1.RemoveAt(index)
        List2.RemoveAt(index)
        List3.Remove(index)
    End Sub

    Public Function IndexOf(key1 As k1, key2 As k2) As Int32
        If key1 Is Nothing OrElse key2 Is Nothing Then
            Return -1
        End If
        For x As Integer = 0 To List1.Count - 1
            If key1.Equals(List1(x)) AndAlso key2.Equals(List2(x)) Then
                Return x
            End If
        Next
        Return -1
    End Function

    Public Sub Add(key1 As k1, key2 As k2, value As v)
        If ContainsKeys(key1, key2) Then
            Throw New ArgumentException("Duplicate key cannot be added.")
        End If
        List1.Add(key1)
        List2.Add(key2)
        List3.Add(List3.Count, value)
    End Sub
#End Region

#Region "internal class keyspair"
    Public Structure KeysPair
        Public Property Key1() As k1
            Get
                Return m_Key1
            End Get
            Set(value As k1)
                m_Key1 = value
            End Set
        End Property
        Private m_Key1 As k1
        Public Property Key2() As k2
            Get
                Return m_Key2
            End Get
            Set(value As k2)
                m_Key2 = value
            End Set
        End Property
        Private m_Key2 As k2
    End Structure
#End Region
End Class