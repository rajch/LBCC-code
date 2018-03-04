Option Strict On
Option Explicit On

Public Class ParseStatus
    Private m_Code As Integer
    Private m_Description As String
    Private m_Row As Integer
    Private m_Column As Integer

    Public Sub New(ByVal newCode As Integer, _
            ByVal newDescription As String, _
            ByVal newColumn As Integer, _
            ByVal newRow As Integer _
            )
        m_Code = newCode
        m_Description = newDescription
        m_Column = newColumn
        m_Row = newRow
    End Sub

    Public ReadOnly Property Code() As Integer
        Get
            Return m_Code
        End Get
    End Property

    Public ReadOnly Property Description() As String
        Get
            Return m_Description
        End Get
    End Property

    Public ReadOnly Property Column() As Integer
        Get
            Return m_Column
        End Get
    End Property

    Public ReadOnly Property Row() As Integer
        Get
            Return m_Row
        End Get
    End Property
End Class