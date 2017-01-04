Imports System
Imports System.IO
Imports System.Text

Public Class Parser

#Region "Fields"
    Private m_InputStream As TextReader
    Private m_Gen As CodeGen

    ' The current line being translated
    Private m_ThisLine As String
    Private m_LineLength As Integer = 0

    ' The position of the line being translated
    Private m_linePos As Integer = 0
    ' The character position currently being looked at
    Private m_CharPos As Integer = 0

    ' The current program element
    Private m_CurrentTokenBldr As StringBuilder
#End Region

#Region "Helper Functions"
    Private Function CreateError( _
        ByVal errorcode As Integer, _
        ByVal errorDescription As String _
        ) As ParseStatus

        Dim result As ParseStatus
        Dim message As String

        Select Case errorcode
            Case 0
                message = "Ok."
            Case 1
                message = String.Format( _
                            "Expected '{0}'", _
                            errorDescription _
                )
        End Select

        result = New ParseStatus(errorcode, _
                    message, _
                    m_CharPos + 1, _
                    m_linePos)


        Return result
    End Function
#End Region

#Region "Recognizers"
    Private Function IsNumeric(ByVal c As Char) As Boolean
        Dim result As Boolean

        If Char.IsDigit(c) Then
            ' Use the CLR's built-in digit recognition
            result = True
        Else
            result = False
        End If

        Return result
    End Function
#End Region

#Region "Scanner"
    Private ReadOnly Property LookAhead() As Char
        Get
            Dim result As Char
            If m_CharPos < m_LineLength Then
                result = m_ThisLine.Chars(m_CharPos)
            Else
                result = " "c
            End If
            Return result
        End Get
    End Property

    Private ReadOnly Property CurrentToken() As String
        Get
            Return m_CurrentTokenBldr.ToString()
        End Get
    End Property

    Private ReadOnly Property TokenLength() As Integer
        Get
            Return m_CurrentTokenBldr.Length()
        End Get
    End Property

    Private ReadOnly Property EndOfLine() As Boolean
        Get
            Return m_CharPos >= m_LineLength
        End Get
    End Property

    Private Function ScanLine() As Boolean
        Dim result As Boolean
        Dim line As String

        line = m_InputStream.ReadLine()
        If line Is Nothing Then
            result = False
        Else
            ' set up line, line length, 
            ' increment line counter, 
            ' and set character position back to 0
            m_ThisLine = line
            m_LineLength = m_ThisLine.Length
            m_linePos += 1
            m_CharPos = 0
            result = True
        End If

        Return result
    End Function

    Private Sub ScanNumber()
        m_CurrentTokenBldr = New StringBuilder

        Do While m_CharPos < m_LineLength
            If Not IsNumeric(LookAhead) Then
                Exit Do
            End If
            m_CurrentTokenBldr.Append(LookAhead)
            m_CharPos += 1
        Loop
    End Sub
#End Region

#Region "Parser"
    Private Function ParseLine() As ParseStatus
        Dim result As ParseStatus

        result = ParseNumber()

        If result.code = 0 Then
            m_Gen.EmitWriteLine()
        End If

        Return result
    End Function

    Private Function ParseNumber() As ParseStatus
        Dim result As ParseStatus

        ' Ask the scanner to read a number
        ScanNumber()

        If TokenLength = 0 Then
            result = CreateError(1, " a number")
        Else
            ' Get the current token, and
            ' Emit it
            m_Gen.EmitNumber(CInt(CurrentToken))
            result = CreateError(0, "Ok")
        End If

        Return result
    End Function
#End Region

    Public Function Parse() As ParseStatus
        Dim result As ParseStatus
        If ScanLine() Then
            result = ParseLine()
        End If
        Return result
    End Function

    Public Sub New( _
        ByVal newStream As TextReader, _
        ByVal newGen As CodeGen _
        )

        m_InputStream = newStream
        m_Gen = newGen
    End Sub
End Class
