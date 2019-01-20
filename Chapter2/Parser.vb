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
                            "Expected {0}", _
                            errorDescription _
                )
            Case Else
                message = "Unknown error."
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
        ElseIf "+-".IndexOf(c) <> -1 And _
            TokenLength = 0 Then
            ' If the symbol being cheked is + or -
            ' AND we are at the START of the current
            ' token
            result = True
        Else
            result = False
        End If

        Return result
    End Function

    Private Function IsWhiteSpace(ByVal c As Char) As Boolean
        Return Char.IsWhiteSpace(c)
    End Function

    Private Function IsMulOrDivOperator(ByVal c As Char) As Boolean
        Return "*/".IndexOf(c) > -1
    End Function

    Private Function IsAddOrSubOperator(ByVal c As Char) As Boolean
        Return "+-".IndexOf(c) > -1
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

    Private Sub SkipWhiteSpace()
        Do While IsWhiteSpace(LookAhead)
            If EndOfLine() Then
                Exit Do
            Else
                m_CharPos += 1
            End If
        Loop
    End Sub

    Private Sub ScanMulOrDivOperator()
        m_CurrentTokenBldr = New StringBuilder

        If IsMulOrDivOperator(LookAhead) Then
            m_CurrentTokenBldr.Append(LookAhead)
            m_CharPos += 1
        End If
    End Sub

    Private Sub ScanAddOrSubOperator()
        m_CurrentTokenBldr = New StringBuilder

        If IsAddOrSubOperator(LookAhead) Then
            m_CurrentTokenBldr.Append(LookAhead)
            m_CharPos += 1
        End If
    End Sub

    Private Sub SkipCharacter()
        m_CharPos += 1
    End Sub
#End Region

#Region "Parser"
    Private Function ParseLine() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        result = ParseNumericExpression()

        If result.code = 0 Then
            If Not EndOfLine() Then
                result = CreateError(1, "end of statement")
            Else
                m_Gen.EmitWriteLine()
            End If
        End If

        Return result
    End Function

    Private Function ParseNumber() As ParseStatus
        Dim result As ParseStatus

        ' Ask the scanner to read a number
        ScanNumber()

        If TokenLength = 0 Then
            result = CreateError(1, "a number")
        ElseIf TokenLength = 1 AndAlso _
                Not Char.IsDigit(CurrentToken.Chars(0)) _
                Then 
            result = CreateError(1, "a number")
        Else
            ' Get the current token, and
            ' Emit it
            m_Gen.EmitNumber(CInt(CurrentToken))
            result = CreateError(0, "Ok")
        End If

        Return result
    End Function

    Private Function ParseFactor() As ParseStatus
        Dim result As ParseStatus

        If LookAhead.Equals("("c) Then
            SkipCharacter()

            result = ParseNumericExpression()

            If result.Code = 0 Then
                If Not LookAhead.Equals(")"c) Then
                    result = CreateError(1, ")")
                Else
                    SkipCharacter()
                End If
            End If
        Else
            result = ParseNumber()
        End If

        SkipWhiteSpace()

        Return result
    End Function

    Private Function ParseMulOrDivOperator() As ParseStatus
        Dim result As ParseStatus
        Dim currentoperator As String = CurrentToken

        SkipWhiteSpace()

        result = ParseFactor()

        If result.Code = 0 Then
            If currentoperator = "*" Then
                m_Gen.EmitMultiply()
            Else
                m_Gen.EmitDivide()
            End If
        End If

        Return result
    End Function

    Private Function ParseTerm() As ParseStatus
        Dim result As ParseStatus

        result = ParseFactor()

        Do While result.Code = 0 _
            AndAlso _
            IsMulOrDivOperator(LookAhead)

            ScanMulOrDivOperator()

            If TokenLength = 0 Then
                result = CreateError(1, "* or /")
            Else
                result = ParseMulOrDivOperator()
                SkipWhiteSpace()
            End If
        Loop

        Return result
    End Function

    Private Function ParseAddOrSubOperator() As ParseStatus
        Dim result As ParseStatus
        Dim currentoperator As String = CurrentToken

        SkipWhiteSpace()

        result = ParseTerm()

        If result.Code = 0 Then
            If currentoperator = "+" Then
                m_Gen.EmitAdd()
            Else
                m_Gen.EmitSubtract()
            End If
        End If

        Return result
    End Function

    Private Function ParseNumericExpression() As ParseStatus
        Dim result As ParseStatus

        result = ParseTerm()

        Do While result.Code = 0 _
            AndAlso _
            IsAddOrSubOperator(LookAhead)

            ScanAddOrSubOperator()

            If TokenLength = 0 Then
                result = CreateError(1, "+ or -")
            Else
                result = ParseAddOrSubOperator()
                SkipWhiteSpace()
            End If
        Loop

        Return result
    End Function
#End Region

    Public Function Parse() As ParseStatus
        Dim result As ParseStatus
        result = CreateError(0, "Ok.")
        
        Do While ScanLine()
            result = ParseLine()
            If result.Code <> 0 Then
                Exit Do
            End If
        Loop
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
