Option Strict On
Option Explicit On

Imports System
Imports System.IO
Imports System.Text

Public Partial Class Parser

#Region "Fields"
    Private m_InputStream As TextReader
    Private m_Gen As CodeGen

    ' The current line being translated
    Private m_ThisLine As String
    Private m_LineLength As Integer = 0

    ' The position of the line being translated
    Private m_LinePos As Integer = 0
    ' The character position currently being looked at
    Private m_CharPos As Integer = 0

    ' The current program element
    Private m_CurrentTokenBldr As StringBuilder

    ' Whether the last string scanned was empty
    Private m_EmptyStringFlag As Boolean = False

    ' The type of the last processed expression
    Private m_LastTypeProcessed As Type

    ' Block stack
    Private m_BlockStack As New BlockStack

    ' Symbol table
    Private m_SymbolTable As New SymbolTable
#End Region

#Region "Helper Functions"
    Private Function MakeUniqueVariableName() As String
        Dim result As New StringBuilder("_clrcompiler_t_i4_")

        Do
            result.Append(CInt(Rnd() * 10))
        Loop Until Not m_SymbolTable.Exists(result.ToString)

        Return result.ToString
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
            ' If the symbol being checked is + or -
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

    Private Function IsConcatOperator(c As Char) As Boolean
        Return "&+".IndexOf(c) > -1
    End Function

    Private Function IsRelOperator(c As Char) As Boolean
        Return "=><!".IndexOf(c)>-1
    End Function

    Private Function IsNotOperatorSymbol(ByVal c As Char) As Boolean
        Return "nNoOtT!".IndexOf(c) > -1
    End Function

    Private Function IsNotOperator(ByVal c As Char) As Boolean
        Dim result As Boolean = False

        If c = "!"c Then
            result = True
        ElseIf Char.ToLowerInvariant(c) = "n"c Then
            If PeekAhead(3).ToLowerInvariant() = "not" Then
                result = True
            End If
        End If

        Return result
    End Function

    Private Function IsAndOperator(ByVal c As Char) As Boolean
        Return "aAnNdD&".IndexOf(c) > -1
    End Function

    Private Function IsOrOperator(ByVal c As Char) As Boolean
        Return "oOrR|".IndexOf(c) > -1
    End Function

    Private Function IsNameCharacter(ByVal c As Char) As Boolean
        Dim result As Boolean = False
        If Char.IsDigit(c) AndAlso _
                    TokenLength > 0 Then
            ' Digits allowed after start of name
            result = True
        ElseIf c.Equals("_"c) Then
            result = True
        ElseIf Char.IsLetter(c) Then
            result = True
        End If
        Return result
    End Function

    Private Function IsAssignmentCharacter(Byval c As Char) As Boolean
        Return ":=".IndexOf(c) > -1
    End Function

    Private Function IsNameStartCharacter(Byval c As Char) As Boolean
        Return c.Equals("_"c) OrElse Char.IsLetter(c)
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

    Private ReadOnly Property CurrentLine() As Integer
        Get
            Return m_linePos
        End Get
    End Property

    Private ReadOnly Property CurrentPosition() As Integer
        Get
            Return m_CharPos
        End Get
    End Property

    Private Function PeekAhead(ByVal count As Integer) As String
        Dim result As String
        Dim charcount As Integer = m_LineLength - m_CharPos

        If charcount <= 0 Then
            result = ""                    
        Else
            If charcount > count Then charcount = count
            result = m_ThisLine.Substring(m_CharPos,charcount)
        End if

        Return result
    End Function

    Private Sub SkipCharacter()
        m_CharPos += 1
    End Sub

    Private Sub SkipWhiteSpace()
        Do While IsWhiteSpace(LookAhead)
            If EndOfLine() Then
                Exit Do
            Else
                SkipCharacter()
            End If
        Loop
    End Sub

    Private Sub SkipRestOfLine()
        m_CharPos = m_LineLength
    End Sub

    Private Sub Backtrack()
        If TokenLength > 0 Then
            m_CharPos -= TokenLength
            ResetToken()
        End If
    End Sub

    Private Sub AppendToToken()
        m_CurrentTokenBldr.Append(LookAhead)
        m_CharPos += 1
    End Sub

    Private Sub ResetToken()
        m_CurrentTokenBldr = New StringBuilder()
    End Sub

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
            m_LinePos += 1
            m_CharPos = 0
            result = True
        End If

        Return result
    End Function

    Private Sub ScanNumber()
        ResetToken()

        Do While m_CharPos < m_LineLength
            If Not IsNumeric(LookAhead) Then
                Exit Do
            End If
            AppendToToken()
        Loop
    End Sub

    Private Sub ScanMulOrDivOperator()
        ResetToken()

        If IsMulOrDivOperator(LookAhead) Then
            AppendToToken()
        End If
    End Sub

    Private Sub ScanAddOrSubOperator()
        ResetToken()

        If IsAddOrSubOperator(LookAhead) Then
            AppendToToken()
        End If
    End Sub

    Private Sub ScanString()
        m_EmptyStringFlag = False
        ResetToken()

        If Not LookAhead.Equals(""""c) Then
            Exit Sub
        End If

        Do While LookAhead.Equals(""""c)
            SkipCharacter()
            Do While Not LookAhead.Equals(""""c)

                If EndOfLine Then
                    ResetToken()
                    Exit Sub
                End If

                AppendToToken()
            Loop

            SkipCharacter()

            If LookAhead.Equals(""""c) Then
                m_CurrentTokenBldr.Append(LookAhead)
            End If
        Loop

        If TokenLength = 0 Then
            m_EmptyStringFlag = True
        End If
    End Sub

    Private Sub ScanConcatOperator
        ResetToken()
        
        If IsConcatOperator(LookAhead) Then
            AppendToToken()
        End If	
    End Sub

    Private Sub ScanRelOperator
        ResetToken()
        
        Do While IsRelOperator(LookAhead)
            AppendToToken()
        Loop

        Select Case CurrentToken
            Case "=","==", "===", "<>", "!=", "!==", ">", "<", ">=", "=>","<=","=<"
                ' Valid relational operator
            Case Else
                ResetToken()
        End Select
    End Sub

    Private Sub ScanNotOperator()
        ResetToken()

        Do While IsNotOperatorSymbol(LookAhead)
            AppendToToken()
        Loop

        Select Case CurrentToken.ToLowerInvariant()
            Case "not", "!"
                ' Valid NOT operator
            Case Else
                ResetToken()
        End Select
    End Sub

    Private Sub ScanAndOperator()
        ResetToken()

        Do While IsAndOperator(LookAhead)
            AppendToToken()
        Loop

        Select Case CurrentToken.ToLowerInvariant()
            Case "and", "&"
                ' Valid AND operator
            Case Else
                ResetToken()
        End Select
    End Sub

    Private Sub ScanOrOperator()
        ResetToken()

        Do While IsOrOperator(LookAhead)
            AppendToToken()
        Loop

        Select Case CurrentToken.ToLowerInvariant()
            Case "or", "|"
                ' Valid OR operator
            Case Else
                ResetToken()
        End Select
    End Sub

    Private Sub ScanName()
        ResetToken()
        Do While IsNameCharacter(LookAhead)
            AppendToToken()
            If EndOfLine Then
                Exit Do
            End If
        Loop
    End Sub

    Private Sub ScanAssignmentOperator()
        ResetToken()
        Do While IsAssignmentCharacter(LookAhead)
            AppendToToken()
        Loop

        Select Case CurrentToken
            Case "=", ":="
                ' Valid Assignment operator
            Case Else
                ResetToken()
        End Select
    End Sub
#End Region

#Region "Parser"
    Private Function ParseLine() As ParseStatus
        Dim result As ParseStatus
        
        SkipWhiteSpace()
        
        m_LastTypeProcessed = Nothing
        
        If EndOfLine() Then
            ' An empty line is valid
            result = Ok()
        Else
            ScanName()

            If TokenLength = 0 Then
                result = CreateError(1, "a statement.")
            Else
                Dim name As String
                name = CurrentToken

                If IsValidCommand(name) Then
                    result = ParseCommand()
                ElseIf m_inCommentBlock Then
                    result = Ok()
                ElseIf IsValidType(name) Then
                    result = ParseTypeFirstDeclaration()
                Else
                    result = ParseAssignmentStatement()
                End If
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
            result = Ok()
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
        ElseIf IsNameStartCharacter(LookAhead) Then
            result = ParseVariable(GetType(System.Int32))
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

        m_LastTypeProcessed = Type.GetType("System.Int32")

        Return result
    End Function

    Private Function ParseString() As ParseStatus
        Dim result As ParseStatus
        
        ScanString()
        
        If TokenLength = 0 And Not m_EmptyStringFlag Then
            result = CreateError(1, "a valid string.")
        Else
            m_Gen.EmitString(CurrentToken)
            result = Ok()
        End If
        
        SkipWhiteSpace()
        
        Return result
    End Function

    Private Function ParseStringFactor() As ParseStatus
        Dim result As ParseStatus

        If LookAhead.Equals(""""c) Then
            result = ParseString()
        ElseIf IsNameStartCharacter(LookAhead) Then
            result = ParseVariable(GetType(System.String))
        ElseIf LookAhead.Equals("("c) Then
            SkipCharacter()

            result = ParseStringExpression()

            If result.Code = 0 Then
                If Not LookAhead.Equals(")"c) Then
                    result = CreateError(1, ")")
                Else
                    SkipCharacter()
                End If
            End If
        Else
            result = CreateError(1, "a string.")
        End If

        Return result
    End Function

    Private Function ParseConcatOperator() As ParseStatus
        Dim result As ParseStatus
        Dim currentoperator As String = CurrentToken
        
        SkipWhiteSpace()
        
        result = ParseStringFactor()
        
        If result.code = 0 Then
            m_Gen.EmitConcat()
        End If
        
        Return result
    End Function	

    Private Function ParseStringExpression() As ParseStatus
                
        Dim result As ParseStatus
        
        result = ParseStringFactor()
        
        Do While result.code=0 _
            AndAlso _
            IsConcatOperator(LookAhead) 
            
            ScanConcatOperator()
            
            If TokenLength=0 Then
                result = CreateError(1, "& or +")
            Else
                result = ParseConcatOperator()
                SkipWhiteSpace()
            End If
        Loop

        m_LastTypeProcessed = Type.GetType("System.String")

        Return result	
    End Function

    Private Sub GenerateRelOperation( _
            reloperator As String, _
            conditionType As Type)
            
        If conditionType.Equals( _
                Type.GetType("System.Int32")
            ) Then

            Select Case reloperator
                Case "=", "==", "==="
                    m_Gen.EmitEqualityComparison()
                Case ">"
                    m_Gen.EmitGreaterThanComparison()
                Case "<"
                    m_Gen.EmitLessThanComparison()
                Case ">=", "=>"
                    m_Gen.EmitGreaterThanOrEqualToComparison()
                Case "<=", "=<"
                    m_Gen.EmitLessThanOrEqualToComparison()
                Case "<>", "!=", "!=="
                    m_Gen.EmitInEqualityComparison()
            End Select

        ElseIf conditionType.Equals( _
                Type.GetType("System.String")
            ) Then

            Select Case reloperator
                Case "=", "==", "==="
                    m_Gen.EmitStringEquality()
                Case ">"
                    m_Gen.EmitStringGreaterThan()
                Case "<"
                    m_Gen.EmitStringLessThan()
                Case ">=", "=>"
                    m_Gen.EmitStringGreaterThanOrEqualTo()
                Case "<=", "=<"
                    m_Gen.EmitStringLessThanOrEqualTo()
                Case "<>", "!=", "!=="
                    m_Gen.EmitStringInequality()
            End Select
        ElseIf conditionType.Equals( _
                Type.GetType("System.Boolean")
            ) Then

            Select Case reloperator
                Case "=", "==", "==="
                    m_Gen.EmitEqualityComparison()
                Case "<>", "!=", "!=="
                    m_Gen.EmitInEqualityComparison()
            End Select
        End If
    End Sub

    Private Function ParseRelOperator() _
            As ParseStatus
            
        Dim result As ParseStatus
        Dim reloperator As String = CurrentToken
        Dim conditiontype As Type = m_LastTypeProcessed
        
        SkipWhiteSpace()
        
        ' The expression after a relational operator
        ' should match the type of the expression
        ' before it
        If conditiontype.Equals( _
                Type.GetType("System.Int32") _
            ) Then

            result = ParseNumericExpression()
        ElseIf conditiontype.Equals( _
                Type.GetType("System.String") _
            ) Then

            result = ParseStringExpression()
        ElseIf conditiontype.Equals( _
                Type.GetType("System.Boolean") _
            ) Then
            ' Boolean conditions support only equality or
            ' inequality comparisons
            Select Case reloperator
                Case "=","==","===", "<>", "!=", "!=="
                    result = ParseBooleanExpression()
                Case Else
                    result = CreateError(1, "= or <> operator")
            End Select
            
        Else
            result = CreateError(1, "an expression of type " & _
                            conditiontype.ToString())
        End If

        If result.Code = 0 Then
            GenerateRelOperation(reloperator, conditiontype)
        End If

        Return result
    End Function

    Private Function ParseCondition( _
                        Optional lastexpressiontype As Type = Nothing _
        ) As ParseStatus

        Dim result As ParseStatus

        If lastexpressiontype Is Nothing Then
            result = ParseExpression(Type.GetType("System.Boolean"))
        Else
            m_LastTypeProcessed = lastexpressiontype
            result = Ok()
        End If

        If result.Code = 0 Then

            ScanRelOperator()

            If TokenLength = 0 Then
                ' If the last type processed was a Boolean,
                ' a relational operator is not required
                If Not m_LastTypeProcessed.Equals( _
                                                GetType(Boolean) _ 
                                            ) Then
                    result = CreateError(1, "a relational operator.")
                End If
            Else
                result = ParseRelOperator()
                SkipWhiteSpace()
            End If
        End If

        Return result
    End Function

    Private Function ParseBooleanFactor( _
                        Optional lastexpressiontype As Type = Nothing _
        ) As ParseStatus

        Dim result As ParseStatus

        If LookAhead.Equals("("c) Then
            SkipCharacter()

            result = ParseBooleanExpression()

            If result.Code = 0 Then
                If Not LookAhead.Equals(")"c) Then
                    result = CreateError(1, ")")
                Else
                    SkipCharacter()
                End If
            End If
        Else
            result = ParseCondition(lastexpressiontype)
        End If

        SkipWhiteSpace()

        Return result
    End Function

    Private Function ParseNotOperator() _
        As ParseStatus

        Dim result As ParseStatus

        SkipWhiteSpace()

        result = ParseBooleanFactor()

        If result.Code = 0 Then
            m_Gen.EmitLogicalNot()
        End If

        Return result
    End Function

    Private Function ParseNotOperation( _
                        Optional lastexpressiontype As Type = Nothing _
        ) As ParseStatus

        Dim result As ParseStatus

        ' If lastexpressiontype has a value, it means we are halfway
        ' through parsing a boolean condition, and have just met a 
        ' relational operator. A NOT operator cannot appear at this
        ' point, so we just continue to parse a boolean factor
        If lastexpressiontype IsNot Nothing Then
            result = ParseBooleanFactor(lastexpressiontype)
        Else
            If IsNotOperator(LookAhead) Then
                ScanNotOperator()

                If TokenLength = 0 Then
                    result = CreateError(1, _
                        "NOT")
                Else
                    result = ParseNotOperator()
                End If
            Else
                result = ParseBooleanFactor()
            End If
        End If

        Return result
    End Function

    Private Function ParseAndOperator() _
            As ParseStatus

        Dim result As ParseStatus

        SkipWhiteSpace()

        result = ParseNotOperation()

        If result.Code = 0 Then
            m_Gen.EmitBitwiseAnd()
        End If

        Return result
    End Function

    Private Function ParseAndOperation( _
                        Optional lastexpressiontype As Type = Nothing _
        ) As ParseStatus

        Dim result As ParseStatus

        result = ParseNotOperation(lastexpressiontype)

        Do While result.Code = 0 And _
            IsAndOperator(LookAhead)

            ScanAndOperator()

            If TokenLength = 0 Then
                result = CreateError(1, _
                    "AND")
            Else
                result = ParseAndOperator()
                SkipWhiteSpace()
            End If
        Loop

        Return result
    End Function


    Private Function ParseOrOperator() _
            As ParseStatus

        Dim result As ParseStatus

        SkipWhiteSpace()

        result = ParseAndOperation()

        If result.Code = 0 Then
            m_Gen.EmitBitwiseOr()
        End If

        Return result
    End Function

    Private Function ParseOrOperation( _
                        Optional lastexpressiontype As Type = Nothing _
        ) As ParseStatus

        Dim result As ParseStatus

        result = ParseAndOperation(lastexpressiontype)

        Do While result.Code = 0 And _
            IsOrOperator(LookAhead)

            ScanOrOperator()

            If TokenLength = 0 Then
                result = CreateError(1, _
                    "OR")
            Else
                result = ParseOrOperator()
                SkipWhiteSpace()
            End If
        Loop

        Return result
    End Function

    Private Function ParseBooleanExpression( _
                        Optional lastexpressiontype As Type = Nothing _
        ) As ParseStatus

        Dim result As ParseStatus

        result = ParseOrOperation(lastexpressiontype)

        If result.Code = 0 Then
            m_LastTypeProcessed = _
                Type.GetType( _
                "System.Boolean")
        End If

        Return result
    End Function

    Private Function ParseExpression( _
                        Optional ByVal expressiontype _
                            As Type = Nothing) _
            As ParseStatus

        Dim result As ParseStatus

        ' Since we are doing the work of the scanner by using the
        ' lookahead character, we need to initialize the token
        ' builder
        ResetToken()

        If LookAhead.Equals(""""c) Then
            result = ParseStringExpression()
        ElseIf IsNumeric(LookAhead) Then
            result = ParseNumericExpression()
        ElseIf IsNameStartCharacter(LookAhead) Then
            result = ParseInitialName()
        ElseIf IsNotOperator(LookAhead) Then
            result = ParseBooleanExpression()
        ElseIf LookAhead.Equals("("c) Then
            ' For now, assuming only numeric expressions can use ()
            result = ParseNumericExpression()
        Else
            result = CreateError(1, "a boolean, numeric or string expression")
        End If

        If result.Code = 0 Then
            If expressiontype Is Nothing Then
                ' If a relational operator follows the expression just parsed
                ' we are in the middle of a Boolean expression. Our work is
                ' not yet done.
                If IsRelOperator(LookAhead) Then
                    ' Parse a boolean expression, letting it know that the
                    ' first part has already been parsed.
                    result = ParseBooleanExpression(m_LastTypeProcessed)
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseCommand() As ParseStatus
        Dim result As ParseStatus

        If TokenLength = 0 Then
            result = CreateError(1, "a valid command")
        Else
            Dim commandname As String = _
                    CurrentToken.ToLowerInvariant()

            If commandname = "comment" Then
                result = ParseCommentCommand()
            ElseIf commandname = "end"
                result = ParseEndCommand()
            ElseIf commandname = "rem" Then
                result = ParseRemCommand()
            ElseIf m_inCommentBlock Then
                ' Ignore rest of line
                SkipRestOfLine()
                ' All is good in a comment block
                result = Ok()
            ElseIf m_SwitchState IsNot Nothing _
                        AndAlso
                    m_SwitchState.CaseFlag Then

                If commandname = "case" Then
                    result = ParseCaseCommand()
                ElseIf commandname = "default" Then
                    result = ParseDefaultCommand()
                Else
                    result = CreateError(1, "Case or Default")
                End If
            Else
                If IsValidCommand(commandname) Then
                    Dim parser as CommandParser = _
                            m_commandTable(commandname)

                    result = parser()
                Else
                    result = CreateError(1, "a valid command")
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseBlock(ByVal newblock As Block) _
                            As ParseStatus

        Dim result As ParseStatus
        result = Ok()
        
        m_BlockStack.Push(newblock)

        Do While ScanLine()
            result = ParseLine()
            If result.Code <> 0 Then
                Exit Do
            End If
        Loop

        ' Block will end when the result code returned
        ' is -1
        If result.Code = -1 Then
            result = Ok()
            m_BlockStack.Pop()
        End If
        
        Return result
    End Function

    Private Function ParseTypeFirstDeclaration() _
                        As ParseStatus
        Dim result As ParseStatus

        Dim typename As String
        typename = CurrentToken

        If IsValidType(typename) Then
            ' Read a variable name
            SkipWhiteSpace()
            ScanName()

            If TokenLength = 0 Then
                result = CreateError(1, "a variable name.")
            Else
                Dim varname As String
                varname = CurrentToken

                If m_SymbolTable.Exists(varname) Then
                    result = CreateError(3, varname)
                Else
                    Dim symbol As New Symbol( _
                                        varname, _
                                        GetTypeForName(typename)
                    )

                    symbol.Handle = m_Gen.DeclareVariable(
                                        symbol.Name,
                                        symbol.Type
                                    )

                    m_SymbolTable.Add(symbol)

                    SkipWhiteSpace()

                    result = ParseDeclarationAssignment(symbol)
                    If result.Code = 0 Then
                        If Not EndOfLine Then
                            result = CreateError(1, "end of statement.")
                        End If
                    End If
                End If
            End If
        Else
            result = CreateError(1, "a valid type.")
        End If

        Return result
    End Function

    Private Function ParseAssignment(variable As Symbol) _
                        As ParseStatus

        Dim result As ParseStatus
     
        SkipWhiteSpace()
        
        Select Case variable.Type.ToString()
            Case "System.Int32"
                result = ParseNumericExpression()
            Case "System.String"
                result = ParseStringExpression()
            Case "System.Boolean"
                result = ParseBooleanExpression()
            Case Else
                result = CreateError(1, " a valid type.")
        End Select 
 
        If result.Code = 0 Then
            ' Generate assignment code
            m_Gen.EmitStoreInLocal(variable.Handle)
        End If

        Return result
    End Function

    Private Function ParseAssignmentStatement() _
                        As ParseStatus

        Dim result As ParseStatus
        
        Dim varname As String
        varname = CurrentToken

        If m_SymbolTable.Exists(varname) Then
            ' Read assignment operator
            SkipWhiteSpace()
            ScanAssignmentOperator()

            If TokenLength = 0 Then
                result = CreateError(1, "= or :=")
            Else
                Dim variable As Symbol
                variable = m_SymbolTable.Fetch(varname)

                result = ParseAssignment(variable)
            End If
        Else
            result = CreateError(4, varname)
        End If

        Return result
    End Function

    Private Function ParseVariable(type As Type) _
                        As ParseStatus

        Dim result As ParseStatus
 
        ' Try to read variable name
        ScanName()
        
        If TokenLength = 0 Then
            result = CreateError(1, "a variable.")
        Else
            Dim varname As String
            varname = CurrentToken

            If Not m_SymbolTable.Exists(varname) Then
                result = CreateError(4, varname)
            Else
                Dim variable As Symbol
                variable = m_SymbolTable.Fetch(varname)

                If Not variable.Type.Equals(type) Then
                    result = CreateError(5, variable.Name)
                Else
                    ' Emit the variable
                    m_Gen.EmitLoadLocal(variable.Handle)
                    SkipWhiteSpace()
                    
                    result = Ok()
                End If
            End If
        End If 

        Return result
    End Function

    Private Function ParseInitialName() As ParseStatus
        Dim result As ParseStatus

        ScanName()

        Dim varname As String
        varname = CurrentToken 

        If CurrentToken.ToLowerInvariant() = "not" Then
            ' This is a boolean. Backtrack and call
            ' boolean parser
            Backtrack()
            result = ParseBooleanExpression()
        Else
            If Not m_SymbolTable.Exists(varname) Then
                result = CreateError(4, varname)
            Else
                Dim variable As Symbol
                variable = m_SymbolTable.Fetch(varname)

                Select Case variable.Type.ToString()
                    Case "System.Int32"
                        ' Backtrack, and call numeric parser
                        Backtrack()
                        result = ParseNumericExpression()
                    Case "System.String"
                        ' Backtrack, and call string parser
                        Backtrack()
                        result = ParseStringExpression()
                    Case "System.Boolean"
                        ' Backtrack, parse variable because boolean factor
                        ' can't, then continue as a boolean expression
                        Backtrack()
                        result = ParseVariable(GetType(Boolean))
                        If result.Code = 0 Then
                            result = ParseBooleanExpression(GetType(Boolean))
                        End If
                    Case Else
                        result = CreateError(1, " an Integer or String variable.")
                End Select

            End If
        End If

        Return result
    End Function

    Private Function ParseDeclarationAssignment(ByVal variable As Symbol) _
                                                            As ParseStatus
        Dim result As ParseStatus
        
        If IsAssignmentCharacter(LookAhead) Then
            ScanAssignmentOperator()

            If TokenLength = 0 Then
                result = CreateError(1, "= or :=")
            Else
                result = ParseAssignment(variable)
            End If
        Else
            result = Ok()
        End If

        Return result
    End Function

    Private Function ParseNoiseWord( _
                        ByVal word As String,
                        Optional ByVal silent As Boolean = True
                    ) As ParseStatus

        Dim result As ParseStatus = Ok()

        If Not EndOfLine Then 
            ' If the next token is the noise word
            ' followed by a whitespace, skip it silently.
            ' If it isn't, raise and error if required, 
            ' but don't move the current position.

            Dim peekby As Integer = word.Length + 1
            Dim peektoken As String = PeekAhead(peekby)

            word = word.ToLowerInvariant()
            peektoken = peektoken.ToLowerInvariant()

            If peektoken.StartsWith(word) _
                    AndAlso _
                (
                    word.Length = peektoken.Length _
                        OrElse _
                    IsWhiteSpace( _
                        peektoken.Chars(peektoken.Length-1) _ 
                    ) _
                ) Then

                ' The token matches, and the next character
                ' is empty or whitespace. Scan and ignore
                '  the word
                ScanName()
                SkipWhiteSpace()

            Else
                If Not silent Then
                    result = CreateError(1, word)
                End If
            End If
        End if
        
        Return result
    End Function

    Private Function ParseLastNoiseWord( _
                        ByVal word As String _ 
                    ) As ParseStatus

        Dim result As ParseStatus

        ' The noise word should be the last word in
        ' a line
        result = ParseNoiseWord(word, False)
        If result.Code = 0 Then
            If Not EndOfLine Then
                result = CreateError(1, "end of statement")
            End If
        End If

        Return result
    End Function 
#End Region

    Public Function Parse() As ParseStatus
        Dim result As ParseStatus
        result = Ok()

        Do While ScanLine()
            result = ParseLine()
            If result.Code <> 0 Then
                Exit Do
            End If
        Loop

        If result.Code = 0 AndAlso _
                    (Not m_BlockStack.IsEmpty()) Then

            result = CreateError(1, _
                            "end " & _
                            m_BlockStack.CurrentBlock.BlockType _
            )
        End If

        Return result
    End Function

    Public Sub New( _
        ByVal newStream As TextReader, _
        ByVal newGen As CodeGen _
        )

        m_InputStream = newStream
        m_Gen = newGen

        InitErrors()
        InitTypes()
        InitLoops()
        InitCommands()
    End Sub
End Class
