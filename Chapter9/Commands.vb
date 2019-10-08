Option Strict On
Option Explicit On

Imports System.Collections.Generic

Public Partial Class Parser

#Region "Fields"
    Private Delegate Function CommandParser() As ParseStatus
    Private m_commandTable As Dictionary(Of String, CommandParser)
    Private m_inCommentBlock As Boolean = False 
    Private m_typeTable As Dictionary(Of String, Type)
    Private m_ElseFlag As Boolean = False
    Private m_loopTable As List(Of String)
    Private m_ErrorTable As Dictionary(Of Integer, String)
#End Region

#Region "Helper Functions"
    Private Sub AddCommand( _
                commandName As String, _
                commandParser As CommandParser _
            )
        m_commandTable.Add( _
            commandName.ToLowerInvariant(), _
            commandParser _
        )
    End Sub

    Private Function IsValidCommand(commandName As String) As Boolean
        Return m_commandTable.ContainsKey(commandName.ToLowerInvariant())
    End Function

    Private Sub InitCommands()
        m_commandTable = New Dictionary(Of String, CommandParser)

        ' Add commands here
        AddCommand("print",     AddressOf ParsePrintCommand)
        AddCommand("rem",       AddressOf ParseRemCommand)
        AddCommand("comment",   AddressOf ParseCommentCommand)
        AddCommand("end",       AddressOf ParseEndCommand)
        AddCommand("dim",       AddressOf ParseDimCommand)
        AddCommand("var",       AddressOf ParseDimCommand)
        AddCommand("if",        AddressOf ParseIfCommand)
        AddCommand("else",      AddressOf ParseElseCommand)
        AddCommand("elseif",    AddressOf ParseElseIfCommand)
        AddCommand("while",     AddressOf ParseWhileCommand)
        AddCommand("exit",      AddressOf ParseLoopControlCommand)
        AddCommand("continue",  AddressOf ParseLoopControlCommand)
        AddCommand("repeat",    AddressOf ParseRepeatCommand)
        AddCommand("until",     AddressOf ParseUntilCommand)
        AddCommand("loop",      AddressOf ParseLoopCommand)
        AddCommand("for",       AddressOf ParseForCommand)
        AddCommand("next",      AddressOf ParseNextCommand)
    End Sub

    Private Sub AddType(typeName As String, type As Type)
        m_typeTable.Add(typeName.ToLowerInvariant(), type)
    End Sub

    Private Function IsValidType(typeName As String) As Boolean
        Return m_typeTable.ContainsKey(typeName.ToLowerInvariant())
    End Function

    Private Function GetTypeForName(typeName as String) As Type
        Return m_typeTable(typeName.ToLowerInvariant())
    End Function

    Private Sub InitTypes()
        m_typeTable = new Dictionary(Of String, Type)

        ' Add types here
        AddType("integer", GetType(System.Int32))
        AddType("string", GetType(System.String))
        AddType("boolean", GetType(System.Boolean))
        AddType("int", GetType(System.Int32))
        AddType("bool", GetType(System.Boolean))
    End Sub

    Private Sub AddLoop(loopName As String)
        m_loopTable.Add(loopName)
    End Sub

    Private Function IsValidLoop(loopName As String) As Boolean
        Return m_loopTable.Contains(loopName)
    End Function

    Private Sub InitLoops()
        m_loopTable = New List(Of String)

        ' Add loops here
        AddLoop("while")
        AddLoop("repeat")
        AddLoop("loop")
        AddLoop("for")
    End Sub

    Private Function CreateError( _
                        ByVal errorcode As Integer, _
                        ByVal errorDetail As String,  _
                        Optional ByVal beforeToken As Boolean = False
                    ) As ParseStatus

        Dim result As ParseStatus
        Dim message As String

        Dim errorpos As Integer

        If beforeToken Then
            ' Some errors happen
            ' at the scan position
            errorpos = m_CharPos
        Else
            ' Others happen after a token
            ' has been scanned
            errorpos = m_CharPos - TokenLength
        End If

        ' Columns should be 1-based for reporting
        errorpos = errorpos + 1
        
        If m_ErrorTable.ContainsKey(errorcode) Then
            message = String.Format( _
                        m_ErrorTable(errorcode), _
                        errorDetail _
            )
        Else
            message = String.Format(
                        "Unknown error code {0}",
                        errorcode
            )
        End if


        result = New ParseStatus(errorcode, _
                    message, _
                    errorpos, _
                    m_LinePos)


        Return result
    End Function

    Public Function BlockEnd() As ParseStatus
        Return CreateError(-1, "", True)
    End Function

    Private Function Ok() As ParseStatus
        Return CreateError(0, "", True)
    End Function

    Private Sub InitErrors()
        m_ErrorTable = New Dictionary(Of Integer, String)

        ' Add Error Numbers here
        m_ErrorTable.Add(-1,    "{0}")
        m_ErrorTable.Add(0,     "{0}")
        m_ErrorTable.Add(1,     "Expected {0}")
        m_ErrorTable.Add(2,     "{0}")
        m_ErrorTable.Add(3,     "Cannot redeclare variable '{0}'")
        m_ErrorTable.Add(4,     "Variable '{0}' not declared")
        m_ErrorTable.Add(5,     "Type mismatch for Variable '{0}'")
        m_ErrorTable.Add(6,     "'{0}' was unexpected at this time")
    End Sub
#End Region

#Region "Commands"
    Private Function ParsePrintCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        result = ParseExpression()
        If result.code = 0 Then
            If m_LastTypeProcessed.Equals( _
                                    Type.GetType("System.Int32") _
                                ) Then
                m_Gen.EmitWriteLine()
            ElseIf m_LastTypeProcessed.Equals( _
                                        Type.GetType("System.String") _
                                ) Then
                m_Gen.EmitWriteLineString()
            ElseIf m_LastTypeProcessed.Equals( _
                                        Type.GetType("System.Boolean") _
                                ) Then
                m_Gen.EmitWriteLineBoolean()
            End If
            
            If Not EndOfLine Then
                result = CreateError(1, "end of statement.")
            End If
        End If

        Return result
    End Function

    Private Function ParseRemCommand() As ParseStatus
        ' Ignore the rest of the line
        SkipRestOfLine()
        Return Ok()
    End Function

    Private Function ParseCommentCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()
        
        If Not EndOfLine Then
            result = CreateError(1, "end of statement")
        Else
            Dim newblock As Block
            Dim oldcommentstate As Boolean
        
            newblock = New Block("comment", -1, -1)
        
            oldcommentstate = m_inCommentBlock
        
            m_inCommentBlock = True
            result = ParseBlock(newblock)
        
            m_inCommentBlock = oldcommentstate
        End If

        Return result
    End Function

    Private Function ParseEndCommand() As ParseStatus
        Dim result As ParseStatus

        Dim block As Block = m_BlockStack.CurrentBlock

        ' We should be in a block when the End command
        ' is encountered.
        If Not block Is Nothing Then
            SkipWhiteSpace()
            ScanName()

            ' The token after the End command should 
            ' match the current block type
            If block.IsOfType(CurrentToken) Then
                ' If this is 'End Repeat'
                If CurrentToken.ToLowerInvariant() = "repeat" Then
                    result = ParseUntilCommand()
                Else
                    result = BlockEnd()
                End If
            Else
                ' Unless we are inside a comment block
                If m_inCommentBlock Then
                    result = Ok()
                Else
                    result = CreateError(1, block.BlockType)
                End If
            End If
        Else
            result = CreateError(2, "Not inside a block")
        End If

        Return result
    End Function

    Public Function ParseDimCommand() As ParseStatus
        Dim result As ParseStatus

        ' Read a variable name
        SkipWhiteSpace()
        ScanName()

        If TokenLength = 0 Then
            result = CreateError(1, "a variable name.")
        Else
            Dim varname As String
            varname = CurrentToken

            If m_SymbolTable.Exists(varname) Then
                ' Variable name already declared
                result = CreateError(3, CurrentToken)
            Else
                SkipWhiteSpace()

                ' Check and ignore "As"
                ParseNoiseWord("As")

                ' Read type
                ScanName()

                Dim typename As String
                typename = CurrentToken

                If TokenLength = 0 OrElse _
                        Not IsValidType(typename) Then
                    result = CreateError(1, "a valid type.")
                Else
                    Dim symbol As New Symbol( _
                                        varname,
                                        GetTypeForName(typename)
                    )

                    symbol.Handle = m_Gen.DeclareVariable( _
                                                symbol.Name, _
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
        End If

        Return result
    End Function

    Public Function ParseIfCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()
        result = ParseBooleanExpression()

        If result.Code=0 Then
            SkipWhiteSpace()

            result = ParseLastNoiseWord("Then")

            If result.Code=0 Then
                ' Store old value of Else flag for nesting
                Dim oldelseflag As Boolean = m_ElseFlag
                ' ElseFlag should be false 
                ' at start of a new If block
                m_ElseFlag = False

                Dim endpoint As Integer = m_Gen.DeclareLabel()

                ' If the condition just emitted is false, emit jump
                ' to endpoint
                m_Gen.EmitBranchIfFalse(endpoint)

                ' Parse the "If" block
                Dim ifblock As Block = New Block( _
                                    "if", _
                                    endpoint, _
                                    endpoint
                )

                result = ParseBlock(ifblock)

                ' If the block was successfully parsed, emit
                ' the endpoint label, and restore saved else
                ' flag for nesting
                If result.Code = 0  Then

                    ' If there is a dangling StartPoint, emit
                    ' it first
                    If ifblock.StartPoint<>ifblock.EndPoint _
                            AndAlso _
                        Not m_ElseFlag Then

                        m_Gen.EmitLabel(ifblock.StartPoint)
                    End If

                    m_Gen.EmitLabel(ifblock.EndPoint)
                    m_ElseFlag = oldelseflag
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseElseCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        ' There should be nothing after Else on a line
        If Not EndOfLine Then
            result = CreateError(1, "end of statement")
        Else
            Dim currBlock As Block = m_BlockStack.CurrentBlock

            ' We should be in an If block, and the Else flag should
            ' not be set
            If currBlock Is Nothing _
                    OrElse _
                currBlock.BlockType <> "if" _
                    OrElse _
                m_ElseFlag Then

                result = CreateError(6, "Else")
            Else
                ' The end point is the same as the start point
                ' This means that only an If statement has been
                ' parsed before. Generate a new end point
                If currBlock.EndPoint = currBlock.StartPoint Then
                    currBlock.EndPoint = m_Gen.DeclareLabel()
                End If

                ' Emit a jump to the end point
                ' Because a True If condition should never cause
                ' code in the Else block to be executed
                m_Gen.EmitBranch(currBlock.EndPoint)

                ' Emit the "start" point (of the Else block now)
                ' because a False If condition should cause
                ' execution to continue after the Else command.
                m_Gen.EmitLabel(currBlock.StartPoint)

                ' Set the Else flag
                m_ElseFlag = True

                result = Ok()
            End If
        End If
        
        Return result
    End Function

    Private Function ParseElseIfCommand As ParseStatus
        Dim result As ParseStatus

        Dim currBlock As Block = m_BlockStack.CurrentBlock
        
        ' The ElseIf command can only be in an If block, and the
        ' Else flag should not be set
        If currBlock Is Nothing _
                OrElse _
            currBlock.BlockType <> "if" _
                OrElse _
            m_ElseFlag Then

            result = CreateError(6, "ElseIf")
        Else
        
            ' If the endpoint is the same as the startpoint, this
            ' is the first ElseIf. Generate new endpoint. This 
            ' will mark the end of the If block.
            If currBlock.EndPoint = currBlock.StartPoint Then
                currBlock.EndPoint = m_Gen.DeclareLabel()
            End If

            ' Emit jump to the endpoint, because the ElseIf condition 
            ' and block should not be processed if the If condition 
            ' was true.
            m_Gen.EmitBranch(currBlock.EndPoint)

            ' Emit the "start" point. This marks the start of the
            ' ElseIf block
            m_Gen.EmitLabel(currBlock.StartPoint)

            SkipWhiteSpace()
            ' Parse the ElseIf condition
            result = ParseBooleanExpression()

            ' If successful
            If result.Code = 0 Then

                ' "Eat" the optional "Then"
                SkipWhiteSpace()

                result = ParseLastNoiseWord("Then")

                If result.Code = 0 Then        
                    ' Generate new "start" point. This will mark
                    ' the next elseif statement, or an else statement
                    currBlock.StartPoint = m_Gen.DeclareLabel()

                    ' If the condition is FALSE, jump to the 
                    ' "start" point
                    m_Gen.EmitBranchIfFalse(currBlock.StartPoint)
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseWhileCommand() As ParseStatus
        Dim result As ParseStatus
               
        SkipWhiteSpace()
        If EndOfLine Then
            result = CreateError(1, "a boolean expression")
        Else
            ' Generate and emit the startpoint
            Dim startpoint As Integer = m_Gen.DeclareLabel()
            m_Gen.EmitLabel(startpoint)

            ' Parse the Boolean expression
            result = ParseBooleanExpression()

            If result.Code = 0 Then
                ' There should be nothing else on the line
                If Not EndOfLine Then
                    result = CreateError(1, "end of statement")
                Else
                    ' Generate endpoint, and emit a conditional
                    ' jump to it
                    Dim endpoint As Integer = m_Gen.DeclareLabel()
                    m_Gen.EmitBranchIfFalse(endpoint)

                    ' Parse the block
                    Dim whileblock As New Block( _
                            "while", _
                            startpoint, _
                            endpoint _
                    )
                    result = ParseBlock(whileblock)

                    If result.Code = 0 Then
                        ' Emit jump back to startpoint
                        m_Gen.EmitBranch(whileblock.StartPoint)

                        ' Emit endpoint
                        m_Gen.EmitLabel(whileblock.EndPoint)
                    End If
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseLoopControlCommand() As ParseStatus
        Dim result As ParseStatus

        ' The current token is either Exit or Continue
        Dim cmdName As String = CurrentToken.ToLowerInvariant()

        ' It should be followed by a valid loop type, and we 
        ' should be inside that type of loop

        ' Read the Exit loop type
        SkipWhiteSpace()
        ScanName()

        Dim loopkind As String = CurrentToken.ToLowerInvariant()

        If Not IsValidLoop(loopkind) Then
            result = CreateError(1, "a valid loop type")
        Else
            ' Try to get the block from the stack
            Dim loopBlock As Block = _
                    m_BlockStack.GetClosestOuterBlock(loopkind)

            If loopBlock Is Nothing Then
                result = CreateError(6, cmdName & " " & loopkind)
            Else
                SkipWhiteSpace()

                If Not EndOfLine Then
                    ' There should be a When clause here    
                    result = ParseWhenClause()

                    If result.Code=0 Then
                        If cmdName = "exit" Then
                            ' Emit conditional jump to EndPoint
                            m_Gen.EmitBranchIfTrue(loopBlock.EndPoint)
                        ElseIf cmdName = "continue" Then
                            ' Emit conditional jump to StartPoint
                            m_Gen.EmitBranchIfTrue(loopBlock.StartPoint)
                        Else
                            result = CreateError(6, cmdName)
                        End if
                    End If
                Else
                    result = Ok()
                    
                    If cmdName = "exit" Then
                        ' Emit jump to EndPoint
                        m_Gen.EmitBranch(loopBlock.EndPoint)
                    ElseIf cmdName = "continue" Then
                        ' Emit jump to StartPoint
                        m_Gen.EmitBranch(loopBlock.StartPoint)
                    Else
                        result = CreateError(6, cmdName)
                    End if
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseRepeatCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()
        
        ' There should be nothing after Repeat
        If Not EndOfLine Then
            result = CreateError(1, "end of statement")
        Else
            Dim startpoint As Integer
            Dim endpoint As Integer

            ' Generate and emit startpoint
            startpoint = m_Gen.DeclareLabel()
            m_Gen.EmitLabel(startpoint)

            ' Generate endpoint
            endpoint = m_Gen.DeclareLabel()

            ' Parse the Repeat block
            Dim repeatblock As New Block( _ 
                                    "repeat", _
                                    startpoint, _
                                    endpoint _ 
            )


            result = ParseBlock(repeatblock)

            If result.Code = 0 Then
                ' Emit endpoint
                m_Gen.EmitLabel(repeatblock.EndPoint)
            End If
        End If

        Return result
    End Function

    Private Function ParseUntilCommand() As ParseStatus
        Dim result As ParseStatus

        ' Until is only valid if we are currently in a repeat block
        Dim repeatblock As Block = m_BlockStack.CurrentBlock

        If repeatblock.BlockType <> "repeat" Then
            result = CreateError(6, "Until")
        Else
            SkipWhiteSpace()

            ' There should be a boolean expression after Until
            If EndOfLine Then
                result = CreateError(1, "a boolean expression")
            Else
                result = ParseBooleanExpression()

                If result.Code = 0 Then
                    ' If the boolean condition evaluates to FALSE
                    ' we must jump to the start point.
                    m_Gen.EmitBranchIfFalse(repeatblock.StartPoint)

                    ' End the block
                    result = BlockEnd()
                End If
            End If
        End If

        Return result
    End Function

    Private Function ParseLoopCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        ' There should be nothing after Loop
        If Not EndOfLine Then
            result = CreateError(1, "end of statement")
        Else
            Dim startpoint As Integer
            Dim endpoint As Integer

            ' Generate and emit startpoint
            startpoint = m_Gen.DeclareLabel()
            m_Gen.EmitLabel(startpoint)

            ' Generate endpoint
            endpoint = m_Gen.DeclareLabel()

            ' Create the Loop block
            Dim loopBlock As New Block( _
                                "loop", _
                                startpoint, _
                                endpoint _
            )

            result = ParseBlock(loopBlock)

            If result.Code = 0 Then
                ' Emit jump to startpoint
                m_Gen.EmitBranch(loopBlock.StartPoint)

                ' Emit endpoint
                m_Gen.EmitLabel(loopBlock.EndPoint)
            End If
        End If

        Return result
    End Function

    Private Function ParseForCommand() As ParseStatus
        Dim result As ParseStatus

        ' Try to scan a variable name
        ResetToken()
        SkipWhiteSpace()
        ScanName()

        If TokenLength = 0 Then
            Return CreateError(1, "a variable")
        End If

        ' Check if variable exists, and is numeric
        Dim varname As String = CurrentToken.ToLowerInvariant()
        If Not m_SymbolTable.Exists(varname) Then
            Return CreateError(4, CurrentToken)
        End If

        Dim variable As Symbol = m_SymbolTable.Fetch(varname)
        If Not variable.Type.Equals(GetType(Integer)) Then
            Return CreateError(1, "a numeric variable")
        End If

        ' Parse the next two tokens as an assignment statement
        result = ParseAssignmentStatement()

        If result.Code<>0 Then
            Return result
        End If

        ' Parse the To Clause
        SkipWhiteSpace()
        ScanName()

        If TokenLength=0 OrElse _
            CurrentToken.ToLowerInvariant()<>"to" Then

            Return CreateError(1, "To")
        End If

        ' Parse the <end> numberic expression
        SkipWhiteSpace()
        result = ParseNumericExpression()

        If result.Code<>0 Then
            Return result
        End If

        ' Declare a uniquely named variable and emit a store
        ' to it, thus storing the <end> value.
        Dim endVariableName As String = MakeUniqueVariableName()
        Dim endVariable As Integer =  _
                m_Gen.DeclareVariable( _
                        endVariableName, _
                        GetType(System.Int32) _
                )

        m_Gen.EmitStoreInLocal(endVariable)

        ' Process the option <step> value
        Dim defaultStep As Boolean = True
        Dim stepVariableName As String
        Dim stepVariable As Integer

        SkipWhiteSpace()
        If Not EndOfLine Then
            ScanName()

            If TokenLength=0 OrElse _
                CurrentToken.ToLowerInvariant()<>"step" Then
                Return CreateError(1, "step")
            End If

            SkipWhiteSpace()
            result = ParseNumericExpression()

            If result.Code<>0 Then
                Return result
            End If

            stepVariableName = MakeUniqueVariableName()
            stepVariable = m_Gen.DeclareVariable( _
                                stepVariableName, _
                                GetType(System.Int32) _
                        )

            m_Gen.EmitStoreInLocal(stepVariable)
            defaultStep = False
        End If

        ' At this point, initialization of the loop is complete
        ' We need to jump to the comparison. The jump is required
        ' because at this point, we will emit the code that
        ' increments the loop counter, and this incrementing should
        ' not happen on the first iteration of the loop. So, we
        ' declare a label that marks the start of the comparison,
        ' and emit a jump to it.
        Dim comparisonStart As Integer = m_Gen.DeclareLabel()
        m_Gen.EmitBranch(comparisonStart)

        ' Declare startpoint and endpoint, and emit the startpoint.
        ' At this point, we will emit code that increments the loop
        ' counter, and then code that checks the counter against the
        ' <end> value. This is the real starting point of the loop.
        Dim startpoint As Integer = m_Gen.DeclareLabel()
        Dim endpoint As Integer = m_Gen.DeclareLabel()

        m_Gen.EmitLabel(startpoint)

        ' Increment the counter variable by the step value
        m_Gen.EmitLoadLocal(variable.Handle)
        
        If defaultStep Then
            m_Gen.EmitNumber(1)
        Else
            m_Gen.EmitLoadLocal(stepVariable)
        End If
        
        m_Gen.EmitAdd()
        m_Gen.EmitStoreInLocal(variable.Handle)

        ' The comparison starts here. So, we emit the comparisonStart
        ' label.
        m_Gen.EmitLabel(comparisonStart)

        ' If the default step value has not been used
        ' Emit a check to see whether the step value is negative
        ' If step is negative, emit <end> and <variable> in that order,
        ' and jump to the comparison. In all other cases, emit 
        ' <variable> and <end> in that order.
        ' Then do the comparison
        Dim normalOrderLabel As Integer
        Dim actualCompareLabel As Integer

        If Not defaultStep Then
            With m_Gen
                normalOrderLabel = .DeclareLabel()
                actualCompareLabel = .DeclareLabel()
                ' Check if step is negative
                .EmitLoadLocal(stepVariable)
                .EmitNumber(0)
                .EmitGreaterThanComparison()
                ' It is. Normal comparison
                .EmitBranchIfTrue(normalOrderLabel)
                ' It isn't. Emit end and counter variables,
                ' in that order
                .EmitLoadLocal(endVariable)
                .EmitLoadLocal(variable.Handle)
                ' Jump to actual comparison
                .EmitBranch(actualCompareLabel)
                ' The normal order of comparison
                ' will happen now
                .EmitLabel(normalOrderLabel)
            End With
        End If

        ' Emit counter and end variables, in that order 
        m_Gen.EmitLoadLocal(variable.Handle)
        m_Gen.EmitLoadLocal(endVariable)
        
        ' The negative step case jumps to this point
        If Not defaultStep Then
            m_Gen.EmitLabel(actualCompareLabel)
        End If

        m_Gen.EmitGreaterThanComparison()
        ' If the comparison succeeds, the loop is over
        m_Gen.EmitBranchIfTrue(endPoint)

        ' Now, Parse the block
        Dim forBlock As Block = New Block( _ 
                                    "for", _
                                    startpoint, _
                                    endpoint _
        )

        result = ParseBlock(forBlock)
        If result.Code = 0 Then
            ' Jump to the startpoint, where the <variable> 
            ' will be incremented and the comparison will be done
            m_Gen.EmitBranch(forBlock.StartPoint)

            ' Emit the endpoint
            m_Gen.EmitLabel(forBlock.EndPoint)
        End If

        Return result
    End Function

    Private Function ParseNextCommand() As ParseStatus
        Dim result As ParseStatus
        
        If (Not m_BlockStack.IsEmpty) _
                AndAlso _
            m_BlockStack.CurrentBlock.BlockType = "for" Then

            result = BlockEnd()
        Else
            result = CreateError(6, "Next")
        End If

        Return result
    End Function
#End Region

#Region "Clauses"
    Private Function ParseWhenClause() As ParseStatus
        Dim result As ParseStatus

        ScanName()

        If CurrentToken.ToLowerInvariant()<>"when" Then
            result = CreateError(1, "when")
        Else
            SkipWhiteSpace()

            result = ParseBooleanExpression()
        End If

        Return result
    End Function
#End Region

End Class