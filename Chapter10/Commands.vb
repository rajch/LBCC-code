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
    Private m_SwitchState As SwitchState
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
        AddCommand("print",         AddressOf ParsePrintCommand)
        AddCommand("rem",           AddressOf ParseRemCommand)
        AddCommand("comment",       AddressOf ParseCommentCommand)
        AddCommand("end",           AddressOf ParseEndCommand)
        AddCommand("dim",           AddressOf ParseDimCommand)
        AddCommand("var",           AddressOf ParseDimCommand)
        AddCommand("if",            AddressOf ParseIfCommand)
        AddCommand("else",          AddressOf ParseElseCommand)
        AddCommand("elseif",        AddressOf ParseElseIfCommand)
        AddCommand("while",         AddressOf ParseWhileCommand)
        AddCommand("exit",          AddressOf ParseLoopControlCommand)
        AddCommand("continue",      AddressOf ParseLoopControlCommand)
        AddCommand("repeat",        AddressOf ParseRepeatCommand)
        AddCommand("until",         AddressOf ParseUntilCommand)
        AddCommand("loop",          AddressOf ParseLoopCommand)
        AddCommand("for",           AddressOf ParseForCommand)
        AddCommand("next",          AddressOf ParseNextCommand)
        AddCommand("switch",        AddressOf ParseSwitchCommand)
        AddCommand("case",          AddressOf ParseCaseCommand)
        AddCommand("default",       AddressOf ParseDefaultCommand)
        AddCommand("fallthrough",   AddressOf ParseFallThroughCommand)
        AddCommand("stop",          AddressOf ParseStopCommand)
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

        GenerateSequencePoint()

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
                    ' For any End command other than
                    ' End Repeat, we need to generate
                    ' a NOP.
                    GenerateSequencePoint()
                    GenerateNOP()
                    
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

        GenerateSequencePoint()

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

                ' Note that the sequence point and the
                ' NOP are emitted AFTER the startpoint
                ' label. This will ensure that when the
                ' debugger jumps to the Else part of an
                ' If block, it will highlight the Else
                ' statement
                GenerateSequencePoint()
                GenerateNOP()

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

            ' Note that the sequence point is emitted AFTER the
            ' startpoint label. This will ensure that when the
            ' debugger jumps to an ElseIf in an If block, it will
            ' highlight the ElseIf statement
            GenerateSequencePoint()

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

        ' TODO: Explore this
        GenerateSequencePoint()
        
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

                GenerateSequencePoint()

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

            ' Note that the sequence point and the
            ' NOP are emitted AFTER the startpoint
            ' label. This will ensure that when the
            ' debugger jumps back, it will highlight
            ' the Repeat statement
            GenerateSequencePoint()
            GenerateNOP()

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
        If m_BlockStack.IsEmpty _
                OrElse _
            m_BlockStack.CurrentBlock.BlockType <> "repeat" Then

            result = CreateError(6, "Until")
        Else
            Dim repeatblock As Block = m_BlockStack.CurrentBlock

            SkipWhiteSpace()

            ' There should be a boolean expression after Until
            If EndOfLine Then
                result = CreateError(1, "a boolean expression")
            Else
                GenerateSequencePoint()

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

            ' Note that the sequence point and the
            ' NOP are emitted AFTER the startpoint
            ' label. This will ensure that when the
            ' debugger jumps back, it will highlight
            ' the Loop statement
            GenerateSequencePoint()
            GenerateNOP()

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

        GenerateSequencePoint()

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

    Private Function ParseSwitchCommand() As ParseStatus
        Dim result As ParseStatus

        SkipWhiteSpace()

        GenerateSequencePoint()

        ' Parse an Expression
        result = ParseExpression()

        If result.Code<>0 Then
            Return result
        End If

        ' Store old Switch state
        ' This behavior is necessary because Switch
        ' statements can be nested.
        Dim oldSwitchState As SwitchState = m_SwitchState

        ' Create new state, with an internal variable to
        ' store the Switch expression just parsed
        m_SwitchState = New SwitchState()
        With m_SwitchState
            .CaseFlag = True
            .DefaultFlag = False
            .FallThroughFlag = False
            .ExpressionVariable = _
                    m_Gen.DeclareVariable( _
                        MakeUniqueVariableName(), _
                        m_LastTypeProcessed _
                    )
            .ExpressionType = m_LastTypeProcessed
        End With

        ' Emit a store to the internal variable. ParseExpression
        ' has left the value of the expression on the stack
        m_Gen.EmitStoreInLocal(m_SwitchState.ExpressionVariable)

        ' There should be nothing else on the line
        SkipWhiteSpace()
        If Not EndOfLine Then
            Return CreateError(1, "end of statement")
        End If

        ' Pre-set the endpoint. This will not change
        Dim endpoint As Integer = m_Gen.DeclareLabel()

        ' Parse the block
        Dim switchBlock As New Block( _
                                "switch", _
                                -1, _
                                endpoint _
        )

        result = ParseBlock(switchBlock)

        If result.Code = 0 Then
            ' If there is a dangling startpoint
            ' emit it
            If switchBlock.StartPoint <> -1 Then
                m_Gen.EmitLabel(switchBlock.StartPoint)
            End If

            ' Emit the endpoint
            m_Gen.EmitLabel(switchBlock.EndPoint)

            ' Restore the old Switch State
            m_SwitchState = oldSwitchState
        End If

        Return result
    End Function

    Private Function ParseCaseCommand() As ParseStatus
        Dim result As ParseStatus

        If m_BlockStack.IsEmpty _
                OrElse _
            m_BlockStack.CurrentBlock.BlockType <> "switch" Then

            Return CreateError(6, "Case")
        End If

        Dim switchblock As Block = m_BlockStack.CurrentBlock

        ' If default statement has already been processed
        ' get out
        If m_SwitchState.DefaultFlag Then
            Return CreateError(6, "Case")
        End If

        ' This could be the first case statement, or the end
        ' of the previous case statement. In the latter case,
        ' we need to emit a jump to the current block's
        ' endingpoint before we start processing this case
        ' statement
        If Not m_SwitchState.CaseFlag Then
            ' This is the end of the previous case statement
            m_Gen.EmitBranch(switchblock.EndPoint)
        End If

        ' If the fallthrough flag is set, we need to jump over
        ' the condition. So we generate a label and a jump
        ' Fallthrough is weird
        Dim fallThroughLabel As Integer

        If m_SwitchState.FallThroughFlag Then
            fallThroughLabel = m_Gen.DeclareLabel()
            m_Gen.EmitBranch(fallThroughLabel)
        End If

        ' Emit the current startpoint if there is one
        If switchblock.StartPoint <> -1 Then
            m_Gen.EmitLabel(switchblock.StartPoint)
        End If

        ' Note that the sequence point is emitted AFTER
        ' the startpoint label is emitted. This will ensure
        ' that when the debugger jumps to a Case in a Select
        ' block, it will highlight the Case statement
        GenerateSequencePoint()
        
        ' Create new startpoint for the next case
        switchblock.StartPoint = m_Gen.DeclareLabel()

        SkipWhiteSpace()
        ' Depending on the type of the switch expression
        ' do the appropriate type of expression
        With m_SwitchState
            If .ExpressionType.Equals( _
                    GetType(System.Int32) _
                ) Then

                result = ParseNumericExpression()
            ElseIf .ExpressionType.Equals( _
                        GetType(System.String) _
                    ) Then

                result = ParseStringExpression()
            ElseIf .ExpressionType.Equals( _
                        GetType(System.Boolean) _
                    ) Then

                result = ParseBooleanExpression()
            Else
                result = CreateError(1, "a valid expression")
            End If

            If result.Code = 0 Then
                ' Emit the Switch Expression
                m_Gen.EmitLoadLocal(.ExpressionVariable)

                ' Emit comparison
                If .ExpressionType.Equals( _
                        GetType(System.String) _
                    ) Then

                    m_Gen.EmitStringEquality()
                Else
                    m_Gen.EmitEqualityComparison()
                End If

                ' If NOT equal, jump to the next Case
                m_Gen.EmitBranchIfFalse(switchBlock.StartPoint)

                ' Emit the Fallthrough label if the Fallthrough
                ' flag is set, and reset the flag
                If m_SwitchState.FallThroughFlag Then
                    m_Gen.EmitLabel(fallThroughLabel)
                    m_SwitchState.FallThroughFlag = False
                End If

                ' Reset the case flag
                m_SwitchState.CaseFlag = False
            End If
        End With

        Return result
    End Function

    Private Function ParseDefaultCommand() As ParseStatus

        If m_BlockStack.IsEmpty _
                OrElse _
            m_BlockStack.CurrentBlock.BlockType <> "switch" Then

            Return CreateError(6, "Default")
        End If

        Dim switchblock As Block = m_BlockStack.CurrentBlock

        ' If default statement has already been processed
        ' get out
        If m_SwitchState.DefaultFlag Then
            Return CreateError(6, "Default")
        End If

        ' This default could be the ONLY case in the
        ' switch block, or the last case

        ' In the first situation, we need to emit a jump to
        ' the current block's endpoint before we start
        ' processing this
        If Not m_SwitchState.CaseFlag Then
            ' This is the end of a previous case statement
            m_Gen.EmitBranch(switchblock.EndPoint)
        End If

        ' Emit the current startpoint if there is one
        If switchblock.StartPoint <> -1 Then
            m_Gen.EmitLabel(switchblock.StartPoint)
        End If

        ' Note that the sequence point and the
        ' NOP are emitted AFTER the startpoint
        ' label. This will ensure that when the
        ' debugger jumps to the Default part of
        ' a Case block, it will highlight the
        ' Default statement
        GenerateSequencePoint()
        GenerateNOP()

        ' Reset the Case flag
        m_SwitchState.CaseFlag = False

        ' Set the default flag
        m_SwitchState.DefaultFlag = True

        ' No more startpoints needed, as default is the
        ' last case
        switchblock.StartPoint = -1

        Return Ok()
    End Function

    Private Function ParseFallThroughCommand() As ParseStatus
        Dim result As ParseStatus

        If m_BlockStack.IsEmpty _
                OrElse _
            m_BlockStack.CurrentBlock.BlockType <> "switch" Then

            result = CreateError(6, "Fallthrough")
        Else
            m_SwitchState.CaseFlag = True
            m_SwitchState.FallThroughFlag = True
            result = Ok()
        End If

        Return result
    End Function

    Private Function ParseStopCommand() As ParseStatus
        Dim result As ParseStatus

        ' If not a debug build, ignore the Stop
        ' instruction
        If Not m_Gen.DebugBuild Then
            Return Ok()
        End If

        SkipWhiteSpace()
        
        ' If there is no When clause
        If EndOfLine Then
            ' Emit a break
            m_Gen.EmitBreak()
            GenerateSequencePoint()
            GenerateNOP()
            result = Ok()
        Else
            result = ParseWhenClause()
        
            If result.Code = 0 Then

                Dim afterBreakPoint As Integer = _
                                m_Gen.DeclareLabel()

                ' If the condition evaluates to
                ' false, jump past the break
                m_Gen.EmitBranchIfFalse(afterBreakPoint)

                ' Emit the break
                m_Gen.EmitBreak()
                GenerateSequencePoint()
                GenerateNOP()

                ' Emit the label
                m_Gen.EmitLabel(afterBreakPoint)
            End If
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