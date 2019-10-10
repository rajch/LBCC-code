Option Strict On
Option Explicit On 

Imports System
Imports System.IO

Module Compiler
    Private Function ParseCommandLine( _
                                ByVal CmdArgs() As String _
                            ) As CommandArgs

        Dim result As New CommandArgs
        result.Message = "Please specify a source file."

        Dim switch As String
        Dim param As String
        
        For Each cmd As String In CmdArgs
            ' If the command-line paramater starts
            ' with a / or a -, it's a switch. If
            ' it contains a : , then the part after
            ' the : is the paramter to the switch
            If cmd.StartsWith("/") _
                         OrElse _
                    cmd.StartsWith("-") Then

                cmd = Mid(cmd, 2)
    
                Dim paramStart As Integer = _
                            InStr(cmd, ":")
        
                If paramStart <> 0 Then
                    switch = Left( _
                                cmd, _
                                paramStart - 1 _
                    )

                    param = Mid(cmd, paramStart + 1)
                Else
                    switch = cmd
                    param = ""
                End If

                Select Case switch.ToLowerInvariant()
                    Case "debug", "d"
                        result.DebugBuild = True
                    Case "out"
                        result.TargetFile = param
                    Case Else
                        result.Message = String.Format( _
                                "Error: Invalid switch - /{0}", _
                                switch _
                        )
                        Exit For
                End Select
            Else
                ' If the command-line parameter is not
                ' a switch, it is considered to be the
                ' name of the source file.
                result.SourceFile = cmd
                result.Valid = True
            End If
        Next

        Return result
    End Function

    Public Function Main(Byval CmdArgs() As String) As Integer
        Dim reader As TextReader
        Dim gen As CodeGen

        Dim status As ParseStatus
        Dim parser As Parser

        Console.WriteLine("Compiler for CLR")
		
		If CmdArgs.Length=0 Then
            reader = Console.In
            gen = New CodeGen( _
                        "",
                        "Test.exe", _
                        "Test", _
                        False
            )
        Else
            ' Parse the command line. If a source
            ' file has not been specified, show
            ' an error message and quit 
            Dim args As CommandArgs = ParseCommandLine(CmdArgs)
            If Not args.Valid Then
                Console.WriteLine(args.Message)
                Return 1
            End If

			If File.Exists(args.SourceFile) Then
				Dim finfo As New FileInfo(args.SourceFile)
				
				reader = New StreamReader( _
							finfo.FullName)
				
                ' If no target file has been specified, use the same
                ' name as the source file with the extension changed
                ' to .exe
                If args.TargetFile = "" Then
                    args.TargetFile = finfo.Name.Replace( _
                                            finfo.Extension, _
                                            ".exe" _
                    )
                Else
                    ' Validate the target filename
                    Dim finfoValid As FileInfo
                    Try
                        finfoValid = New FileInfo(args.TargetFile)
                    Catch ex As Exception
                        Console.WriteLine("Error: {0}", ex.Message)
                        Return 1
                    End Try

                    ' Target file name must have
                    ' an extension
                    If finfoValid.Extension = "" Then
                        args.TargetFile = finfoValid.Name & ".exe"
                    End If
                End If

                args.AssemblyName = finfo.Name.Replace( _
                            finfo.Extension, _
                            "" _
                )
                
				gen = new CodeGen( _
                        finfo.FullName, _
				        args.TargetFile, _
                        args.AssemblyName, _
                        args.DebugBuild _
                )
			Else
				Console.Write( _
					"Error: Could not find the file {0}.", _
                    args.SourceFile _
                )
				Return 1
			End If
		End If

        parser = New Parser(reader, gen)
        status = parser.Parse()

        With status
            If .Code <> 0 Then
                Console.WriteLine( _
                      "Error at line {0}, column {1} : {2}", _
                      .Row, _
                      .Column, _
                      .Description)
            Else
                gen.Save()
                Console.WriteLine("Done.")
            End If
        End With
        
        Return status.Code
    End Function
End Module