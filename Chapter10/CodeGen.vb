Option Strict On
Option Explicit On


Imports System
Imports System.Reflection
Imports System.Reflection.Emit
Imports System.Collections.Generic
Imports System.Diagnostics.SymbolStore

Public Class CodeGen
	Private m_ILGen As ILGenerator
	Private m_producedAssembly As AssemblyBuilder
	Private m_producedmodule As ModuleBuilder
	Private m_producedtype As TypeBuilder
	Private m_producedmethod As MethodBuilder
	Private m_SaveToFile As String
	Private m_LocalVariables As _
				New Dictionary(Of Integer, LocalBuilder)
	Private m_Labels As New List(Of Label)
	Private m_debugBuild As Boolean
	Private m_symbolStore As ISymbolDocumentWriter

	Public ReadOnly Property DebugBuild() As Boolean
		Get
			Return m_debugBuild
		End Get
	End Property

	Public Sub EmitNumber(ByVal num As Integer)
		m_ILGen.Emit(OpCodes.Ldc_I4, num)
	End Sub

	Public Sub EmitString(ByVal str As String)
		m_ILGen.Emit(OpCodes.Ldstr, str)
	End Sub

	Public Sub EmitAdd()
		m_ILGen.Emit(OpCodes.Add)
	End Sub

	Public Sub EmitSubtract()
		m_ILGen.Emit(OpCodes.Sub)
	End Sub

	Public Sub EmitMultiply()
		m_ILGen.Emit(OpCodes.Mul)
	End Sub

	Public Sub EmitDivide()
		m_ILGen.Emit(OpCodes.Div)
	End Sub

	Public Sub EmitEqualityComparison()
		m_ILGen.Emit(Opcodes.Ceq)
	End Sub

	Public Sub EmitGreaterThanComparison()
		m_ILGen.Emit(OpCodes.Cgt)
	End Sub

	Public Sub EmitLessThanComparison()
		m_ILGen.Emit(OpCodes.Clt)
	End Sub

	Private Sub NegateComparison()
		EmitNumber(0)
		EmitEqualityComparison
	End Sub

	Public Sub EmitInEqualityComparison()
		EmitEqualityComparison
		NegateComparison
	End Sub

	Public Sub EmitGreaterThanOrEqualToComparison()
		EmitLessThanComparison
		NegateComparison
	End Sub

	Public Sub EmitLessThanOrEqualToComparison()
		EmitGreaterThanComparison
		NegateComparison
	End Sub

	Private Sub EmitStringCompare()
		Dim stringtype As Type = _
			Type.GetType("System.String")

		Dim paramtypes() As Type = _
			{stringtype, stringtype}

		m_ILGen.Emit( _
				OpCodes.Call, _
				stringtype.GetMethod( _
						"CompareOrdinal", _
						paramtypes _
				) _
		)
	End Sub

	Public Sub EmitStringEquality()
		EmitStringCompare()
		EmitNumber(0)
		EmitEqualityComparison()
	End Sub

	Public Sub EmitStringInequality()
		EmitStringEquality()
		NegateComparison()
	End Sub

	Public Sub EmitStringGreaterThan()
		EmitStringCompare()
		EmitNumber(0)
		EmitGreaterThanComparison()
	End Sub

	Public Sub EmitStringLessThan()
		EmitStringCompare()
		EmitNumber(0)
		EmitLessThanComparison()
	End Sub

	Public Sub EmitStringGreaterThanOrEqualTo()
		EmitStringCompare()
		EmitNumber(-1)
		EmitGreaterThanComparison()
	End Sub

	Public Sub EmitStringLessThanOrEqualTo()
		EmitStringCompare()
		EmitNumber(1)
		EmitLessThanComparison()
	End Sub

	Public Sub EmitBitwiseAnd()
		m_ILGen.Emit(OpCodes.And)
	End Sub

	Public Sub EmitBitwiseOr()
		m_ILGen.Emit(OpCodes.Or)
	End Sub

	Public Sub EmitLogicalNot()
		NegateComparison()
	End Sub

	Public Sub EmitWriteLine()
		Dim inttype As Type = Type.GetType("System.Int32")
		Dim consoletype As Type = Type.GetType("System.Console")
		Dim paramtypes() As Type = {inttype}

		m_ILGen.Emit( _
			OpCodes.Call, _
			consoletype.GetMethod( _
				"WriteLine", paramtypes _
			) _
		)
	End Sub

	Public Sub EmitWriteLineString()
		Dim strtype As Type = Type.GetType("System.String")
		Dim consoletype As Type = Type.GetType("System.Console")
		Dim paramtypes() As Type = {strtype}

		m_ILGen.Emit( _
			OpCodes.Call, _
			consoletype.GetMethod( _
				"WriteLine", paramtypes _
			) _
		)
	End Sub

	Public Sub EmitWriteLineBoolean()
		Dim booltype As Type = _
			Type.GetType("System.Boolean")
			
		Dim consoletype As Type = _
			Type.GetType("System.Console")
		
		Dim paramtypes() As Type = _
			{ booltype}
		
		m_ILGen.Emit(Opcodes.Call, _
			consoletype.GetMethod( _
			"WriteLine", paramtypes))
	End Sub

	Public Function DeclareLabel() As Integer
		Dim lbl As Label = m_ILGen.DefineLabel()
		m_Labels.Add(lbl)
		Return m_Labels.Count - 1
	End Function

	Public Sub EmitLabel(ByVal labelNumber As Integer)
		Dim lbl As Label = m_Labels(labelNumber)
		m_ILGen.MarkLabel(lbl)
	End Sub

	Public Sub EmitBranchIfFalse(ByVal labelNumber As Integer)
		Dim lbl As Label = m_Labels(labelNumber)
		m_ILGen.Emit(OpCodes.Brfalse, lbl)
	End Sub

	Public Sub EmitBranchIfTrue(ByVal labelNumber As Integer)
		Dim lbl As Label = m_Labels(labelNumber)
		m_ILGen.Emit(OpCodes.Brtrue, lbl)
	End Sub

	Public Sub EmitBranch(ByVal labelNumber As Integer)
		Dim lbl As Label = m_Labels(labelNumber)
		m_ILGen.Emit(OpCodes.Br, lbl)
	End Sub

	Public Sub Save()
		' Emit a RETurn opcode, which is the last
		' opcode for any method
		m_ILGen.Emit(OpCodes.Ret)

		' Actually create the type in the module
		m_producedtype.CreateType()

		' Specify that when the produced assembly
		' is run, execution will start from
		' the produced method (MainMethod). Also, the
		' produced assembly will be a console
		' application.
		m_producedAssembly.SetEntryPoint( _
			m_producedmethod, _
			PEFileKinds.ConsoleApplication
		)

		' If this is a debug build, set the User
		' Entry Point, which is the point where a
		' debugger will start stepping
		If m_debugBuild Then
			Dim isMono As Boolean = _
				Type.GetType("Mono.Runtime") IsNot Nothing

			If Not isMono Then
			 	m_producedmodule.SetUserEntryPoint( _
			 							m_producedmethod _
			 	)
			End If
		End If

		m_producedAssembly.Save(m_SaveToFile)
	End Sub

	Public Sub EmitConcat()
		Dim stringtype As Type = Type.GetType("System.String")
		Dim paramtypes() As Type = { stringtype, stringtype }

		Dim concatmethod as MethodInfo = stringtype.GetMethod( _
				"Concat", paramtypes _
		)

		m_ILGen.Emit( _
			Opcodes.Call, _
			concatmethod
		)
	End Sub

	Public Function DeclareVariable( _
						Name As String, _
						VariableType As System.Type _
					) As Integer 
	
		Dim lb As LocalBuilder 

		lb = m_ILGen.DeclareLocal(VariableType)
		m_LocalVariables.Add(lb.LocalIndex, lb)

		If m_debugBuild Then
			lb.SetLocalSymInfo(Name)
		End If

		Return lb.LocalIndex
	End Function

	Public Sub EmitStoreInLocal( _
				Index As Integer _
			)
		
		Dim lb As LocalBuilder 
	
		lb = m_LocalVariables(Index)
		m_ILGen.Emit(OpCodes.Stloc, lb)

	End Sub

	Public Sub EmitLoadLocal( _
				Index As Integer _
			)
		
		Dim lb As LocalBuilder 
	
		lb = m_LocalVariables(Index)
		m_ILGen.Emit(OpCodes.Ldloc, lb)

	End Sub

	Public Sub EmitSequencePoint( _
					ByVal startLine As Integer, _
					ByVal startColumn As Integer, _
					ByVal endLine As Integer, _
					ByVal endColumn As Integer _
				)

		m_ILGen.MarkSequencePoint( _
						m_symbolStore, _
						startLine, _
						startColumn, _
						endLine, _
						endColumn _
				)

	End Sub

	Public Sub EmitNOP()
		m_ILGen.Emit(OpCodes.Nop)
	End Sub

	Public Sub EmitBreak()
		m_ILGen.Emit(OpCodes.Break)
	End Sub

	Public Sub New( _
			ByVal sourceFileName As String, _
			ByVal outputFileName As String, _
			ByVal assemblyName As String, _
			ByVal debugBuild As Boolean _
		)

		m_SaveToFile = outputFileName
		m_debugBuild = debugBuild

		' Compiling a CLR language produces an assembly.
		' An assembly has one or more modules.
		' Each module has one or more types:
		' 				(structures or classes)
		' Each type has one or more methods.
		' Methods are where actual code resides.

		' Create a new assembly. An assembly must have a name.
		' An assembly's name consists of up to four parts:
		' a simple name, a four-part version number, the default
		' Culture of the assembly, and optionally a public key token.
		' At the moment, we will only set the simple name. The
		' version number will default to 0.0.0.0, and the culture
		' will default to neutral.
		Dim an As New assemblyName()
		an.Name = assemblyName

		m_producedAssembly = _
			AppDomain.CurrentDomain.DefineDynamicAssembly( _
								an, AssemblyBuilderAccess.Save _
			)
			
		' In the newly created assembly, create a
		' module with the same name.
		' The last parameter of DefineDynamicModule
		' indicates whether debug information will
		' be emitted for this assembly
		m_producedmodule = m_producedAssembly.DefineDynamicModule( _
											assemblyName, _
											outputFileName, _
											m_debugBuild _
		)

		' If this is a debug build, create a symbol
		' store for the module.
		If m_debugBuild Then
			m_symbolStore = m_producedmodule.DefineDocument( _
											sourceFileName, _
											Nothing, _
											Nothing, _
											SymDocumentType.Text _
			)
		End If

		' In the newly created module, create a class called
		' "MainClass".

		m_producedtype = m_producedmodule.DefineType("MainClass")

		' In MainClass, create a Shared (static) method
		' with Public scope, called "MainMethod".
		m_producedmethod = 	m_producedtype.DefineMethod(
										"MainMethod", _
										MethodAttributes.Public _
											Or _
										MethodAttributes.Static, _
										Nothing, _
										Nothing _
		)

		' All IL code that we produce will be contained
		' in MainMethod.
		m_ILGen = m_producedmethod.GetILGenerator()
	End Sub
End Class
