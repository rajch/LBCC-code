Option Strict On
Option Explicit On


Imports System
Imports System.Reflection
Imports System.Reflection.Emit


Public Class CodeGen
	Private m_ILGen As ILGenerator
	Private m_producedAssembly As AssemblyBuilder
	Private m_producedmodule As ModuleBuilder
	Private m_producedtype As TypeBuilder
	Private m_producedmethod As MethodBuilder
	Private m_SaveToFile As String

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

	Private Sub NegateComparison
		EmitNumber(0)
		EmitEqualityComparison
	End Sub

	Public Sub EmitInEqualityComparison
		EmitEqualityComparison
		NegateComparison
	End Sub

	Public Sub EmitGreaterThanOrEqualToComparison
		EmitLessThanComparison
		NegateComparison
	End Sub

	Public Sub EmitLessThanOrEqualToComparison
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

	Public Sub New(ByVal FileName As String)
		m_SaveToFile = FileName

		' Compiling a CLR language produces an assembly.
		' An assembly has one or more modules.
		' Each module has one or more types:
		' (structures or classes)
		' Each type has one or more methods.
		' Methods are where actual code resides.
		' Create an assembly called "MainAssembly".

		Dim an As New AssemblyName
		an.Name = "MainAssembly"

		m_producedAssembly = AppDomain.CurrentDomain.DefineDynamicAssembly( _
							an, AssemblyBuilderAccess.Save _
		)

		' In MainAssembly, create a module called
		' "MainModule".
		m_producedmodule = m_producedAssembly.DefineDynamicModule( _
							"MainModule", FileName, False
		)

		' In MainModule, create a class called
		' "MainClass".
		m_producedtype = m_producedmodule.DefineType("MainClass")

		' In MainClass, create a Shared (static) method
		' with Public scope, called "MainMethod".
		m_producedmethod = m_producedtype.DefineMethod( _
						"MainMethod", _
						MethodAttributes.Public Or MethodAttributes.Static, _
						Nothing, _
						Nothing _
		)

		' All IL code that we produce will be contained
		' in MainMethod.
		m_ILGen = m_producedmethod.GetILGenerator

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
End Class
