Option Strict On
Option Explicit On

Module Tester2
	Sub Main()
		Dim cg As New CodeGen("hello.exe")

		cg.EmitNumber(2)
		cg.EmitWriteLine()
		cg.EmitNumber(2)
		cg.EmitAdd()

		cg.Save()
	End Sub
End Module
