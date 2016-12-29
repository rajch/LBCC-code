Option Strict On
Option Explicit On

Module Tester3
	Sub Main()
		Dim cg As New CodeGen("hello.exe")

		cg.EmitNumber(2)
		cg.EmitWriteLine()

		cg.EmitNumber(2)
		cg.EmitNumber(0)
		cg.EmitDivide()
		cg.EmitWriteLine()

		cg.Save()
	End Sub
End Module
