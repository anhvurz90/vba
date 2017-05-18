4.VB Grammar: {
	4.1.Objects: {
		- Object types:
			+ Application,
			+ workbook (file)
			+ worksheet
			+ range
			+ chart
		
		- Examples:
			+ Application.Workbooks
			+ Workbooks.Item(1)
			+ Workbooks("Seles.xls")
			+ Workbooks("Seles.xls").Worksheets("Sheet1").Range("B3")
	}
	4.2.Methods: {
		- for Range:
			Activate, Clear, Copy, Cut, Delete, Select
		- Example:
			Range("B3").Select
	}
	4.3.Properties: {
		- for Range:
			ColumnWidth, Font, Formula, Text, Value
		- Example:
			+ ActiveCell.FormulaR1C1 = "Nguyen Van Hung"
			+ Range("C3").ColumnWidth = 14
	}
	4.4.Variables: {
		- Case insensitive
		- Dim variable_name As data_type:
		- Byte, Boolean, Integer, Long, SIngle, Currency, Date,
		Object, String, Variant
		- Example:
			+ Sub DataType() 
				Dim Age As Integer 'Age is an integer'
				Dim eName As String
				Age = 22
				eName = "Nguyen Van Hung"
				MsgBox "Name: " & eName & vbTab & "Age: " & Age
			End Sub
			
			+ Const Pi = 3.14159
			
			+ Sub VD_Bienso() 
				Dim Marks as Range
				Dim C, D As Integer
				Set Marks = Range("B1:B10")
				D = 0
				For Each C in Marks
					If C.value < 40 then
						D = D + 1
					End If
				Next C
				MsgBox "New value of D: " & D
			End Sub
	}
	4.5.Arrays: {
		- Dim Arr(4)
		- Dim Myfriends(1 to 30) As String
		- Dim Noisuy(1 to 20, 1 to 30) As Single
		- Dim Array("Michael", "David", "Peter", "Jackson")
		- UBound: last element, LBound: first element
		- 	Option Base 1 ( index of array from 1 instead of 0)
			Sub assignArray()
				Dim Arr(4) As String
				Arr(1) = "Thang 1"
				Arr(2) = "Thang 2"
				Arr(3) = "Thang 3"
				Arr(4) = "Thang 4"
				MsgBox Arr(1) & Chr(13) & Arr(2) & vbNewLine & Arr(3) & vbCrLf & Arr(4)
			End Sub
	}
	4.6.With - End With: {}
}
8.Tham chieu den o^ va` vu`ng: {
	8.1.Tham chieu kieu A1: {
		- Ex:
			Range("B1")
			Range("B1:B6")
			Range("B2:B7, F4:K30")
			Range("C:C")
			Range("7:7")
			Range("D:G")
			Range("2:6")
			Range("2:2, 5:5, 8:8")
			Range("B:B, D:D, G:G")
			
			Range("A1:A3").Select
		- Ex2: 
			Sub Thunghiem()
				Workbook("Popupmenu").Sheets("Sheet1").Range("B3").Select
				ActiveCell.FormulaR1C1 = "Bo mon DCCT"
				Selection.Font.Bold = True
				Selection.Font.Italic = True
				Selection.Font.ColorIndex = 3
				With Selection.Interior
					.ColorIndex = 6
					.Pattern = xlSolid
				End With
				Range("B4").Select
			End Sub
	}
	8.2.Index Numbers: {
		- Ex:
			Cells(4, 1)
			Cells() -> all cells in sheet
			Worksheets("Sheet2").Cells(3, 2).Value = 2000
	}
	8.3.Rows  and Columns: {
		- Ex:
			Rows(4)
			Rows
			Columns(4) == Columns("D")
			Columns
			Worksheets("Week4").Rows(2).Font.Bold = True
			
	}
	8.4.Named ranges: {
		- Range("[Quanly.xls]Danhsach!Congty").Font.Bold = True
		- Range("Congty").Font.Bold = False
		
		- Workbooks("Congty.xls").Names.Add Name:="Congty",_RefersTo:="Danhsach!D1:D10"
		   Range("Congty").Font.Italic = True
	}
	8.5.Multiple ranges: {
		Worksheets("Bang").Range("A1:C3,H4:L8,P14:Z34").ClearContents
		Range("Danhsach1, Danhsach2, Danhsach3").ClearContents
	}
	8.6.Offset cells: {
		Sub Offset() 
			Range("B1").Activate
			ActiveCell.Offset(1, 1).Font.ColorIndex = 3
			ActiveCell.Offset(4, 1).Font.Bold = True
			ActiveCell.Offset(8, 1).Value = "Xi nghiep khao sat dia ky thuat"
			ActiveCell.Offset(8, 1).Font.Size = 12
			Range("E9").Activate
			ActiveCell.Offset(-1, -2).Font.Italic = True
		End Sub
	}
	8.7.R1C1: {
		= Row Column
		- Ex1: assign B5 = Sum("B2:B4")
			Range("B5").Select
			ActiveCell.FormulaR1C1 = "=Sum(R[-3]C:R[-1]C)"
		- Ex2: assign D5 = F2 - F4
			Range("B5").Formula = "=R[-3]C[2]-R[-1]C[2]"
		- Ex3: 
			Range("G6").Select
			ActiveCell.FormulaR1C1 = "=R[-1]C*R[-2]C"
			Selection.Copy
			Selection.PasteSpecial Paste:= xlValues
			Application.CutCopyMode = False
	}
}
9.Conditions: {
	9.1.If: {
		Sub Hocluc()
			Sheets("Sheet1").Select
			Range("A1").Select
			If ActiveCell > 8 Then
				Range("B2").Value = "Hoc luc gioi"
			ElseIf ActiveCell > 6.5 Then
				Range("B2").Value = "Hoc luc kha"
			ElseIf ActiveCell > 5 Then
				Range("B2").Value = "Hoc luc trung binh"
			Else
				Range("B2").Value = "Hoc luc kem"
			End If
		End Sub
		
		Sub User_If()
			If ActiveCell.Value = "" Then Exit Sub
			If ActiveCell.Value >= 40 Then
				ActiveCell.Offset(0, 1).Value = "Tot"
			Else
				ActiveCell.Offset(0, 1).Value = "Xau"
			End If
		End Sub
	}
	9.2.Select Case: {
		Sub Trangthai()
			Sheets("Sheet1").Select
			Doset = Cells(2, 2).Value
			Select Case Doset:
				Case 1, 1 to 10
					Cells(2, 3).Value = "Chay"
				Case 0.75 to 1
					Cells(2, 3) = "Deo chay"
				Case 0.5 to 0.75
					Cells(2, 3) = "Deo mem"
				Case 0.25 to 0.5
					Cells(2, 3) = "Deo cung"
				Case 0 to 0.25
					Cells(2, 3) = "Nua cung"
				Case < 0
					Cells(2, 3) = "Cung"
			End Select
		End Sub
	}
	9.3.And, Or
}
10.Dialog in VBA: {
	10.1.Message Box: {
		Sub Nhangui()
			Dim Truonghop As Integer
			Truonghop = MsgBox("Ban co muon thoat khoi chuong trinh ko",
				vbYesNoCancel + vbQuestion + vbDefaultButton1, 
				"Chuong trinh tinh luong")
			if Truonghop = vbYes THen
				MsgBox "Ban vua chon nut Yes.", vbInformation
			ElseIf Truonghop = vbNo Then
				MsgBox "Ban vua chon nut No.", vbCritical
			ElseIf Truonghop = vbCancel Then
				MsgBox "Ban vua bam nut Cancel.", vbExclamation
			End If
		End Sub
	}
	10.2.Input Box: {
		Sub Input()
			Dim Dangmang
			Dim Cot, Hang As Integer
			Set Mang = Application.InputBox("Vao mang:", "Title", Type:= 8)
			Cot = Dangmang.Columns.Count
			Hang = Dangmang.Rows.Count
			MsgBox "So cot la: " & Cot
			MsgBox "So hang la: " & Hang
			MsgBox "Dia chi o dau la: " & Dangmang.Cells(1, 1).Address
			MsgBox "Dia chi o cuoi la: " & Dangmang.Cells(Cot, Hang).Address
		End Sub
	}
}
11.Loop: {
	11.1.Do ... Loop: {
		Sub Do()
			m = 4
			Do
				m = m + 1
				MsgBox m
				If m > 10 Then Exit Do
			Loop
		End Sub
	}
	11.2.Do While ... Loop: {
		Sub DoW_Loop()
			i = 1
			Do While i <= 10
				Cells(i, 1) = i
				i = i + 1
				MsgBox i
			Loop
		End Sub
	}
	11.3.Do Until ... Loop: {
		Sub Do_LoopW()
			i = 1
			Do
				Cells(i, 3) = i
				i = i + 1
				MsgBox i
			Loop While i <= 10
		End Sub
	}
	11.4.Do Until ... Loop: {
		Sub DoU_Loop()
			i = 1
			Do Until i = 10
				Cells(i, 5) = i
				i = i + 1
				MsgBox i
			Loop
		End Sub
	}
	11.5.For ... Next: {
		- No Step Usage:
			Sub ForNext() 
				For i = 1 To 5
					Cells(10, i) = i
					MsgBox i
				Next
			End Sub
			
			Sub ForNext_Step()
				For i = 1 to 7 Step 2
					Cells(12, i) = i
					MsgBox i
				Next
			End Sub
			
		- Step Usage:
		
			
	}
	11.6.For Each ... Next: {
		Sub ShowWorkSheets() 
			Dim mySheet As Worksheet
			Dim i As Integer : i = 1
			For Each mySheet in Worksheets
				MsgBox mySheet.Name
				i = i + 1
			Next mySheet
			MsgBox "So sheet in workbook is: " & i
		End Sub
	}
	11.7.Exit: {
		Sub ExitStatementDemo() 
			Dim I, MyNum
			Do
				For I = 1 to 1000
					MyNum = Int(Rnd * 1000)
					Select Case MyNum
						Case 7: Exit For
						Case 29: Exit Do
						Case 54: Exit Sub
					End Select
				Next I
			Loop
		End Sub
	}
	11.8.Loop in loop: {
		Sub CellsExample()
			For i = 1 To 5
				For j = 1 To 5
					Cells(i, j) = "Row " & i & " Col " & j
				Next j
			Next i
		End Sub
	}
}
12.Addin: {
	- Alt + F11
	- Function BMI(w, h)
		BMI = W * 703 / h ^ 2
	  End Function
}
13.Request: {
	- Ex1:
		Dim oRequest As Object
		Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
		oRequest.Open "GET", ...url...
		oRequest.Send
		MsgBox oRequest.ResponseText
	
	- Ex2:
		Dim oRequest As Object
		Set oRequest = CreateObject("WinHttp.WinHttpRequest.5.1")
		oRequest.Open "POST", "https://mysite.com/licensing/getstatus.php"
		oRequest.SetRequestHeader "Content-Typ", "application/x-www-form-urlencoded"
		oRequest.Send "var1=123&anothervar=test"
		MsgBox oRequest.ResponseText
}
14.Tip & tricks: {
	- Edit VBA Code: Alt + F11.
	- Choose Macro: Alt + F8
}