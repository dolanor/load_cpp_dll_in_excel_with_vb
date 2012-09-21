'ThisWorkbook
Private Declare Sub Dummy Lib "libforexcel.dll" ()

Private Sub Workbook_Open()
	ChDrive ActiveWorkbook.Path
	ChDir ActiveWorkbook.Path
	Dummy
End Sub
