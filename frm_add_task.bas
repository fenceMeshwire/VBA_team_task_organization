Option Explicit
' -------------------------------------------------------------------------
' Please make the form first and attach this code to the form.
' Make and rename two WorkSheets to Sheet1 and Sheet2.
' -------------------------------------------------------------------------
Private Sub cmb_Cancel_Click()

Unload frm_add_task

End Sub

' -------------------------------------------------------------------------
Private Sub cmb_OK_Click()

Dim intCounter As Integer
Dim lngCellFree As Long

lngCellFree = Sheet1.Cells(Sheet1.Rows.Count, 1).End(xlUp).Row + 1

With Sheet1
  ' Add data to spreadsheet:
  .Cells(lngCellFree, 1).Value = Me.txt_task_description.Value
  If Me.chk_daily = True Then .Cells(lngCellFree, 2).Value = "X"
  If Me.chk_weekly = True Then .Cells(lngCellFree, 3).Value = "X"
  If Me.chk_monthly = True Then .Cells(lngCellFree, 4).Value = "X"
  If Me.chk_quarterly = True Then .Cells(lngCellFree, 5).Value = "X"
  If Me.chk_yearly = True Then .Cells(lngCellFree, 6).Value = "X"
  ' Format cells:
  For intCounter = 2 To 6
    .Cells(lngCellFree, intCounter).HorizontalAlignment = xlCenter
  Next intCounter
  .Cells(lngCellFree, 1).HorizontalAlignment = xlLeft
  ' Draw border around task list
  .Range("A" & lngCellFree).BorderAround Weight:=xlThin
  .Range("B" & lngCellFree).BorderAround Weight:=xlThin
  .Range("C" & lngCellFree).BorderAround Weight:=xlThin
  .Range("D" & lngCellFree).BorderAround Weight:=xlThin
  .Range("E" & lngCellFree).BorderAround Weight:=xlThin
  .Range("F" & lngCellFree).BorderAround Weight:=xlThin
  ' Clear task value in the form:
  Me.txt_task_description.Value = ""
  Me.cmb_member.Value = ""
End With

End Sub
' -------------------------------------------------------------------------
Private Sub UserForm_Initialize()

Dim intCounter As Integer
Dim varNamesArray As Variant

varNamesArray = getNamesArray()

For intCounter = LBound(varNamesArray) To UBound(varNamesArray)
  Me.cmb_member.AddItem varNamesArray(intCounter)
Next intCounter

End Sub
' -------------------------------------------------------------------------
Function getNamesArray() As Variant

Dim intCounter As Integer
Dim varEmployee, varSortedArray As Variant

varEmployee = Array("Sam", "Sally", "John", "Andrea", "Richard")
varSortedArray = get_sorted_array_az(varEmployee)

getNamesArray = varSortedArray

End Function
' -------------------------------------------------------------------------
Function get_sorted_array_az(ByRef varEmployee As Variant) As Variant

Dim a, b As Long
Dim strStorage As String

For a = LBound(varEmployee) To UBound(varEmployee) - 1
  For b = a + 1 To UBound(varEmployee)
    If UCase(varEmployee(a)) > UCase(varEmployee(b)) Then
      strStorage = varEmployee(b)
      varEmployee(b) = varEmployee(a)
      varEmployee(a) = strStorage
    End If
  Next b
Next a

get_sorted_array_az = varEmployee

End Function


