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
  If Me.txt_task_description.Value <> "" Then
    .Cells(lngCellFree, 1).Value = Me.txt_task_description.Value
  Else
    MsgBox ("Please enter a task.")
    Exit Sub
  End If
  
  If Me.cmb_member.Value <> "" Then
    .Cells(lngCellFree, 2).Value = Me.cmb_member.Value
  Else
    MsgBox ("Please enter a team member.")
    Exit Sub
  End If
  
  If Me.chk_daily = True Then .Cells(lngCellFree, 3).Value = "X"
  If Me.chk_weekly = True Then .Cells(lngCellFree, 4).Value = "X"
  If Me.chk_monthly = True Then .Cells(lngCellFree, 5).Value = "X"
  If Me.chk_quarterly = True Then .Cells(lngCellFree, 6).Value = "X"
  If Me.chk_yearly = True Then .Cells(lngCellFree, 7).Value = "X"
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
  .Range("G" & lngCellFree).BorderAround Weight:=xlThin
  ' Clear task value in the form:
  Me.txt_task_description.Value = ""
  Me.cmb_member.Value = ""
  Me.chk_daily = False
  Me.chk_yearly = False
  Me.chk_monthly = False
  Me.chk_quarterly = False
  Me.chk_weekly = False
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

Call create_heading

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

Private Sub create_heading()

Dim intCounter As Integer

With Sheet1

.Cells(1, 1).Value = "Task"
.Cells(1, 2).Value = "Team member"
.Cells(1, 3).Value = "Daily"
.Cells(1, 4).Value = "Weekly"
.Cells(1, 5).Value = "Monthly"
.Cells(1, 6).Value = "Quarterly"
.Cells(1, 7).Value = "Yearly"

.Range("A1").BorderAround Weight:=xlThin
.Range("B1").BorderAround Weight:=xlThin
.Range("C1").BorderAround Weight:=xlThin
.Range("D1").BorderAround Weight:=xlThin
.Range("E1").BorderAround Weight:=xlThin
.Range("F1").BorderAround Weight:=xlThin
.Range("G1").BorderAround Weight:=xlThin

.Range("A1").Font.Bold = True
.Range("B1").Font.Bold = True
.Range("C1").Font.Bold = True
.Range("D1").Font.Bold = True
.Range("E1").Font.Bold = True
.Range("F1").Font.Bold = True
.Range("G1").Font.Bold = True

For intCounter = 1 To 7
  .Cells(1, intCounter).HorizontalAlignment = xlCenter
Next intCounter
  
End With

End Sub
