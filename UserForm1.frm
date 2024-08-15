VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} UserForm1 
   Caption         =   "UserForm1"
   ClientHeight    =   6810
   ClientLeft      =   120
   ClientTop       =   465
   ClientWidth     =   16710
   OleObjectBlob   =   "UserForm1.frx":0000
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "UserForm1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False


Private Sub classtxt_AfterUpdate()
If Not IsNumeric(classtxt.value) Then
        MsgBox "only numbers allowed in the class field:"
        Cancel = True
        End If
End Sub

Private Sub ComboBox1_Change()
Dim varNRows As Integer
Select Case ComboBox1.value
Case "classone"
    With ListBox1
    varNRows = Sheets("Sheet1").Cells(Rows.Count, "A").End(xlUp).Row
    
    
    .RowSource = "sheet1!A6:B6" & varNRows
    .ColumnHeads = False
    .ColumnCount = 2
    .ColumnWidths = "80;80"
    End With
  
 Case "classtwo"
 With ListBox1
 varNRows = Sheets("data").Cells(Rows.Count, "A").End(xlUp).Row
    .RowSource = "data!A5:B5" & varNRows
    End With
  Case "classthree"
  With ListBox1
    varNRows = Sheets("Sheet3").Cells(Rows.Count, "A").End(xlUp).Row
    .RowSource = "sheet3!A1:B1" & varNRows
    .ColumnHeads = True
    .ColumnCount = 2
    .ColumnWidths = "10;10"
    End With
 

Case Else

    
End Select

End Sub




Private Sub CommandButton1_Click()
   Dim value As Integer
   
   If Len(nametxt.value) = 0 Then
   
   MsgBox "enter name and class information ", vbInformation
   
   
   Else
   
   
   


    ' Insert data from text boxes into the new row

 Select Case classtxt.value
     Case 1
       Call data_table("Sheet1", "totaltable")
     Case 2
       Call data_table("data", "class2table")
    
        
      Case Else
         nametxt.value = ""
         classtxt.value = ""
    
         MsgBox "Enter a valid value:", vbInformation
 End Select
    ' Optionally, you can display a message box confirming submission
    
    End If
    
   
   
    



End Sub

Private Sub Workbook_Open()

    ' Activates "Summary" worksheet when workbook is opened
    ThisWorkbook.Worksheets("Sheet1").Activate
    ThisWorkbook.Worksheets("Sheet2").Activate

End Sub


Private Sub MultiPage1_Change()

End Sub




Private Sub CommandButton2_Click()
UserForm1.Hide
Select Case ComboBox1.value
Case "classone"
    Sheet1.PrintPreview
    UserForm1.Show
    
 Case "classtwo"
 
  Case "classthree"
  Case Else
 End Select
 
End Sub




Private Sub ListBox1_Change()
With ListBox1


End With

End Sub



Private Sub UserForm_Initialize()
With ComboBox1
.AddItem "classone"
.AddItem "classtwo"
.AddItem "classthree"
End With
End Sub

