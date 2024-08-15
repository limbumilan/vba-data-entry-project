Attribute VB_Name = "Module1"

        
        
        Sub data_table(name, table)
        
        
        Dim wb As Workbook
        Dim ws As Worksheet
        Dim tbl As ListObject
        ThisWorkbook.Activate
        

         Set wb = ActiveWorkbook
         
         
         Set ws = wb.Worksheets(name)
         
         Set tbl = ws.ListObjects(table)
          Set newRow = tbl.ListRows.Add
          With ws
          .Select
          
       
           With newRow.Range
       
        .Cells(1, 1).value = UserForm1.nametxt.value
        .Cells(1, 2).value = UserForm1.classtxt.value
        
        ' Add more cells if needed
  

    ' Clear the text boxes
        UserForm1.nametxt.value = ""
        UserForm1.classtxt.value = ""
       MsgBox "Data has been added to the table!", vbInformation
         End With
        End With

  

        End Sub


