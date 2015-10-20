Attribute VB_Name = "CombineSpreadsheets"

Sub CombineSpreadsheets()
    Dim Path As String
    'Make sure the Path has a trailing "\"
    Path = "C:\share\Output\"
    FileName = Dir(Path & "*.csv")
    'FileName = Dir(Path & "*.xls")
    'FileName = Dir(Path & "*.xlsx")
    
    Do While FileName <> ""
      Workbooks.Open FileName:=Path & FileName, ReadOnly:=True
         For Each Sheet In ActiveWorkbook.Sheets
         Sheet.Copy After:=ThisWorkbook.Sheets(1)
      Next Sheet
         Workbooks(FileName).Close
         FileName = Dir()
      Loop
      
End Sub
