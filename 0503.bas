Attribute VB_Name = "Module1"
Option Explicit
Sub ³W«ß¦X¨Ö()
Dim shtsIdx As Integer
Dim k As Long
For shtsIdx = 1 To Sheets.Count
Sheets(shtsIdx).Activate
For k = 2 To 11 Step 3
    Dim rangeStr As String
    rangeStr = "A" & k & ":A" & k + 2
    
    Range(rangeStr).Merge
Next
Next
End Sub
