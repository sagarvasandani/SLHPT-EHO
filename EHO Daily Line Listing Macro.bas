Attribute VB_Name = "Module1"
Sub daily_report()

' turns off screen updating, makes it run faster
Application.ScreenUpdating = False
Application.DisplayAlerts = False

'define variables etc.
Dim GI As Variant
Dim Diagnosis As Variant
Dim fndList As Variant
Dim rplcList As Variant
Dim LR As Long
Dim r As Long
Dim wb As Workbook
Dim ws As Worksheet
Dim fip As String
Dim fop As String
Dim ld As String
Dim path As String
Dim dt As String





'contents
ld = "Lab data"
dt = Format(CStr(Now), "yymmdd")
fop = "S:\02  SURVEILLANCE\04 Reporting\Daily\Macro output"   'change "text to the desired FOlder Path
fip = "\" & dt & " " & ld 'this names the FIle path
path = fop & fip


'Array contents
PCR = Array("Typhoid Fever", "Salmonellosis", "Shigellosis", "E.coli infection", "E.coli infection, VTEC", "E.coli infection, VTEC O157", "Listeriosis", "Paratyphoid Fever")
Diagnosis = Array("Meningitis, other viral", "Pneumonia", "Meningitis, unspecified")
GI = Array("Aeromonas infection", "Campylobacteriosis", "Cholera", "Clostridium perfringens Infection", "Cryptosporidiosis", "E.coli infection, VTEC", "Enterococcal infection", "Giardiasis", "Listeriosis", "Mycobacterium infection, unspecified", "Paratyphoid Fever", "Salmonella infection, unspecified", "Salmonellosis", "Scombroid Toxin Poisoning", "Shigellosis (Bacillary Dysentery)", "Staphylococcus aureus intoxication", "Streptococcal Group C/G infection", "Typhoid Fever", "Unavailable", "Vibrio parahaemolyticus intoxication")
fndList = Array("Bexley London Boro", "Bromley London Boro", "Croydon London Boro", "Greenwich London Boro", "Kingston upon Thames London Boro", "Lambeth London Boro", "Lewisham London Boro", "Merton London Boro", "Richmond upon Thames London Boro", "Southwark London Boro", "Sutton London Boro", "Wandsworth London Boro")
rplcList = Array("Bexley", "Bromley", "Croydon", "Greenwich", "Kingston", "Lambeth", "Lewisham", "Merton", "Richmond", "Southwark", "Sutton", "Wandsworth")

'Delete columns that aren't needed
Range("b1:c1,g1,m1:n1,p1:v1,z1:ab1,ad1:af1,ah1:al1,an1:ar1,av1:ca1,cc1:cf1,ci1:ck1,cn1:cr1,cw1:cy1,da1:dd1,dh1:er1").EntireColumn.Delete

'Formats text
Range("A1") = "Case Identifier"
Range("A1:AF20000").Select
With Selection
    .Font.Name = "Calibri"
   .Font.Size = 8
   .Cells.VerticalAlignment = xlTop
  .Cells.HorizontalAlignment = xlLeft
  .Cells.WrapText = True
   .EntireColumn.AutoFit
End With

'this bit formats the header and autofilters
Range("A1:AF1").Select
With Selection
.Cells.Interior.ColorIndex = 15
.Font.Bold = True
End With
Range("A2").Select



'to filter out POSSIBLE "salmonellosis", "Shigellosis", "E.coli infection", "E.coli infection, VTEC",
' "E.coli infection, VTEC O157","Typhoid fever","listeriosis",
With ActiveSheet
.UsedRange.AutoFilter Field:=11, Criteria1:=PCR, Operator:=xlFilterValues
.UsedRange.AutoFilter Field:=12, Criteria1:="Possible", Operator:=xlFilterValues
.UsedRange.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
.AutoFilterMode = False
Range("A2").Select
End With

'this bit cleans out ILI
With ActiveSheet
.UsedRange.AutoFilter Field:=11, Criteria1:="Influenza or Flu-like illness", Operator:=xlFilterValues
.UsedRange.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
.AutoFilterMode = False
Range("A2").Select
End With

'this bit looks in column K and replaces "Unavailable" with "Food poisoning"
Columns("k").Select
Selection.Replace What:="Unavailable", Replacement:="Food Poisoning", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
Range("A2").Select

Range("s2", Range("s65536").End(xlUp)).Select
Selection.Replace What:="", Replacement:="MISSING", LookAt:=xlPart, _
        SearchOrder:=xlByRows, MatchCase:=False
Range("A2").Select

'Loop through each item in find/replace lists
  For x = LBound(fndList) To UBound(fndList)
    'Loop through each worksheet in ActiveWorkbook
      For Each sht In ActiveWorkbook.Worksheets
        sht.Cells.Replace What:=fndList(x), Replacement:=rplcList(x), _
          LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, _
          SearchFormat:=False, ReplaceFormat:=False
      Next sht
  
  Next x
        
'this bit deletes ALL NON GI infections
LR = Range("m" & Rows.Count).End(xlUp).Row
    For r = LR To 2 Step -1
        If IsError(Application.Match(Range("M" & r).Value, GI, False)) Then
            Rows(r).EntireRow.Delete
        End If
    Next r
    
Range("A2").Select

'this bit catches any discarded cases
With ActiveSheet
.UsedRange.AutoFilter Field:=12, Criteria1:="Discarded", Operator:=xlFilterValues
.UsedRange.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
.AutoFilterMode = False
Range("A2").Select
End With

'this bit catches any meningo or pneumonia cases (or anything else you want to put in diagnosis array)
With ActiveSheet
.UsedRange.AutoFilter Field:=11, Criteria1:=Diagnosis, Operator:=xlFilterValues
.UsedRange.Offset(1, 0).Rows.SpecialCells(xlCellTypeVisible).Delete (xlShiftUp)
.AutoFilterMode = False
Range("A2").Select
End With

' this bit looks at sheet and column creates new sheets based on unique names in column specified
ActiveSheet.Name = "Export"
Const sname As String = "Export" 'change to whatever starting sheet
Const s As String = "S" 'change to whatever is in LA column
Dim d As Object, a, cc&
Dim p&, i&, rws&, cls&
Set d = CreateObject("scripting.dictionary")
With Sheets(sname)
    rws = .Cells.Find("*", , , , xlByRows, xlPrevious).Row
    cls = .Cells.Find("*", , , , xlByColumns, xlPrevious).Column
    cc = .Columns(s).Column
End With
For Each sh In Worksheets
    d(sh.Name) = 1
Next sh


With Sheets.Add(after:=Sheets(sname))
Sheets(sname).Cells(1).Resize(rws, cls).Copy .Cells(1)
.Cells(1).Resize(rws, cls).Sort .Cells(cc), 2, Header:=xlYes
a = .Cells(cc).Resize(rws + 1, 1)
p = 2
For i = 2 To rws + 1
    If a(i, 1) <> a(p, 1) Then
        If d(a(p, 1)) <> 1 Then
            Sheets.Add.Name = a(p, 1)
            .Cells(1).Resize(, cls).Copy Cells(1)
            .Cells(p, 1).Resize(i - p, cls).Copy Cells(2, 1)
        End If
        p = i
    End If
Next i



Application.DisplayAlerts = False
    .Delete
Application.DisplayAlerts = True

End With
Sheets(sname).Activate

'this bit puts in some text
For Each ws In Worksheets
With ws
Call ActiveSheet.Hyperlinks.Add(.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Offset(1), https://extranet.phe.gov.uk/sites/SFT/SLHP/_layouts/15/start.aspx#/SitePages/Home.aspx, "", "UKHSA SLHPT/EHO SharePoint", "All up to date GI Disease protocols, Guidance and questionnaires can be accessed via UKHSA SLHPT/EHO SharePoint at: https://extranet.phe.gov.uk/sites/SFT/SLHP/_layouts/15/start.aspx#/SitePages/Home.aspx")
Call ActiveSheet.Hyperlinks.Add(.Cells(ActiveSheet.Rows.Count, 1).End(xlUp).Offset(2), mailto:jill.bedesha@phe.gov.uk?subject=Line listing query, "", "Email Jill Bedesha", "If you require access/have any issues please email jill.bedesha@phe.gov.uk")
End With
Next ws


' this bit looks at all the names in the workbook and saves them as a new file
For Each ws In ThisWorkbook.Worksheets
If Not ws.Name Like "Sheet*" Then
       Set wb = ws.Application.Workbooks.Add
      ws.Copy Before:=wb.Sheets(1)
    wb.SaveAs path & " " & ws.Name, Excel.XlFileFormat.xlOpenXMLWorkbook
   wb.Close
       Set wb = Nothing
End If
Next ws
    
Application.ScreenUpdating = True
Application.DisplayAlerts = True


  
End Sub




