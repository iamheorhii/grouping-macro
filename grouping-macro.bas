Attribute VB_Name = "Module1"
Option Explicit

' ============================================================
' CAN/LID Grouping Macro (General Template)
'
' What it does:
' - Reads all rows from the first worksheet (source)
' - Classifies rows as CAN or LID based on user-defined rules
' - Groups by Material + Plant, then creates CAN-centric groups
' - Outputs results into a newly created sheet: "CAN LID MAP"
'
' You MUST implement your own CAN/LID classification rules in
' IsCanRow / IsLidRow (see below).
' ============================================================

Public Sub Create_CANLID_GROUPS()
    Dim wsSrc As Worksheet, wsOut As Worksheet
    Dim lastRow As Long, lastCol As Long
    Dim data As Variant, headers As Variant
    Dim r As Long, c As Long
    Dim outRow As Long, groupNo As Long
    
    Dim mapMP As Object               ' key: material|plant -> dict("CAN") as Collection, dict("LID") as Collection
    Dim canRowByKey As Object         ' key: plant|<can_id> -> rowIndex
    Dim lidsByCanKey As Object        ' key: plant|<can_id> -> Collection(rowIndex)
    Dim seenSignatures As Object      ' signature -> True
    
    Dim material As String, plant As String, descr As String
    Dim mpKey As String, canKey As String
    
    Dim rowIdx As Variant, lidIdx As Variant
    Dim tmp As Variant
    
    ' --- Source/Output sheets
    Set wsSrc = ThisWorkbook.Worksheets(1)
    Set wsOut = RecreateSheet(ThisWorkbook, "CAN LID MAP")
    
    ' --- Read data into memory
    lastRow = wsSrc.Cells(wsSrc.Rows.Count, "B").End(xlUp).Row
    If lastRow < 2 Then
        MsgBox "No data rows found (expected at least 2 rows).", vbExclamation
        Exit Sub
    End If
    
    lastCol = wsSrc.Cells(1, wsSrc.Columns.Count).End(xlToLeft).Column
    headers = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(1, lastCol)).Value
    data = wsSrc.Range(wsSrc.Cells(1, 1), wsSrc.Cells(lastRow, lastCol)).Value
    
    ' --- Build Material|Plant buckets and keep CAN/LID row indexes
    Set mapMP = CreateObject("Scripting.Dictionary")
    
    For r = 2 To lastRow
        material = Trim$(CStr(data(r, 2)))          ' Column B: Material (customize if needed)
        plant = UCase$(Trim$(CStr(data(r, 5))))     ' Column E: Plant (customize if needed)
        descr = UCase$(Trim$(CStr(data(r, 10))))    ' Column J: Description (customize if needed)
        
        If Len(material) = 0 Or Len(plant) = 0 Then
            ' skip incomplete keys
        Else
            If IsCanRow(plant, descr, data, r) Or IsLidRow(plant, descr, data, r) Then
                mpKey = material & "|" & plant
                If Not mapMP.Exists(mpKey) Then
                    Dim bucket As Object
                    Set bucket = CreateObject("Scripting.Dictionary")
                    bucket.Add "CAN", New Collection
                    bucket.Add "LID", New Collection
                    mapMP.Add mpKey, bucket
                End If
                
                If IsCanRow(plant, descr, data, r) Then mapMP(mpKey)("CAN").Add r
                If IsLidRow(plant, descr, data, r) Then mapMP(mpKey)("LID").Add r
            End If
        End If
    Next r
    
    ' --- Build CAN key index and attach LIDs
    Set canRowByKey = CreateObject("Scripting.Dictionary")
    Set lidsByCanKey = CreateObject("Scripting.Dictionary")
    
    Dim mp As Variant
    For Each mp In mapMP.Keys
        plant = Split(CStr(mp), "|")(1)
        
        For Each rowIdx In mapMP(mp)("CAN")
            ' Column H is used here as the "ID" to group by (customize if needed)
            canKey = plant & "|" & Trim$(CStr(data(CLng(rowIdx), 8))) ' Column H
            
            If Not canRowByKey.Exists(canKey) Then
                canRowByKey.Add canKey, CLng(rowIdx)
                Set lidsByCanKey(canKey) = New Collection
            End If
            
            ' Current logic links all LIDs within the same Material|Plant bucket to each CAN.
            ' If you need stricter matching, implement it inside ShouldAttachLidToCan().
            For Each lidIdx In mapMP(mp)("LID")
                If ShouldAttachLidToCan(data, CLng(canRowByKey(canKey)), CLng(lidIdx)) Then
                    If Not CollectionContainsRow(lidsByCanKey(canKey), CLng(lidIdx)) Then
                        lidsByCanKey(canKey).Add CLng(lidIdx)
                    End If
                End If
            Next lidIdx
        Next rowIdx
    Next mp
    
    ' --- Output header
    wsOut.Cells(1, 1).Value = "Group No."
    wsOut.Cells(1, 1).Font.Bold = True
    
    For c = 1 To lastCol
        wsOut.Cells(1, c + 1).Value = headers(1, c)
        wsOut.Cells(1, c + 1).Font.Bold = True
        wsOut.Cells(1, c + 1).Interior.Color = RGB(206, 206, 206)
    Next c
    
    ' --- Output groups with signature de-duplication
    Set seenSignatures = CreateObject("Scripting.Dictionary")
    outRow = 2
    groupNo = 1
    
    Dim k As Variant
    For Each k In canRowByKey.Keys
        If lidsByCanKey(k).Count = 0 Then
            ' no LIDs linked -> skip group
        Else
            Dim signature As String
            signature = BuildGroupSignature(data, canRowByKey(k), lidsByCanKey(k))
            
            If Not seenSignatures.Exists(signature) Then
                seenSignatures.Add signature, True
                
                ' CAN row (bold)
                wsOut.Cells(outRow, 1).Value = groupNo
                For c = 1 To lastCol
                    wsOut.Cells(outRow, c + 1).Value = data(CLng(canRowByKey(k)), c)
                Next c
                wsOut.Rows(outRow).Font.Bold = True
                outRow = outRow + 1
                
                ' LID rows (sorted by row index for stable output)
                Dim lidRows() As Long
                lidRows = CollectionToLongArray(lidsByCanKey(k))
                If UBound(lidRows) >= LBound(lidRows) Then
                    QuickSortLong lidRows, LBound(lidRows), UBound(lidRows)
                End If
                
                Dim ii As Long
                For ii = LBound(lidRows) To UBound(lidRows)
                    wsOut.Cells(outRow, 1).Value = groupNo
                    For c = 1 To lastCol
                        wsOut.Cells(outRow, c + 1).Value = data(lidRows(ii), c)
                    Next c
                    outRow = outRow + 1
                Next ii
                
                outRow = outRow + 1
                groupNo = groupNo + 1
            End If
        End If
    Next k
    
    wsOut.Columns.AutoFit
    MsgBox "CAN/LID grouping complete. Groups created: " & (groupNo - 1), vbInformation
End Sub

' -----------------------
' USER CONFIG SECTION
' -----------------------

' Decide whether row r should be treated as CAN.
' Implement your own logic here. Keep it deterministic and fast.
Private Function IsCanRow(ByVal plant As String, ByVal descr As String, ByRef data As Variant, ByVal r As Long) As Boolean
    IsCanRow = False
    
    ' EXAMPLES (REMOVE/EDIT):
    ' If plant = "XY" And descr Like "XX*" Then IsCanRow = True
    ' If plant = "XY" And descr Like "XX*" Then IsCanRow = True
    
    ' GENERAL PLACEHOLDER:
    ' Insert your rules here, e.g.
    ' If InStr(1, descr, "CAN", vbTextCompare) > 0 Then IsCanRow = True
End Function

' Decide whether row r should be treated as LID.
Private Function IsLidRow(ByVal plant As String, ByVal descr As String, ByRef data As Variant, ByVal r As Long) As Boolean
    IsLidRow = False
    
    ' EXAMPLES (REMOVE/EDIT):
    ' If plant = "XY" And descr Like "XX*" Then IsLidRow = True
    ' If plant = "YY" And descr Like "XX*" Then IsLidRow = True
    
    ' GENERAL PLACEHOLDER:
    ' Insert your rules here, e.g.
    ' If InStr(1, descr, "LID", vbTextCompare) > 0 Then IsLidRow = True
End Function

' Optional: control whether a LID should be attached to a given CAN inside the same Material|Plant bucket.
' Current default keeps the original behavior (attach all LIDs in the bucket).
Private Function ShouldAttachLidToCan(ByRef data As Variant, ByVal canRow As Long, ByVal lidRow As Long) As Boolean
    ShouldAttachLidToCan = True
    
    ' Example stricter rule (customize):
    ' Attach only if Column H matches:
    ' ShouldAttachLidToCan = (Trim$(CStr(data(canRow, 8))) = Trim$(CStr(data(lidRow, 8))))
End Function

' -----------------------
' OUTPUT HELPERS
' -----------------------

Private Function BuildGroupSignature(ByRef data As Variant, ByVal canRow As Long, ByVal lidRows As Collection) As String
    ' Signature based on Column H values (customize if Column H is not your identifier).
    Dim parts() As String
    Dim i As Long, n As Long
    n = lidRows.Count + 1
    ReDim parts(1 To n)
    
    parts(1) = Trim$(CStr(data(canRow, 8)))
    
    Dim v As Variant
    i = 2
    For Each v In lidRows
        parts(i) = Trim$(CStr(data(CLng(v), 8)))
        i = i + 1
    Next v
    
    QuickSortString parts, LBound(parts), UBound(parts)
    BuildGroupSignature = Join(parts, "|")
End Function

Private Function CollectionContainsRow(ByVal col As Collection, ByVal rowIndex As Long) As Boolean
    Dim v As Variant
    For Each v In col
        If CLng(v) = rowIndex Then
            CollectionContainsRow = True
            Exit Function
        End If
    Next v
    CollectionContainsRow = False
End Function

Private Function CollectionToLongArray(ByVal col As Collection) As Long()
    Dim arr() As Long
    Dim i As Long
    
    If col.Count = 0 Then
        ReDim arr(0 To -1)
        CollectionToLongArray = arr
        Exit Function
    End If
    
    ReDim arr(1 To col.Count)
    For i = 1 To col.Count
        arr(i) = CLng(col(i))
    Next i
    
    CollectionToLongArray = arr
End Function

Private Function RecreateSheet(ByVal wb As Workbook, ByVal sheetName As String) As Worksheet
    Application.DisplayAlerts = False
    On Error Resume Next
    wb.Worksheets(sheetName).Delete
    On Error GoTo 0
    Application.DisplayAlerts = True
    
    Set RecreateSheet = wb.Worksheets.Add(After:=wb.Worksheets(wb.Worksheets.Count))
    RecreateSheet.Name = sheetName
End Function

' -----------------------
' SORT HELPERS
' -----------------------

Private Sub QuickSortLong(arr() As Long, ByVal first As Long, ByVal last As Long)
    Dim lo As Long, hi As Long, pivot As Long, tmp As Long
    lo = first: hi = last
    pivot = arr((first + last) \ 2)
    
    Do While lo <= hi
        Do While arr(lo) < pivot: lo = lo + 1: Loop
        Do While arr(hi) > pivot: hi = hi - 1: Loop
        
        If lo <= hi Then
            tmp = arr(lo): arr(lo) = arr(hi): arr(hi) = tmp
            lo = lo + 1: hi = hi - 1
        End If
    Loop
    
    If first < hi Then QuickSortLong arr, first, hi
    If lo < last Then QuickSortLong arr, lo, last
End Sub

Private Sub QuickSortString(arr() As String, ByVal first As Long, ByVal last As Long)
    Dim lo As Long, hi As Long, pivot As String, tmp As String
    lo = first: hi = last
    pivot = arr((first + last) \ 2)
    
    Do While lo <= hi
        Do While arr(lo) < pivot: lo = lo + 1: Loop
        Do While arr(hi) > pivot: hi = hi - 1: Loop
        
        If lo <= hi Then
            tmp = arr(lo): arr(lo) = arr(hi): arr(hi) = tmp
            lo = lo + 1: hi = hi - 1
        End If
    Loop
    
    If first < hi Then QuickSortString arr, first, hi
    If lo < last Then QuickSortString arr, lo, last
End Sub

