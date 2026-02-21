Attribute VB_Name = "M_Data_Comps_SupplierPick"
Option Explicit

'===============================================================================
' Module: M_Data_Comps_SupplierPick
'
' Purpose:
'   Robust, user-recoverable supplier picker by SupplierName (users do NOT need SupplierID).
'   - Keeps prompting until a valid supplier is selected or user explicitly cancels.
'   - Matches both raw substring and normalized substring (handles &, punctuation, weird spaces).
'   - Self-diagnosing: on 0 matches, shows sample SupplierName values + normalized forms,
'     and reports which required headers were found/missing.
'
' Inputs (Tabs/Tables/Headers):
'   - Sheet: Suppliers
'   - Table: TBL_SUPPLIERS
'   - Required headers:
'       SupplierID
'       SupplierName
'       SupplierDefaultLT
'
' Outputs / Side effects:
'   - Returns supplierId, supplierName, supplierDefaultLT via ByRef args
'   - Displays MsgBox diagnostics when no matches / config issues
'
' Preconditions / Postconditions:
'   - TBL_SUPPLIERS exists and has rows
'
' Errors & Guards:
'   - Fails fast with clear message if required headers are missing
'   - Does not silently fail; always re-prompts or allows explicit cancel
'
' Version: v1.0.0
' Author: Keith + GPT
' Date: 2025-12-29
'===============================================================================

'==========================
' PUBLIC TEST HARNESS
'==========================
Public Sub Test_SupplierSearch()
    Const SH_SUPPLIERS As String = "Suppliers"
    Const LO_SUPPLIERS As String = "TBL_SUPPLIERS"

    Dim loSupp As ListObject
    Dim supplierId As String, supplierName As String, supplierDefaultLT As Variant
    Dim ok As Boolean

    Set loSupp = ThisWorkbook.Worksheets(SH_SUPPLIERS).ListObjects(LO_SUPPLIERS)

    ok = SupplierPick_ByName(loSupp, supplierId, supplierName, supplierDefaultLT)

    If ok Then
        MsgBox "Selected:" & vbCrLf & _
               "SupplierName: " & supplierName & vbCrLf & _
               "SupplierID: " & supplierId & vbCrLf & _
               "SupplierDefaultLT: " & CStr(supplierDefaultLT), vbOKOnly, "Test Supplier Search"
    Else
        MsgBox "Cancelled (no selection).", vbOKOnly, "Test Supplier Search"
    End If
End Sub

'==========================
' MAIN PICKER (CALL THIS)
'==========================
Public Function SupplierPick_ByName(ByVal loSupp As ListObject, ByRef supplierId As String, ByRef supplierName As String, ByRef supplierDefaultLT As Variant) As Boolean
    Const title As String = "Pick Supplier"
    Const MAX_SHOW As Long = 25
    Const SAMPLE_SHOW As Long = 12

    Dim idxId As Long, idxName As Long, idxLT As Long
    Dim arrId As Variant, arrName As Variant, arrLT As Variant

    Dim term As String, termNorm As String
    Dim i As Long, hitCount As Long
    Dim hitsIdx() As Long

    Dim menu As String
    Dim choiceText As String
    Dim choiceN As Long

    SupplierPick_ByName = False
    supplierId = vbNullString
    supplierName = vbNullString
    supplierDefaultLT = vbNullString

    If loSupp Is Nothing Then
        MsgBox "Suppliers table reference is Nothing.", vbOKOnly, title
        Exit Function
    End If

    If loSupp.DataBodyRange Is Nothing Then
        MsgBox "Suppliers table '" & loSupp.Name & "' has no rows.", vbOKOnly, title
        Exit Function
    End If

    ' Resolve required columns (strict header match, case-insensitive)
    idxId = GetColIndex(loSupp, "SupplierID")
    idxName = GetColIndex(loSupp, "SupplierName")
    idxLT = GetColIndex(loSupp, "SupplierDefaultLT")

    If idxId = 0 Or idxName = 0 Or idxLT = 0 Then
        MsgBox "Supplier picker cannot run because required headers were not found." & vbCrLf & vbCrLf & _
               "Expected headers in table '" & loSupp.Name & "':" & vbCrLf & _
               "  SupplierID, SupplierName, SupplierDefaultLT" & vbCrLf & vbCrLf & _
               "Found indices:" & vbCrLf & _
               "  SupplierID=" & CStr(idxId) & vbCrLf & _
               "  SupplierName=" & CStr(idxName) & vbCrLf & _
               "  SupplierDefaultLT=" & CStr(idxLT) & vbCrLf & vbCrLf & _
               "Action: rename the table headers to match exactly, or update the code to your actual header text.", _
               vbOKOnly, title
        Exit Function
    End If

    ' Pull arrays
    arrId = loSupp.ListColumns(idxId).DataBodyRange.value
    arrName = loSupp.ListColumns(idxName).DataBodyRange.value
    arrLT = loSupp.ListColumns(idxLT).DataBodyRange.value

RetrySearch:
    term = InputBox( _
        "Supplier search:" & vbCrLf & _
        "- Type part of the supplier name (e.g., B&B or Thread)." & vbCrLf & _
        "- Leave blank and click OK to cancel.", _
        title)

    term = Trim$(term)
    If Len(term) = 0 Then
        SupplierPick_ByName = False
        Exit Function
    End If

    termNorm = NormalizeForMatch(term)

    hitCount = 0
    ReDim hitsIdx(1 To 1)

    For i = 1 To UBound(arrName, 1)
        Dim candidate As String, candNorm As String
        candidate = CStr(arrName(i, 1))
        candNorm = NormalizeForMatch(candidate)

        If InStr(1, candidate, term, vbTextCompare) > 0 _
           Or (Len(termNorm) > 0 And InStr(1, candNorm, termNorm, vbTextCompare) > 0) Then

            hitCount = hitCount + 1
            If hitCount = 1 Then
                hitsIdx(1) = i
            Else
                ReDim Preserve hitsIdx(1 To hitCount)
                hitsIdx(hitCount) = i
            End If
        End If
    Next i

    If hitCount = 0 Then
        Dim msg As String
        Dim nShow As Long

        nShow = WorksheetFunction.Min(SAMPLE_SHOW, UBound(arrName, 1))

        msg = "No suppliers matched: '" & term & "'" & vbCrLf & vbCrLf & _
              "Normalized search: " & termNorm & vbCrLf & vbCrLf & _
              "Sample SupplierName values (first " & CStr(nShow) & "):" & vbCrLf

        For i = 1 To nShow
            msg = msg & "  - " & CStr(arrName(i, 1)) & "    (Norm: " & NormalizeForMatch(CStr(arrName(i, 1))) & ")" & vbCrLf
        Next i

        msg = msg & vbCrLf & _
              "Try a shorter fragment (e.g., THREAD), or omit punctuation." & vbCrLf & _
              "If you expected a match, verify the supplier name exists in Suppliers.TBL_SUPPLIERS[SupplierName]."

        MsgBox msg, vbOKOnly, title
        GoTo RetrySearch
    End If

    If hitCount = 1 Then
        i = hitsIdx(1)
        supplierId = CStr(arrId(i, 1))
        supplierName = CStr(arrName(i, 1))
        supplierDefaultLT = arrLT(i, 1)
        SupplierPick_ByName = True
        Exit Function
    End If

RetryPick:
    menu = "Multiple suppliers matched '" & term & "'. Choose a number:" & vbCrLf & vbCrLf

    For i = 1 To hitCount
        menu = menu & CStr(i) & ") " & CStr(arrName(hitsIdx(i), 1)) & "  [" & CStr(arrId(hitsIdx(i), 1)) & "]" & vbCrLf
        If i >= MAX_SHOW Then
            menu = menu & vbCrLf & "(Showing first " & CStr(MAX_SHOW) & ". Refine your search to narrow results.)"
            Exit For
        End If
    Next i

    choiceText = InputBox(menu & vbCrLf & "Enter a number (or leave blank to re-search).", title)
    choiceText = Trim$(choiceText)

    If Len(choiceText) = 0 Then GoTo RetrySearch

    If Not IsNumeric(choiceText) Then
        MsgBox "Please enter a valid number from the list.", vbOKOnly, title
        GoTo RetryPick
    End If

    choiceN = CLng(choiceText)
    If choiceN < 1 Or choiceN > hitCount Then
        MsgBox "Choice out of range. Pick a number shown in the list.", vbOKOnly, title
        GoTo RetryPick
    End If

    i = hitsIdx(choiceN)
    supplierId = CStr(arrId(i, 1))
    supplierName = CStr(arrName(i, 1))
    supplierDefaultLT = arrLT(i, 1)

    SupplierPick_ByName = True
End Function

'==========================
' NORMALIZATION HELPERS
'==========================
Private Function NormalizeForMatch(ByVal s As String) As String
    ' Uppercase, replace & with AND, remove punctuation to spaces, collapse spaces.
    Dim t As String
    Dim i As Long
    Dim ch As String
    Dim out As String

    t = UCase$(Trim$(Replace(CStr(s), ChrW(160), " "))) ' also remove NBSP (common copy/paste issue)
    t = Replace(t, "&", " AND ")

    out = vbNullString
    For i = 1 To Len(t)
        ch = Mid$(t, i, 1)
        If (ch >= "A" And ch <= "Z") Or (ch >= "0" And ch <= "9") Then
            out = out & ch
        Else
            out = out & " "
        End If
    Next i

    NormalizeForMatch = CollapseSpaces(out)
End Function

Private Function CollapseSpaces(ByVal s As String) As String
    Dim t As String
    t = Trim$(s)
    Do While InStr(1, t, "  ", vbBinaryCompare) > 0
        t = Replace(t, "  ", " ")
    Loop
    CollapseSpaces = t
End Function

'==========================
' LISTOBJECT HELPERS
'==========================
Private Function GetColIndex(ByVal lo As ListObject, ByVal header As String) As Long
    Dim lc As ListColumn
    For Each lc In lo.ListColumns
        If StrComp(lc.Name, header, vbTextCompare) = 0 Then
            GetColIndex = lc.Index
            Exit Function
        End If
    Next lc
    GetColIndex = 0
End Function


