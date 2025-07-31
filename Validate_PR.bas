Attribute VB_Name = "Module3"
Option Explicit

Sub validate_PR()
    ' Setup and Initialization
    Dim Application1 As Object, Connection As Object, session As Object
    Dim ClosefixedPR As Worksheet, CombinePR As Worksheet, Plan_Order As Worksheet
    Dim itotalRowsCombiPR As Long, lastRow As Long
    Dim clipboardData As String
    Dim wbSAP As Workbook, wbTarget As Workbook, wbSource As Workbook, wb As Workbook
    Dim wsSource As Worksheet, wsTarget As Worksheet
    Dim found As Boolean, isOpen As Boolean
    Dim nPlant As String, wbNamePattern As String
    Dim nPart As Variant, startTime As Double
    Dim i As Long, rowIndex As Integer
    Dim fieldID As String, obj As Object
    
    ' Initialize SAP GUI Scripting
    Set Application1 = GetObject("SAPGUI").GetScriptingEngine
    
    ' Reference worksheets
    Set ClosefixedPR = ThisWorkbook.Sheets("Close Fixed PR")
    Set CombinePR = ThisWorkbook.Sheets("Combine PR")
    Set Plan_Order = ThisWorkbook.Sheets("Plan Order")

    ' Check SAP Logon status
    If Application1.Children.Count = 0 Then
        MsgBox "Please log on to the SAP system before proceeding!", vbCritical
        Exit Sub
    End If

    Set Connection = Application1.Children(0)
    Set session = Connection.Children(0)

    Application.DisplayAlerts = False

    ' Extract Data and Setup Sheet Formulas
    itotalRowsCombiPR = CombinePR.Cells(Rows.Count, "A").End(xlUp).Row
    lastRow = CombinePR.Cells(CombinePR.Rows.Count, "G").End(xlUp).Row
    CombinePR.Range("C4:G" & lastRow).ClearContents

    For i = 4 To itotalRowsCombiPR
        CombinePR.Range("E" & i).Formula = "=TEXTBEFORE(TEXTAFTER(C" & i & ",""purchase requisition ""), "" "")"
        CombinePR.Range("F" & i).Formula = "=IFERROR(XLOOKUP(A" & i & ",Summary!A:A,Summary!B:B),0)"
        CombinePR.Range("G" & i).Formula = "=IF(F" & i & ">=B" & i & ",""Ok"",IF(F" & i & "=0,""Check if SA part"",""Not enough Plan Order""))"
    Next i

    ' Validate Plant Code
    nPlant = Trim(CStr(CombinePR.Range("B2").Value))
    If nPlant = "" Then
        MsgBox "Please fill in the Plant Code", vbExclamation
        Exit Sub
    End If

    ' Clear Plan Order Sheet
    Plan_Order.Cells.Clear

    ' Run SAP Transaction SQ00
    With session
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nsq00"
        .findById("wnd[0]").sendVKey 0
        .findById("wnd[0]/tbar[1]/btn[19]").press
        .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").currentCellRow = 8
        .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").selectedRows = "8"
        .findById("wnd[1]/usr/cntlGRID1/shellcont/shell").doubleClickCurrentCell
        .findById("wnd[0]/usr/ctxtRS38R-QNUM").Text = "PU-088MRP_EFFI"
        .findById("wnd[0]/tbar[1]/btn[8]").press
        .findById("wnd[0]/usr/ctxtLANGUAGE-LOW").Text = "EN"
        .findById("wnd[0]/usr/ctxtPLANT-LOW").Text = nPlant
        .findById("wnd[0]/usr/btn%_SP$00019_%_APP_%-VALU_PUSH").press
        .findById("wnd[1]/tbar[0]/btn[16]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/btn%_SP$00003_%_APP_%-VALU_PUSH").press
        .findById("wnd[1]/tbar[0]/btn[16]").press
        .findById("wnd[1]/tbar[0]/btn[8]").press
        .findById("wnd[0]/usr/rad%EXCEL").Select
        .findById("wnd[0]/usr/ctxtSP$00003-LOW").Text = "PA"
        .findById("wnd[0]/usr/ctxtMATERIAL-LOW").Text = "1"
        .findById("wnd[0]/usr/btn%_MATERIAL_%_APP_%-VALU_PUSH").press
    End With

    ' Input Materials into SAP Selection
    If itotalRowsCombiPR = 4 Then
        ReDim nPart(1 To 1, 1 To 1)
        nPart(1, 1) = CombinePR.Range("A4").Value
    Else
        nPart = CombinePR.Range("A4:A" & itotalRowsCombiPR).Value
    End If

    For i = 1 To UBound(nPart, 1)
        If i <= 7 Then
            rowIndex = (i - 1) Mod 7
        Else
            rowIndex = ((i - 1) Mod 7) + 1
        End If

        fieldID = "wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" & _
                  "tblSAPLALDBSINGLE/ctxtRSCSEL_255-SLOW_I[1," & rowIndex & "]"

        If rowIndex = 1 And i > 7 Then
            session.findById("wnd[1]/usr/tabsTAB_STRIP/tabpSIVA/ssubSCREEN_HEADER:SAPLALDB:3010/" & _
                             "tblSAPLALDBSINGLE").VerticalScrollbar.Position = i - 1
        End If

        On Error Resume Next
        Set obj = session.findById(fieldID)
        If Err.Number <> 0 Then
            Debug.Print "Error: Control not found - " & fieldID
            Exit For
        Else
            obj.Text = nPart(i, 1)
        End If
        On Error GoTo 0
    Next i

    ' Execute SAP Program and Download Excel
    session.findById("wnd[1]/tbar[0]/btn[8]").press
    session.findById("wnd[0]/tbar[1]/btn[8]").press
    session.findById("wnd[1]/usr/subSUBSCREEN_STEPLOOP:SAPLSPO5:0150/sub:SAPLSPO5:0150/radSPOPLI-SELFLAG[0,0]").Select
    session.findById("wnd[1]/tbar[0]/btn[0]").press
    session.findById("wnd[1]/tbar[0]/btn[0]").press

    ' Wait for SAP to Export Excel
    wbNamePattern = "Worksheet in Basis*"
    startTime = Timer: found = False

    Do While Not found And Timer - startTime < 60
        DoEvents
        For Each wb In Application.Workbooks
            If wb.Name Like wbNamePattern Then
                Set wbSource = wb
                found = True
                Exit For
            End If
        Next wb
    Loop

    If Not found Then
        MsgBox "Workbook not opened within 60 seconds. Exiting.", vbExclamation
        Exit Sub
    End If

    ' Copy Data to Plan Order Sheet
    Set wsSource = wbSource.Sheets(1)
    Set wsTarget = Plan_Order
    wsTarget.Cells.Clear
    wsSource.UsedRange.Copy Destination:=wsTarget.Range("A1")

    ' Format column C as numbers
    With wsTarget.Columns("C")
        .NumberFormat = "0.00"
        .Value = .Value
    End With

    MsgBox "Validate_PR successfully!", vbInformation, "Done"
End Sub


