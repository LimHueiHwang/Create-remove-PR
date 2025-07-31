Attribute VB_Name = "Module2"
Option Explicit

Sub Create_PR()

    ' Declare variables
    Dim Application1 As Object
    Dim Connection As Object
    Dim session As Object
    Dim ClosefixedPR As Worksheet
    Dim CombinePR As Worksheet
    Dim itotalRowsCombiPR As Long
    Dim i As Long
    Dim nPlant As String
    Dim DRemoveVendor As String
    Dim rowComPR As Long
    Dim nMaterial As Range
    Dim nQuantity As Range
    Dim MRPelement As String
    Dim recordDate As Date
    Dim recordQty As Long
    Dim Result As String
    Dim tableViewSize As Long
    Dim nR As Long
    Dim nScroll As Long
    Dim storednR As Long

    ' Set SAP application and worksheets
    Set Application1 = GetObject("SAPGUI").GetScriptingEngine
    Set ClosefixedPR = ThisWorkbook.Sheets("Close Fixed PR")
    Set CombinePR = ThisWorkbook.Sheets("Combine PR")

    ' Check if SAP is logged in
    If Application1.Children.Count > 0 Then
        Set Connection = Application1.Children(0)
    Else
        MsgBox "Please logon to SAP system before proceed!!!"
        Exit Sub
    End If

    ' Set session
    Set session = Connection.Children(0)

    ' Turn off alerts to prevent pop-ups during automation
    Application.DisplayAlerts = False

    ' Get total rows in Combine PR sheet
    itotalRowsCombiPR = CombinePR.Cells(Rows.Count, "A").End(xlUp).Row

    ' Clear previous results in Combine PR
    CombinePR.Range("C4:E" & itotalRowsCombiPR).Value = ""

    ' Loop through each row in Combine PR
    For i = 4 To itotalRowsCombiPR
        ' Fill in column E with extracted PO number from column C
        CombinePR.Range("E" & i).Formula = "=TEXTBEFORE(TEXTAFTER(C" & i & ",""purchase requisition ""), "" "")"
    Next i

    ' Get plant and vendor settings
    nPlant = Trim(CStr(CombinePR.Range("B2").Value))
    DRemoveVendor = Trim(CStr(CombinePR.Range("H1").Value))

    ' Validate plant code
    If nPlant = "" Then
        MsgBox "Please fill in the Plant Code"
        Exit Sub
    End If

    ' Loop through each row to process MRP data
    For rowComPR = 4 To itotalRowsCombiPR
        Set nMaterial = CombinePR.Range("A" & rowComPR)
        Set nQuantity = CombinePR.Range("B" & rowComPR)

        ' Refresh MRP by navigating in SAP
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd03"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/ctxtRM61X-MATNR").Text = nMaterial
        session.findById("wnd[0]/usr/ctxtRM61X-BERID").Text = nPlant
        session.findById("wnd[0]/usr/ctxtRM61X-WERKS").Text = nPlant
        session.findById("wnd[0]/tbar[0]/btn[0]").press

        ' Navigate to MRP data
        session.findById("wnd[0]").maximize
        session.findById("wnd[0]/tbar[0]/okcd").Text = "/nmd04"
        session.findById("wnd[0]/tbar[0]/btn[0]").press
        session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-MATNR").Text = nMaterial
        session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-BERID").Text = nPlant
        session.findById("wnd[0]/usr/tabsTAB300/tabpF01/ssubINCLUDE300:SAPMM61R:0301/ctxtRM61R-WERKS").Text = nPlant
        session.findById("wnd[0]/tbar[0]/btn[0]").press

        ' Ensure it's in GR view
        If Trim(session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/btnBUTTON_DAT00").Text) <> "AV" Then
            session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/btnBUTTON_DAT00").press
        End If

        ' Apply filter in SAP
        session.findById("wnd[0]/tbar[1]/btn[29]").press
        session.findById("wnd[0]/usr/subINCLUDE12XX:SAPMM61R:1200/cmbRM61R-FILBZ").Key = "SAP00001"

        ' Loop through MRP data rows
        tableViewSize = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ").VisibleRowCount
        For nR = 0 To tableViewSize - 1
            MRPelement = session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-DELB0[2," & nR & "]").Text

            ' If element matches criteria, process
            If Not (MRPelement Like "[_]*") Then
                recordDate = CDate(session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/ctxtMDEZ-DAT00[1," & nR & "]").Text)
                recordQty = CLng(session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-MNG01[8," & nR & "]").Text)

                ' Process for "SchLne" MRP element
                If MRPelement = "SchLne" Then
                    session.findById("wnd[0]/tbar[1]/btn[41]").press
                    session.findById("wnd[0]/usr/ctxtRM61X-BANER").Text = "3"
                    session.findById("wnd[0]/usr/ctxtRM61X-LIFKZ").Text = "1"
                    session.findById("wnd[0]/usr/ctxtRM61X-DISER").Text = "1"
                    session.findById("wnd[0]/usr/ctxtRM61X-PLMOD").Text = "1"
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]").sendVKey 0
                    session.findById("wnd[0]/tbar[1]/btn[6]").press
                    Exit For
                End If
            End If
        Next nR

        ' If matching MRP found, process PR creation
        If MRPelement = "PlOrd." Then
            session.findById("wnd[0]/usr/subINCLUDE1XX:SAPMM61R:0750/tblSAPMM61RTC_EZ/txtMDEZ-EXTRA[5," & nR & "]").SetFocus
            session.findById("wnd[0]").sendVKey 2
            session.findById("wnd[1]/tbar[0]/btn[27]").press
            session.findById("wnd[0]/usr/txtMDBA-MENGE").Text = Trim(nQuantity)
            session.findById("wnd[0]").sendVKey 0

            ' Handle quantity logic
            If nQuantity < recordQty Then
                session.findById("wnd[0]/usr/txtPLAF-GSMNG").Text = ""
            End If

            ' Remove vendor if necessary
            If DRemoveVendor = "X" Then
                session.findById("wnd[0]/usr/ctxtRM61P-EPSTP").Text = ""
                session.findById("wnd[0]/usr/ctxtMDBA-FLIEF").Text = ""
                session.findById("wnd[0]/usr/ctxtMDBA-KONNR").Text = ""
                session.findById("wnd[0]/usr/txtMDBA-KONPS").Text = ""
                session.findById("wnd[0]/usr/txtMDBA-RESWK").Text = ""
                session.findById("wnd[0]").sendVKey 0
            End If

            ' Finalize PR creation
            session.findById("wnd[0]/tbar[0]/btn[11]").press
            Result = session.findById("wnd[0]/sbar").Text
            CombinePR.Range("C" & rowComPR) = Result
        End If
    Next rowComPR

    ' Notify completion
    MsgBox "Done. Please check the result."

End Sub

