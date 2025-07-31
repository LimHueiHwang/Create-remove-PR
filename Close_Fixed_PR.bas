Attribute VB_Name = "Module1"
Option Explicit

Sub Close_Fixed_PR()

    ' Declare SAP GUI objects
    Dim SapGuiApp As Object
    Dim SapConnection As Object
    Dim SapSession As Object
    
    ' Declare Excel worksheet and variables
    Dim wsCloseFixedPR As Worksheet
    Dim lastRow As Long
    Dim currentPRCell As Range
    Dim rowIndex As Long
    Dim sapMessage As String

    ' Initialize SAP GUI
    On Error Resume Next
    Set SapGuiApp = GetObject("SAPGUI").GetScriptingEngine
    On Error GoTo 0
    
    If SapGuiApp Is Nothing Then
        MsgBox "Please log in to SAP before running this script.", vbCritical
        Exit Sub
    End If

    ' Connect to first available SAP connection/session
    Set SapConnection = SapGuiApp.Children(0)
    Set SapSession = SapConnection.Children(0)

    ' Reference worksheet
    Set wsCloseFixedPR = ThisWorkbook.Sheets("Close Fixed PR")

    ' Turn off alerts
    Application.DisplayAlerts = False

    ' Get the last row with data in column A
    lastRow = wsCloseFixedPR.Cells(wsCloseFixedPR.Rows.Count, "A").End(xlUp).Row

    ' Clear old result columns (B and C)
    wsCloseFixedPR.Range("B4:C" & lastRow).ClearContents

    ' Start ME52N transaction
    With SapSession
        .findById("wnd[0]").maximize
        .findById("wnd[0]/tbar[0]/okcd").Text = "/nme52n"
        .findById("wnd[0]").sendVKey 0
    End With

    ' Loop through each PR in column A starting from row 4
    For rowIndex = 4 To lastRow
        Set currentPRCell = wsCloseFixedPR.Range("A" & rowIndex)

        ' Click on "Other Purchase Requisition"
        SapSession.findById("wnd[0]/tbar[1]/btn[17]").press

        ' Enter PR number
        With SapSession.findById("wnd[1]/usr/subSUB0:SAPLMEGUI:0003/ctxtMEPO_SELECT-BANFN")
            .Text = currentPRCell.Value
            .caretPosition = Len(.Text)
        End With
        SapSession.findById("wnd[1]").sendVKey 0

        ' Check for PR existence or messages
        sapMessage = SapSession.findById("wnd[0]/sbar/pane[0]").Text

        If Not sapMessage Like "*does not exist*" Then
            ' Untick FIXKZ, Tick EBAKZ
            With SapSession
                .findById("wnd[0]/usr/.../chkMEREQ3321-FIXKZ").Selected = False
                .findById("wnd[0]/usr/.../chkMEREQ3321-EBAKZ").Selected = True
                .findById("wnd[0]/usr/.../chkMEREQ3321-FIXKZ").SetFocus

                ' Click Save
                .findById("wnd[0]/tbar[0]/btn[11]").press
            End With
        End If

        ' Capture message after save
        sapMessage = SapSession.findById("wnd[0]/sbar/pane[0]").Text

        ' Handle "No changes made" case
        If sapMessage = "" Then
            SapSession.findById("wnd[1]/tbar[0]/btn[0]").press
            wsCloseFixedPR.Range("B" & rowIndex).Value = "No changes made"
        Else
            wsCloseFixedPR.Range("B" & rowIndex).Value = sapMessage
        End If

    Next rowIndex

    MsgBox "Process completed. Please check column B for results.", vbInformation

End Sub

