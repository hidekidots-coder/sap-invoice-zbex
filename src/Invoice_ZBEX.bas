Option Explicit

' ============================================================
' SAP INVOICE ZBEX PROCESSOR
' ============================================================
'
' Purpose:
'   Automate the insertion of ZBEX output message in SAP invoices
'   (transaction VF02) based on a list in Excel.
'
' Technologies:
'   VBA
'   SAP GUI Scripting
'   Microsoft Excel
'
' ============================================================

Public Sub Invoice_ZBEX()

    On Error GoTo ErroHandler

    Dim SapGuiAuto As Object
    Dim SAPApp As Object
    Dim SAPCon As Object
    Dim session As Object
    
    Dim ws As Worksheet
    Dim Invoice As String
    Dim last4 As String
    
    Dim tbl As Object
    Dim campo As Object
    Dim i As Long
    Dim encontrouLinhaVazia As Boolean

    ' --------------------------------------------------------
    ' Configuration
    ' --------------------------------------------------------
    
    Set ws = ThisWorkbook.Sheets("invoice")     ' Remove the trailing space from the sheet name
    
    ' --------------------------------------------------------
    ' Connect to SAP GUI
    ' --------------------------------------------------------
    
    Set SapGuiAuto = GetObject("SAPGUI")
    Set SAPApp = SapGuiAuto.GetScriptingEngine
    Set SAPCon = SAPApp.Children(0)
    Set session = SAPCon.Children(0)

    ' --------------------------------------------------------
    ' Main Loop
    ' --------------------------------------------------------
    
    Do While Trim(ws.Range("A2").Value) <> ""
        
        Invoice = Trim(CStr(ws.Range("A2").Value))
        last4 = Right(Invoice, 4)
        
        On Error GoTo ErroInvoice
        
        ' Wait until SAP is free
        Do While session.Busy
            DoEvents
        Loop
        
        session.findById("wnd[0]").maximize
        session.StartTransaction "VF02"
        
        session.findById("wnd[0]/usr/ctxtVBRK-VBELN").Text = Invoice
        session.findById("wnd[0]").sendVKey 0          ' Enter
        
        ' Go to Messages tab
        session.findById("wnd[0]").sendVKey 20         ' Shift+F8 / Messages
        
        ' ====================================================
        ' Insert ZBEX in the first empty line
        ' ====================================================
        
        encontrouLinhaVazia = False
        Set tbl = session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3")
        
        For i = 0 To tbl.RowCount - 1
            
            Set campo = session.findById( _
                "wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," & i & "]")
            
            If Trim(campo.Text) = "" Then
                campo.SetFocus
                campo.Text = "ZBEX"
                encontrouLinhaVazia = True
                Exit For
            End If
            
        Next i
        
        If Not encontrouLinhaVazia Then
            MsgBox "Não há linha disponível para inserir ZBEX na invoice: " & Invoice, vbExclamation
            GoTo Continuar
        End If
        
        session.findById("wnd[0]").sendVKey 0          ' Enter
        
        ' ====================================================
        ' Configure the message
        ' ====================================================
        
        session.findById("wnd[0]").sendVKey 3          ' Back / Details
        
        ' Wait until SAP stabilizes
        Do While session.Busy
            DoEvents
        Loop
        
        session.findById("wnd[0]/usr/ctxtNAST-LDEST").Text = "SPOOLONLY"
        session.findById("wnd[0]/usr/txtNAST-ANZAL").Text = "1"
        session.findById("wnd[0]/usr/txtNAST-DSNAM").Text = "INV"
        session.findById("wnd[0]/usr/txtNAST-DSUF1").Text = last4
        
        ' Print Immediately
        session.findById("wnd[0]/usr/chkNAST-DIMME").Selected = True
        session.findById("wnd[0]/usr/chkNAST-DIMME").SetFocus
        
        session.findById("wnd[0]").sendVKey 3          ' Back
        
        ' ====================================================
        ' Ensure focus on ZBEX before F5
        ' ====================================================
        
        Set tbl = session.findById("wnd[0]/usr/tblSAPDV70ATC_NAST3")
        
        For i = 0 To tbl.RowCount - 1
            
            Set campo = session.findById( _
                "wnd[0]/usr/tblSAPDV70ATC_NAST3/ctxtDNAST-KSCHL[1," & i & "]")
            
            If Trim(campo.Text) = "ZBEX" Then
                campo.SetFocus
                campo.CaretPosition = 4
                Exit For
            End If
            
        Next i
        
        ' Wait until SAP is free
        Do While session.Busy
            DoEvents
        Loop
        
        session.findById("wnd[0]").sendVKey 5          ' F5
        
        session.findById("wnd[0]/usr/cmbNAST-VSZTP").Key = "4"
        
        session.findById("wnd[0]").sendVKey 3          ' Back
        session.findById("wnd[0]").sendVKey 11         ' Save
        
        session.findById("wnd[0]/tbar[0]/btn[3]").press   ' Back
        
        ' Remove processed invoice
        ws.Rows(2).Delete Shift:=xlUp
        
        On Error GoTo 0
        GoTo Continuar
        
ErroInvoice:
        MsgBox "Erro na invoice: " & Invoice & vbCrLf & Err.Description, vbCritical
        Exit Sub
        
Continuar:
    Loop
    
    MsgBox "Processo finalizado!", vbInformation

Saida:
    Set campo = Nothing
    Set tbl = Nothing
    Set session = Nothing
    Set SAPCon = Nothing
    Set SAPApp = Nothing
    Set SapGuiAuto = Nothing
    Set ws = Nothing
    Exit Sub

ErroHandler:
    MsgBox "Erro geral: " & Err.Description, vbCritical
    Resume Saida
    
End Sub
