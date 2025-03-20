Attribute VB_Name = "ExcelModule"
Option Explicit

' Constants for string patterns
Private Const INCIDENTE_PATTERN As String = "INC[0-9]+|SOL[0-9]+|RFC[0-9]+"
Private Const APESA_FOLIO_PATTERN As String = "Folio\s[0-9]{6}\-[0-9]{2}"
Private Const SUBJECT_REJECT_PATTERN As String = "Entregado:|Retransmitido:|Leído:|Read:|Aceptada:|Accepted:|Respuesta automática:"
Private Const BODY_CUT_PATTERN As String = "De:"
    ' Configure mailbox
Private Const MAILBOX As String = "afrappe@sociedad-general.mx"
Private Const FOLDER As String = "Bandeja de entrada"

Function GetLastDate() As Date
    On Error GoTo ErrorHandler
    
    Dim lastRow As Long
    With ThisWorkbook.Sheets("Correspondencia")
        lastRow = .Cells(.Rows.Count, 1).End(xlUp).Row
        GetLastDate = .Cells(lastRow, 1).Value
    End With
    Exit Function

ErrorHandler:
    GetLastDate = DateSerial(1900, 1, 1) ' Return default date if error
    Debug.Print "Error in GetLastDate: " & Err.Description
End Function

Function GetSenderEmail(ByVal olMail As Object) As String
    On Error Resume Next
    
    If olMail.Sender.Type = "EX" Then
        GetSenderEmail = olMail.Sender.GetExchangeUser.PrimarySmtpAddress
    Else
        GetSenderEmail = olMail.SenderEmailAddress
    End If
    
    If Err.Number <> 0 Then
        GetSenderEmail = ""
        Debug.Print "Error getting sender email: " & Err.Description
    End If
End Function

Function GetRecipientEmail(ByVal Recipient As Object) As String
    On Error Resume Next
    
    If Recipient.AddressEntry.Type = "EX" Then
        GetRecipientEmail = Recipient.AddressEntry.GetExchangeUser().PrimarySmtpAddress
    Else
        GetRecipientEmail = Recipient.AddressEntry.Address
    End If
    
    If Err.Number <> 0 Then
        GetRecipientEmail = ""
        Debug.Print "Error getting recipient email: " & Err.Description
    End If
End Function

Sub ProcesaCorrespondencia()
    On Error GoTo ErrorHandler
    
    ' Declare variables
    Dim olapp As Outlook.Application
    Dim olNs As Namespace
    Dim Fldr As Outlook.Items
    Dim MFldr As MAPIFolder
    Dim olMail As Object
    Dim ws As Worksheet
    Dim i As Long
    Dim sFilter As String
    Dim lastDate As Date
    
    ' Initialize RegExp objects
    Dim regexSubject As RegExp
    Set regexSubject = InitializeRegex(SUBJECT_REJECT_PATTERN)
    
    ' Get last processed date
    lastDate = GetLastDate()
    
    ' Setup worksheet
    Set ws = ThisWorkbook.Sheets("Proceso")
    InitializeWorksheet ws
    
    ' Setup Outlook
    Set olapp = New Outlook.Application
    Set olNs = olapp.GetNamespace("MAPI")
    

    sFilter = "[ReceivedTime] > '" & Format(lastDate, "dd/mm/yyyy HH:mm") & "'"
    
    Set MFldr = olNs.Folders(MAILBOX).Folders(FOLDER)
    
    ' Process emails
    i = 2
    Dim mailItem As Object
    For Each mailItem In MFldr.Items.Restrict(sFilter)
        If Not regexSubject.Test(mailItem.Subject) Then
            ProcessEmail mailItem, ws, i
            i = i + 1
        End If
        DoEvents
    Next mailItem
    
    Clean:
        Set regexSubject = Nothing
        Set olapp = Nothing
        Set olNs = Nothing
        Exit Sub
        
ErrorHandler:
    MsgBox "Error: " & Err.Description, vbCritical
    Resume Clean
End Sub

Private Function InitializeRegex(ByVal pattern As String) As RegExp
    Set InitializeRegex = New RegExp
    With InitializeRegex
        .Global = True
        .MultiLine = True
        .IgnoreCase = True
        .Pattern = pattern
    End With
End Function

Private Sub InitializeWorksheet(ByVal ws As Worksheet)
    ws.Visible = True
    ws.Activate
    ws.Cells.ClearContents
    
    ' Set headers
    With ws
        .Cells(1, 1) = "Fecha de llegada"
        .Cells(1, 2) = "De"
        .Cells(1, 3) = "De-Email"
        .Cells(1, 4) = "Para"
        .Cells(1, 5) = "Para-Email"
        .Cells(1, 6) = "Tema"
        .Cells(1, 7) = "Hilo"
        .Cells(1, 8) = "UltimoHilo"
    End With
End Sub

Private Sub ProcessEmail(ByVal mailItem As Object, ByVal ws As Worksheet, ByVal row As Long)
    With ws
        .Cells(row, 1) = mailItem.ReceivedTime
        .Cells(row, 2) = mailItem.SenderName
        .Cells(row, 3) = GetSenderEmail(mailItem)
        .Cells(row, 4) = mailItem.Recipients(1).Name
        .Cells(row, 5) = GetRecipientEmail(mailItem.Recipients(1))
        .Cells(row, 6) = mailItem.ConversationTopic
        .Cells(row, 7) = mailItem.ConversationID
        .Cells(row, 8) = Len(mailItem.ConversationIndex)
    End With
End Sub

