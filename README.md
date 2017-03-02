# Email
Email JPEG



Sub Email()
    Send_Email_SG
    Send_Email_CN
End Sub
Sub Send_Email_SG()
    Dim Email_Subject, Email_Send_From, Email_Send_To, _
    Email_Cc, Email_Bcc, Email_Body As String
    Dim Mail_Object, Mail_Single As Variant
    Dim rng As Range
    Dim rptDate As String
    Dim rptD As Date
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExter As String
    Dim activeFullName As String
    Dim wb As Workbook
    Dim tmpPic As String
    Dim oChart As Chart
    Dim oImg As Picture
    Dim i As Integer
    Dim last As Long
    
    rptD = Date
    rptDate = Format(Now() - 1, "dd mmm yyyy")
    Email_Subject = "Daily Attribution Report (ASIA/ASEAN) - " & rptDate
    Email_Send_From = ActiveWorkbook.Sheets("Emails").Range("C2:C2").Value
    Email_Send_To = ActiveWorkbook.Sheets("Emails").Range("C3:C3").Value
    Email_Cc = ActiveWorkbook.Sheets("Emails").Range("C4:C4").Value
    Email_Bcc = ""
    
    last = ThisWorkbook.Worksheets("Report").Range("D10000").End(xlUp).Row
    ThisWorkbook.Worksheets("Report").Activate
    ActiveWindow.DisplayGridlines = False
    createJpg "Report", "D8:W" & last, "report"
    ActiveWindow.DisplayGridlines = True
    Worksheets("Print").Activate
    TempFilePath = Cells(8, 5).Value & "\"
    'activeFullName = ActiveWorkbook.FullName
    'TempFilePath = Left(activeFullName, InStrRev(activeFullName, "\")) 'Environ$("temp") & "\"
'    TempFileName = "Return Page.xls"
'    FileExter = TempFilePath & TempFileName
'    Set wb = Workbooks.Open(FileExter)
'    On Error GoTo 0
'
'    With wb.Sheets("Report")
'        tmpPic = "c:\Hiport\Return.gif" 'TempFilePath & "Return.PNG" --- Environ$("temp") &
'        Set oChart = Charts.Add
'        Set Rng = wb.Sheets("Report").Range("D9:Z50")
'        For i = oChart.SeriesCollection.Count To 1 Step -1
'            oChart.Legend.LegendEntries(i).Delete
'        Next i

'        Rng.CopyPicture xlScreen, xlBitmap
'        oChart.Paste
'        oChart.Export FileName:=tmpPic, filtername:="GIF"
'
'    End With
    On Error GoTo debugs
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
    With Mail_Single
    .Subject = Email_Subject
    .To = Email_Send_To
    .CC = Email_Cc
    .BCC = Email_Bcc
    .Attachments.Add "G:\FMDPERF\FACTSET ATTRIBUTION DAILY\report.bmp"
    '.Body = Email_Body
    .HTMLBody = "Hi,<br>" & _
                "The attribution reports are ready in the folder: <br>" & _
                TempFilePath & "<br><img src ='cid:report.bmp'><br>" & _
                "Regards,<br>"
    '.Attachments.Add tmpPic
    
    '.Display
    .Send
    End With
    'Kill tmpPic
    'wb.Saved = True
    'wb.Close
debugs:
    If Err.Description <> "" Then MsgBox Err.Description
End Sub

Private Sub SaveRngAsBMP(rng As Range, FileName As String)
rng.CopyPicture xlScreen, xlBitmap
SavePicture PastePicture(xlBitmap), FileName
End Sub

Function RangetoHTML(rng As Range)
' Changed by Ron de Bruin 28-Oct-2006
' Working in Office 2000-2013
    Dim fso As Object
    Dim ts As Object
    Dim TempFile As String
    Dim TempWB As Workbook

    TempFile = Environ$("temp") & "\" & Format(Now, "dd-mm-yy h-mm-ss") & ".htm"

    'Copy the range and create a new workbook to past the data in
    rng.Copy
    Set TempWB = Workbooks.Add(1)
    With TempWB.Sheets(1)
        .Cells(1).PasteSpecial Paste:=8
        .Cells(1).PasteSpecial xlPasteValues, , False, False
        .Cells(1).PasteSpecial xlPasteFormats, , False, False
        .Cells(1).Select
        Application.CutCopyMode = False
        On Error Resume Next
        .DrawingObjects.Visible = True
        .DrawingObjects.Delete
        On Error GoTo 0
    End With

    'Publish the sheet to a htm file
    With TempWB.PublishObjects.Add( _
         SourceType:=xlSourceRange, _
         FileName:=TempFile, _
         Sheet:=TempWB.Sheets(1).Name, _
         Source:=TempWB.Sheets(1).UsedRange.Address, _
         HtmlType:=xlHtmlStatic)
        .Publish (True)
    End With

    'Read all data from the htm file into RangetoHTML
    Set fso = CreateObject("Scripting.FileSystemObject")
    Set ts = fso.GetFile(TempFile).OpenAsTextStream(1, -2)
    RangetoHTML = ts.readall
    ts.Close
    RangetoHTML = Replace(RangetoHTML, "align=center x:publishsource=", _
                          "align=left x:publishsource=")

    'Close TempWB
    TempWB.Close savechanges:=False

    'Delete the htm file we used in this function
    Kill TempFile

    Set ts = Nothing
    Set fso = Nothing
    Set TempWB = Nothing
End Function

Sub Send_Email_CN()
    Dim Email_Subject, Email_Send_From, Email_Send_To, _
    Email_Cc, Email_Bcc, Email_Body As String
    Dim Mail_Object, Mail_Single As Variant
    Dim rptDate As String
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim tmpReturnFile As String
    Dim tmpFCQFile As String
    Dim tmpCNESFile As String
    Dim last As Long

    rptDate = Format(Now() - 1, "dd mmm yyyy")
    Email_Subject = "Daily Attribution Report - " & rptDate
    Email_Send_From = ActiveWorkbook.Sheets("Emails").Range("C6:C6").Value
    Email_Send_To = ActiveWorkbook.Sheets("Emails").Range("C7:C7").Value
    Email_Cc = ActiveWorkbook.Sheets("Emails").Range("C8:C8").Value
    Email_Bcc = ""
    last = ThisWorkbook.Worksheets("Report").Range("D10000").End(xlUp).Row
    ThisWorkbook.Worksheets("Report").Activate
    ActiveWindow.DisplayGridlines = False
    createJpg "Report", "D8:W" & last, "report"
    ActiveWindow.DisplayGridlines = True
    Worksheets("Print").Activate
    TempFilePath = Cells(8, 5).Value & "\"
    Worksheets("Print").Activate
    TempFilePath = Cells(8, 5).Value & "\"
    'activeFullName = ActiveWorkbook.FullName
    'TempFilePath = Left(activeFullName, InStrRev(activeFullName, "\")) 'Environ$("temp") & "\"
    TempFileName = "Return Page.xls"
    tmpReturnFile = TempFilePath & TempFileName
    TempFileName = "LCNA.xls"
    tmpFCQFile = TempFilePath & TempFileName
    TempFileName = "CNES.xls"
    tmpCNESFile = TempFilePath & TempFileName
    On Error GoTo 0
    
    On Error GoTo debugs
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
    With Mail_Single
    .Subject = Email_Subject
    .To = Email_Send_To
    .CC = Email_Cc
    .BCC = Email_Bcc
    .Attachments.Add "G:\FMDPERF\FACTSET ATTRIBUTION DAILY\report.bmp"
    '.Body = Email_Body
    '.HTMLBody = "<img src='" & tmpPic & "'>" 'height=480 width=360
    .HTMLBody = "<img src ='cid:report.bmp'><br>Regards,<br>"
    .Attachments.Add tmpReturnFile
    .Attachments.Add tmpFCQFile
    .Attachments.Add tmpCNESFile
    .Send
    End With
debugs:
    If Err.Description <> "" Then MsgBox Err.Description
End Sub
Sub Send_Email_LGEM()

 Dim Email_Subject, Email_Send_From, Email_Send_To, _
    Email_Cc, Email_Bcc, Email_Body As String
    Dim Mail_Object, Mail_Single As Variant
    Dim rng As Range
    Dim rptDate As String
    Dim rptD As Date
    Dim TempFilePath As String
    Dim TempFileName As String
    Dim FileExter As String
    Dim activeFullName As String
    Dim wb As Workbook
    Dim tmpPic As String
    Dim oChart As Chart
    Dim oImg As Picture
    Dim i As Integer
    
    
    rptD = Date
    rptDate = Format(Now() - 1, "dd mmm yyyy")
    Email_Subject = "Daily Attribution Report_LGEM - " & rptDate
    Email_Send_From = ActiveWorkbook.Sheets("Emails").Range("C10:C10").Value
    Email_Send_To = ActiveWorkbook.Sheets("Emails").Range("C11:C11").Value
    Email_Cc = ActiveWorkbook.Sheets("Emails").Range("C12:C12").Value
    Email_Bcc = ""

    last = ThisWorkbook.Worksheets("Report_GEM").Range("D10000").End(xlUp).Row
    ThisWorkbook.Worksheets("Report_GEM").Activate
    ActiveWindow.DisplayGridlines = False
    createJpg "Report_GEM", "D8:W" & last, "report"
    ActiveWindow.DisplayGridlines = True
    Worksheets("Print").Activate
    TempFilePath = Cells(8, 5).Value & "\"
    
On Error GoTo debugs
    Set Mail_Object = CreateObject("Outlook.Application")
    Set Mail_Single = Mail_Object.CreateItem(0)
    With Mail_Single
    .Subject = Email_Subject
    .To = Email_Send_To
    .CC = Email_Cc
    .BCC = Email_Bcc
    .Attachments.Add "G:\FMDPERF\FACTSET ATTRIBUTION DAILY\report.bmp"
    '.Body = Email_Body
    .HTMLBody = "Hi,<br>" & _
                "The attribution reports are ready in the folder: <br>" & _
                TempFilePath & "<br><img src ='cid:report.bmp'><br>" & _
                "Regards,<br>"
    '.Attachments.Add tmpPic
    
    '.Display
    .Send
    End With
    'Kill tmpPic
    'wb.Saved = True
    'wb.Close
debugs:
    If Err.Description <> "" Then MsgBox Err.Description

End Sub


Sub createJpg(Namesheet As String, nameRange As String, nameFile As String)
    ThisWorkbook.Activate
    Worksheets(Namesheet).Activate
    Set Plage = ThisWorkbook.Worksheets(Namesheet).Range(nameRange)
    Plage.CopyPicture
    With ThisWorkbook.Worksheets(Namesheet).ChartObjects.Add(Plage.Left, Plage.Top, Plage.Width, Plage.Height)
        .Activate
        .Chart.Paste
        .Chart.Export "G:\FMDPERF\FACTSET ATTRIBUTION DAILY" & "\" & nameFile & ".bmp", "BMP"
    End With
    Worksheets(Namesheet).ChartObjects(Worksheets(Namesheet).ChartObjects.Count).Delete
Set Plage = Nothing
End Sub

