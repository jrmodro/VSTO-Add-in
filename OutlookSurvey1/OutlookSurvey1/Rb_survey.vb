Imports Microsoft.Office.Tools.Ribbon
Imports Microsoft.Office.Interop
Imports System.IO

Public Class Rb_survey
    Private Sub Button1_Click(sender As Object, e As RibbonControlEventArgs) Handles Button1.Click

        '-----------------------------------------------------------------------

        ' opção 1

        '' Get inspector
        'Dim currentInspector As Microsoft.Office.Interop.Outlook.Inspector = TryCast(Me.Context, Microsoft.Office.Interop.Outlook.Inspector)

        '' If inspector hasn't been initialized
        'If currentInspector Is Nothing Then
        '    Return
        'End If

        '' Get current mail item from inspector
        'Dim currentMailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(currentInspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
        'If currentMailItem IsNot Nothing Then
        '    ' Read a file and put it on body of current email item
        '    Dim stmReader As New StreamReader("c:\Temp\Survey01.html")
        '    currentMailItem.HTMLBody = stmReader.ReadToEnd()
        'End If


        '-----------------------------------------------------------------------

        'opção 2

        ' Get inspector
        Dim currentInspector As Microsoft.Office.Interop.Outlook.Inspector = TryCast(Me.Context, Microsoft.Office.Interop.Outlook.Inspector)

        ' If inspector hasn't been initialized
        If currentInspector Is Nothing Then
            Return
        End If

        ' Get current mail item from inspector
        Dim currentMailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(currentInspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)
        Dim mySurvey As String = ""

        If currentMailItem IsNot Nothing Then
            ' Read a file and put it on body of current email item
            If My.Computer.FileSystem.FileExists("c:\Temp\Survey01.html") = True Then
                mySurvey = My.Computer.FileSystem.ReadAllText("c:\Temp\Survey01.html")
            Else
                MsgBox("File c:\Temp\Survey01.html not found!", MsgBoxStyle.Information, "Information")
                Exit Sub
            End If

            currentMailItem.HTMLBody = currentMailItem.HTMLBody.Substring(0, InStr(currentMailItem.HTMLBody, "<div class=WordSection1>", vbTextCompare) - 1) & mySurvey
        End If


        '-----------------------------------------------------------------------
        'opção 3

        '' Get inspector
        'Dim currentInspector As Microsoft.Office.Interop.Outlook.Inspector = TryCast(Me.Context, Microsoft.Office.Interop.Outlook.Inspector)

        '' If inspector hasn't been initialized
        'If currentInspector Is Nothing Then
        '    Return
        'End If

        'Dim currentMailItem As Microsoft.Office.Interop.Outlook.MailItem = TryCast(currentInspector.CurrentItem, Microsoft.Office.Interop.Outlook.MailItem)

        '' The full path will place the email in the user's temporary folder
        'Dim strTmpPath As String = "c:\Temp\Survey01.msg"

        '' Save the email to the user's temp folder and convert it to a .MSG
        'currentMailItem.SaveAs(strTmpPath, Outlook.OlSaveAsType.olMSG)

        '' Open the email file and read it into a byte array
        'Dim tmpFile As New FileStream(strTmpPath, FileMode.Open, FileAccess.Read)
        'Dim btSaveFile(tmpFile.Length) As Byte
        'tmpFile.Read(btSaveFile, 0, tmpFile.Length)
        'tmpFile.Close()

        'currentMailItem.HTMLBody = Replace(strTmpPath, "</body>", "c:\Temp\Survey01.html", 1, 1, vbTextCompare)

        '-----------------------------------------------------------------------
        'opção 4

        'If (currentMailItem IsNot Nothing) Then
        '    UploadDocument("http://testURL/sites/testing", "Documents", strTmpPath, btSaveFile)
        'End If
        ' Clean up the temporary .MSG file from the user's temporary folder
        'System.IO.File.Delete(strTmpPath)

        'Dim objOL As New Outlook.Application
        'Dim objInsp As Object

        'objInsp = objOL.ActiveInspector.CurrentItem
        'If Not objInsp Is Nothing Then
        '    'Dim wordDoc As Word.Document

        '    'verifica a existência do arquivo
        '    If My.Computer.FileSystem.FileExists("c:\Temp\Survey01.html") = False Then
        '        MsgBox("File c:\Temp\Survey01.html not found!", MsgBoxStyle.Information, "Information")
        '        Exit Sub
        '    End If

        '    'wordDoc = objInsp.HTMLBody 'objInsp.WordEditor
        '    'wordDoc = objInsp.WordEditor
        '    'wordDoc.Application.Selection.InsertFile("c:\Temp\Survey01.html", , False, False, False)
        'End If

        'objOL = Nothing
        'objInsp = Nothing
    End Sub


End Class
