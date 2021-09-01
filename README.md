Sub Send_Email()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Auto Email")

'''''''''' Update Next ''''''''''''''

Call Update_Next_Schedule_Time
Application.OnTime sh.Range("N24").Value, "Send_Email"

''''''''''''''''''''''''''''''''''''''

Dim oa As Object
Dim msg As Object

Set oa = CreateObject("outlook.application")
Set msg = oa.CreateItem(0)

With msg
.To = sh.Range("E6").Value
.CC = sh.Range("E8").Value
.Subject = sh.Range("E10").Value
.Body = sh.Range("E12").Value
.Attachments.Add "C:\Users\info\Desktop\Sample File.txt"
.Send
End With

End Sub
Sub Update_Next_Schedule_Time()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Auto Email")
sh.Unprotect
sh.Range("N24").Value = ""

Dim i As Integer

Dim dt As Date

If Time > Application.WorksheetFunction.Max(sh.Range("N6:N40")) Then
dt = Date + 1
Else
dt = Date
End If

If UCase(Format(dt, "DDD")) = "SAT" Then
If sh.Range("O22").Value = True And sh.Range("O23").Value = False Then
dt = dt + 1
ElseIf sh.Range("O22").Value = True And sh.Range("O23").Value = True Then
dt = dt + 2
End If
ElseIf UCase(Format(Date, "DDD")) = "SUN" Then
If sh.Range("O23").Value = True Then
dt = dt + 1
End If
End If


If dt > Date Then
sh.Range("N24").Value = dt + Application.WorksheetFunction.Min(sh.Range("N6:N20"))
Else

For i = 6 To 20
If Time < sh.Range("N" & i).Value Then
sh.Range("N24").Value = sh.Range("N" & i).Value + dt
Exit For
End If
Next i
End If

sh.Protect

End Sub
Sub Set_Schedule()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Auto Email")
Call Update_Next_Schedule_Time

Application.OnTime sh.Range("N24").Value, "Send_Email"

MsgBox "Schedule Set"

End Sub
Sub Cancel_Schedule()

Dim sh As Worksheet
Set sh = ThisWorkbook.Sheets("Auto Email")

On Error Resume Next
Application.OnTime sh.Range("N24").Value, "Send_Email", , False

MsgBox "Schedule Cancelled"

End Sub

++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++++
Sub ConnectingToOutlook()
Dim out As Outlook.Application
Set out = New Outlook.Application

Dim mit As Outlook.MailItem
Set mit = out.CreateItem(olMailItem)


mit.To = "Suryanalajala@gmail.com"
mit.CC = "Surya.analyst13@gmail.com"
mit.BCC = "surya.nalajala@outlook.com"

mit.Subject = "Regarding Testing Outlook Macro"
mit.Body = "Hello Surya," & vbNewLine & "we are testing a outlook macro.Please dont repond to this mail." & vbNewLine & "Thanks," & vbNewLine & "XYZ"
mit.Attachments.Add "C:\Users\mpkum\Desktop\Simha.xlsx"

mit.Display
mit.Send

End Sub

Sub SendingBulkEmails()

Dim out As Outlook.Application
Set out = New Outlook.Application

Dim mit As Outlook.MailItem


Dim lr As Integer
lr = Range("A" & Rows.Count).End(xlUp).Row

For i = 2 To lr

    Set mit = out.CreateItem(olMailItem)

    mit.To = Range("B" & i).Value
    mit.CC = Range("C" & i).Value
    mit.BCC = Range("D" & i).Value

   mit.Subject = "Regarding Testing Outlook Macro"
   mit.Body = "Hello " & Range("A" & i).Value & "," & vbNewLine & "we are testing a outlook macro.Please dont repond to this mail." & vbNewLine & "Thanks," & vbNewLine & "XYZ"
   mit.Attachments.Add "C:\Users\mpkum\Desktop\Simha.xlsx"

   mit.Display
   mit.Send

Next i

End Sub


Sub ConnectingToSqlServer()              
Dim Con As ADODB.Connection
Set Con = New ADODB.Connection

Dim Rs As ADODB.Recordset
Set Rs = New ADODB.Recordset

Dim Fd As ADODB.Field

Con.ConnectionString = "Provider=SQLOLEDB.1;Integrated Security=SSPI;Persist Security Info=True;Initial Catalog=AdventureWorks2012;Data Source=DESKTOP-AEF2VP2;"

Con.Open
       
       Rs.ActiveConnection = Con
       Rs.LockType = adLockReadOnly
       Rs.CursorType = adOpenDynamic
       'Rs.Source = "Employees"
       Rs.Source = "Select * from Employees where Salary >=50000"
       Rs.Open
       Range("A2").CopyFromRecordset Rs
       i = 1
       For Each Fd In Rs.Fields
       
        Cells(1, i).Value = Fd.Name
       i = i + 1
       Next Fd
       
       Rs.Close
Con.Close

End Sub
