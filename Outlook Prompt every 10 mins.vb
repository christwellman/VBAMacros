Private WithEvents myReminders As Outlook.Reminders
Const taskSubject As String = "Trigger Task"
Const amountOfTime As Long = 10
 
Private Sub Application_Reminder(ByVal Item As Object)
    StartRemindingMe
End Sub
 
Private Sub Application_Startup()
    StartRemindingMe
End Sub
 
Private Sub myReminders_BeforeReminderShow(Cancel As Boolean)
    Dim remind As Outlook.Reminder
     ' check to make sure we're working on the correct reminder
    Set remind = myReminders.Item(1)
    If remind.Caption = taskSubject Then
        MsgBox ("Chris Here is where you made a change")
    End If
     ' cancel the reminder
    Cancel = True
End Sub
 
Public Sub StartRemindingMe()
    Dim olApp As Outlook.Application
    Dim olNS As Outlook.NameSpace
    Dim tsk As Outlook.TaskItem
    Dim tasksFolder As Outlook.MAPIFolder
    Dim tasks As Outlook.Items
    Dim matchingTasks As Outlook.Items
    Dim i As Long
    Dim task As Outlook.TaskItem
     ' only start watching reminders when I say so
    Set olApp = Outlook.Application
    Set myReminders = olApp.Reminders
     ' delete any existing tasks
    Set olNS = olApp.GetNamespace("MAPI")
    Set tasksFolder = olNS.GetDefaultFolder(olFolderTasks)
    Set tasks = tasksFolder.Items
    Set matchingTasks = tasks.Restrict("[Subject] = '" & taskSubject & "'")
    For i = matchingTasks.Count To 1 Step -1
        Set task = matchingTasks.Item(i)
        If task.subject = taskSubject Then
            With task
                .MarkComplete
                .Delete
            End With
        End If
    Next i
     ' create initial task
    Set tsk = Application.CreateItem(olTaskItem)
    With tsk
        .subject = taskSubject
        .StartDate = Format(Now, "mm/dd/yyyy")
        .ReminderSet = True
        .ReminderTime = DateAdd("n", amountOfTime, Now)
        .Save
    End With
End Sub