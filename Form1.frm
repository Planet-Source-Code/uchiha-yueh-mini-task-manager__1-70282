VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form Form1 
   AutoRedraw      =   -1  'True
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Mini-Task Manager"
   ClientHeight    =   4335
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   5655
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   MaxButton       =   0   'False
   ScaleHeight     =   4335
   ScaleWidth      =   5655
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView lvProcess 
      Height          =   3975
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   7011
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin VB.CommandButton cmdEndTask 
      Caption         =   "&End Task"
      Enabled         =   0   'False
      Height          =   375
      Left            =   4560
      TabIndex        =   1
      Top             =   3960
      Width           =   1095
   End
   Begin VB.CommandButton cmdRefresh 
      Caption         =   "&Refresh"
      Height          =   375
      Left            =   3360
      TabIndex        =   0
      Top             =   3960
      Width           =   1095
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim PName As String, PID As Long

Private Sub cmdEndTask_Click()
Dim Process, ans

ans = MsgBox("End '" & PName & "'?", vbQuestion + vbYesNo, "Are you sure?")

If ans = vbNo Then Exit Sub

For Each Process In GetObject("winmgmts:"). _
    ExecQuery("select name from Win32_Process where name='" & PName & _
    "' and processid ='" & PID & "'")
    Process.Terminate (0)
Next

Call cmdRefresh_Click

End Sub

Private Sub cmdRefresh_Click()
Dim Process
Dim lv As ListItem

lvProcess.ListItems.Clear

For Each Process In GetObject("winmgmts:"). _
    ExecQuery("select * from Win32_Process")
    
    Set lv = lvProcess.ListItems.Add(, , Process.Name)
    lv.SubItems(1) = Process.ProcessID
   
Next

cmdEndTask.Enabled = False

End Sub

Private Sub Form_Load()

With lvProcess.ColumnHeaders

    .Add , , "Process", 3900
    .Add , , "Process ID", 1000
    
End With

Call cmdRefresh_Click

End Sub

Private Sub lvProcess_ItemClick(ByVal Item As MSComctlLib.ListItem)

PName = Item.Text
PID = Item.SubItems(1)

cmdEndTask.Enabled = True

End Sub
