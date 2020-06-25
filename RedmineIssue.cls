VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "RedmineIssue"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit

Public id As Integer
Public project, tracker, status, priority, author, assigned_to, subject, due_date As String
Public is_public As Boolean

'Constructor
Private Sub Class_Initialize()
End Sub
'Destructor
Private Sub Class_Terminate()
End Sub

Public Function load(ByVal json As Dictionary)
  Me.id = Int(json("id"))
  Me.tracker = json("tracker")("name")
  Me.status = json("status")("name")
  Me.priority = json("priority")("name")
  Me.author = json("author")("name")
  Me.assigned_to = json("assigned_to")("name")
  Me.subject = json("subject")
  
End Function


