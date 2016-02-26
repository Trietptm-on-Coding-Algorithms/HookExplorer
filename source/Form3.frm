VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form Form3 
   Caption         =   "Search Results"
   ClientHeight    =   5625
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   10530
   LinkTopic       =   "Form3"
   ScaleHeight     =   5625
   ScaleWidth      =   10530
   StartUpPosition =   2  'CenterScreen
   Begin VB.OptionButton Option1 
      Caption         =   "HookMod"
      Height          =   240
      Index           =   3
      Left            =   3825
      TabIndex        =   8
      Top             =   180
      Value           =   -1  'True
      Width           =   1050
   End
   Begin VB.CommandButton cmdSave 
      Caption         =   "Save Results"
      Height          =   420
      Left            =   8685
      TabIndex        =   7
      Top             =   90
      Width           =   1500
   End
   Begin VB.OptionButton Option1 
      Caption         =   "HookProc"
      Height          =   240
      Index           =   2
      Left            =   2565
      TabIndex        =   5
      Top             =   180
      Width           =   1050
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Inst"
      Height          =   240
      Index           =   1
      Left            =   1800
      TabIndex        =   4
      Top             =   180
      Width           =   600
   End
   Begin VB.OptionButton Option1 
      Caption         =   "Name"
      Height          =   240
      Index           =   0
      Left            =   855
      TabIndex        =   3
      Top             =   180
      Width           =   825
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Search All"
      Height          =   375
      Left            =   6975
      TabIndex        =   2
      Top             =   135
      Width           =   1410
   End
   Begin VB.TextBox txtFind 
      Height          =   330
      Left            =   5040
      TabIndex        =   1
      Top             =   135
      Width           =   1860
   End
   Begin MSComctlLib.ListView lvImports 
      Height          =   4905
      Left            =   0
      TabIndex        =   6
      Top             =   585
      Width           =   10230
      _ExtentX        =   18045
      _ExtentY        =   8652
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "IAT Address"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Value"
         Object.Width           =   1587
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Name"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "1st Instruction"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "HookProc"
         Object.Width           =   2540
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "HookMod"
         Object.Width           =   2540
      EndProperty
   End
   Begin VB.Label Label1 
      Caption         =   "Search"
      Height          =   285
      Left            =   135
      TabIndex        =   0
      Top             =   180
      Width           =   645
   End
End
Attribute VB_Name = "Form3"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdSave_Click()
    
    On Error Resume Next
    Dim f As String
    Dim fhandle As Long
    Dim x As CContainer
    Dim li As ListItem
    Dim tmp() As String
    
    Const header = "Dll\tValue\tName\t1stInst\tHookProc\tHookMod"
    
    If lvImports.ListItems.Count = 0 Then
        MsgBox "No results to export", vbInformation
        Exit Sub
    End If
    
    f = Environ("temp") & "\results.txt" 'get tmp file name
    fhandle = FreeFile
    
    Open f For Output As #fhandle
    Print #fhandle, vbCrLf & "Make sure to save results" & vbCrLf
    Print #fhandle, Replace(header, "\t", vbTab)
                    
    For Each li In lvImports.ListItems
        Form1.push tmp, li.Text
        For i = 1 To lvImports.ColumnHeaders.Count
            Form1.push tmp, li.SubItems(i)
        Next
        Print #fhandle, Join(tmp, ",")
        Erase tmp
    Next
    
    Close #fhandle
         
    Shell "notepad """ & f & """", vbNormalFocus
    
    'Sleep 800
    DoEvents
    Kill f
    
End Sub

Private Sub Command1_Click()
    Dim itemIndex As Long
    Dim li As ListItem
    Dim li2 As ListItem
    Dim names As New Collection
    Dim k As String
    Dim txt As String
    Dim lv As ListView
    
    On Error Resume Next
    
    If Len(txtFind) = 0 Then Exit Sub
    If Option1(0).value = True Then itemIndex = 2 'name
    If Option1(1).value = True Then itemIndex = 3 '1st inst
    If Option1(2).value = True Then itemIndex = 4 'hook proc
    If Option1(3).value = True Then itemIndex = 5 'hook module
     
    lvImports.ListItems.Clear
    
    If Form1.lvBound2.Visible Then
        Set lv = Form1.lvBound2
    Else
        Set lv = Form1.lvBound
    End If
    
    For Each li In lv.ListItems
    
        If Form1.lvBound2.Visible Then
            Form1.lvBound2_ItemClick li
        Else
            Form1.lvBound_ItemClick li
        End If
    
        For Each li2 In Form1.lvImports.ListItems
            txt = li2.SubItems(itemIndex)
            k = li2.SubItems(2) & ":" & li2.SubItems(5) 'name:hookmod pair as key
            If InStr(1, txt, txtFind, vbTextCompare) > 0 Then
                If Not KeyExistsInCollection(names, k) Then
                    cloneLi li2, li.SubItems(3), Me.lvImports
                    names.Add k, k
                End If
            End If
        Next
    Next
                    
            
        
End Sub

'assumes you have same number of columns in each already..(copy past listview)
Sub cloneLi(li As ListItem, dll As String, toLv As ListView)
    
    On Error Resume Next
    
    Dim li2 As ListItem
    Set li2 = toLv.ListItems.Add
    'li2.Text = li.Text
    li2.Text = dll
    
    For i = 1 To toLv.ColumnHeaders.Count - 1
        li2.SubItems(i) = li.SubItems(i)
    Next
    
End Sub


Function KeyExistsInCollection(c As Collection, val As String) As Boolean
    On Error GoTo nope
    Dim t
    t = c(val)
    KeyExistsInCollection = True
 Exit Function
nope: KeyExistsInCollection = False
End Function

Private Sub Form_Resize()
    On Error Resume Next
    
    If Me.WindowState = vbMinimized Then Exit Sub
    
    If Me.Height < 8100 Then Me.Height = 8100
    If Me.Width < 9660 Then Me.Width = 9660
    
    lvImports.Width = Me.Width - lvImports.Left - 200
    SizeLV lvImports
    
    lvImports.Height = Me.Height - lvImports.Top - 400
        
End Sub

Sub SizeLV(lv As ListView)
    Dim c As Long
    With lv
        c = .ColumnHeaders.Count
        .ColumnHeaders(c).Width = .Width - lv.ColumnHeaders(c).Left - 300
    End With
End Sub

