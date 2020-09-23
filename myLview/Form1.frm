VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmView 
   Caption         =   "Listview with Alternate Color Plus manual edit..."
   ClientHeight    =   7410
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   10170
   Icon            =   "Form1.frx":0000
   LinkTopic       =   "Form1"
   LockControls    =   -1  'True
   ScaleHeight     =   7410
   ScaleWidth      =   10170
   StartUpPosition =   3  'Windows Default
   Begin MSComctlLib.ListView popList 
      Height          =   7095
      Left            =   150
      TabIndex        =   0
      Top             =   120
      Width           =   9825
      _ExtentX        =   17330
      _ExtentY        =   12515
      View            =   3
      LabelWrap       =   -1  'True
      HideSelection   =   0   'False
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   5
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Column1"
         Object.Width           =   4410
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Column2"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Column3"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Column4"
         Object.Width           =   3528
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Button"
         Object.Width           =   1411
      EndProperty
   End
End
Attribute VB_Name = "frmView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private WithEvents tLBox  As VB.TextBox
Attribute tLBox.VB_VarHelpID = -1
Private WithEvents bbTon   As VB.CommandButton
Attribute bbTon.VB_VarHelpID = -1
Private WithEvents ccBox   As VB.ComboBox
Attribute ccBox.VB_VarHelpID = -1

Dim x1                  As Integer
Dim y1                  As Integer

Private Sub bTon_Click()
MsgBox popList.SelectedItem.Text
End Sub

Private Sub Form_Load()
Dim hHeader As Long, lStyle As Long
Dim sItem As ListItem
Dim i As Integer

For i = 1 To 50
Set sItem = popList.ListItems.Add(, , "Col1 Row" & i)
    sItem.SubItems(1) = "Col2 Row" & i
    sItem.SubItems(2) = "Col3 Row" & i
    sItem.SubItems(3) = "Col4 Row" & i
Next

'call this procedure only if you have finish populate the data on your listview
'or else it will not paint your Listview
Call AltBckColor(Me, popList, RGB(167, 197, 218), &H80000018)

LV_FlatHeaders Me.hWnd, popList.hWnd

End Sub

Private Sub popList_DblClick()
If (GetKeyState(vbKeyLButton) And &H8000) Then
    Set tLBox = AttachList(Me, popList, x1, y1, tLBox)
End If
End Sub

Private Sub popList_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
'be sure to initialeize when your mouse click on an item in Listview
'to get it's current position
x1 = x
y1 = y

End Sub

Private Sub popList_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
x1 = x
y1 = y
Set bbTon = AttachButton(Me, popList, x1, y1, bbTon, 4)
Set ccBox = AttachCbox(Me, popList, x1, y1, ccBox, 3)
End Sub

Private Sub tLBox_LostFocus()
If tLBox.Visible = True Then tLBox.Visible = False
tLBox.Text = ""
End Sub

Private Sub tLBox_Keypress(Keyascii As Integer)
    If (Keyascii = vbKeyReturn) Then
        Call hideTbox(True)
    ElseIf (Keyascii = vbKeyEscape) Then
        Call hideTbox(False)
    End If
End Sub

Friend Sub hideTbox(Apply As Boolean)
    If Apply Then
        popList.ListItems(tHT.lItem + 1).SubItems(tHT.lSubItem) = tLBox
    Else
        popList.ListItems(tHT.lItem + 1).SubItems(tHT.lSubItem) = tLBox.Tag
    End If
    tLBox.Visible = False
    tLBox = ""
End Sub

