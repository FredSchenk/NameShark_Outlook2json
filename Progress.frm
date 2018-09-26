VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} Progress 
   Caption         =   "Progress"
   ClientHeight    =   6450
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   9690
   OleObjectBlob   =   "Progress.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "Progress"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False



Sub UpdateCounters(iDone, iChanged, sFullName)
    Me.tChanged = iChanged
    Me.tDone = iDone
    If ("" <> sFullName) Then Me.lstChangedNames.AddItem sFullName
    Me.tPercText.Text = Int(100 * Me.tDone / Me.tTotal) & " %"
    If (Me.tDone / Me.tTotal > 0.5) Then Me.tPercText.ForeColor = vbWhite Else Me.tPercText.ForeColor = vbBlack
    Me.tPercDone.Width = Int(Me.tPercentage.Width * Int(Me.tDone) / Int(Me.tTotal))
    Me.Repaint
    If (iDone Mod 10 = 0) Then
        DoEvents
    End If
End Sub

Private Sub btnClose_Click()
    Me.Hide
    Unload Me
End Sub

Sub CopyToClipboard()
Dim t As String
Dim MyData As DataObject
    t = ""
    For i = 0 To Me.lstChangedNames.ListCount - 1
        t = t & Me.lstChangedNames.List(i) & vbNewLine
    Next
    Set MyData = New DataObject
    MyData.SetText t
    MyData.PutInClipboard
    Set MyData = Nothing
End Sub
