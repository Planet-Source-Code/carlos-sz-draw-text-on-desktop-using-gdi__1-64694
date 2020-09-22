VERSION 5.00
Begin VB.Form frmMain 
   BackColor       =   &H00E0E0E0&
   BorderStyle     =   0  'None
   Caption         =   "GDI+ Text"
   ClientHeight    =   2580
   ClientLeft      =   495
   ClientTop       =   -75
   ClientWidth     =   3615
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2580
   ScaleWidth      =   3615
   ShowInTaskbar   =   0   'False
   Begin VB.CommandButton cmdDraw 
      Caption         =   " Draw "
      Height          =   495
      Left            =   1200
      TabIndex        =   0
      Top             =   1440
      Width           =   1215
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "Press Esc to exit"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   225
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   1455
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "GDI+ is used to draw a nice text on the desktop. This is not an application, just an example of GDI+'s flexibility. Have fun!"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3375
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'---------------------------------------------------------------------------------------
'
' Module    : GDI+ Text
' DateTime  : 18/03/2006 04:11 AM
' Author    : Carlos Alberto S.
' Purpose   : Draw a nice text on the desktop using GDI+
' Note      : This code is just an example and it was done considering some special
'             needs. It was not cleaned up as it could. But you'll be able to follow
'             the code.
' Credits   : http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=43004&lngWId=1
'             http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=42861&lngWId=1
'             http://www.pscode.com/vb/scripts/ShowCode.asp?txtCodeId=45451&lngWId=1
'             And KSoft for GDI+ hints
'---------------------------------------------------------------------------------------

Option Explicit

'move the form and text
Private OldX As Long
Private OldY As Long
Private MoveIt As Boolean

Private Sub cmdDraw_Click()

    frmText.Show 0, Me
    cmdDraw.Enabled = False

End Sub

Private Sub Form_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    OldX = X
    OldY = Y
    MoveIt = True

End Sub

Private Sub Form_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MoveIt Then
        Me.Left = Me.Left + X - OldX
        Me.Top = Me.Top + Y - OldY
        frmText.Top = frmMain.Top + frmMain.Height + 80
        frmText.Left = (frmMain.Width * 0.5) + frmMain.Left - (frmText.Width * 0.5)
    End If

End Sub

Private Sub Form_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveIt = False

End Sub

Private Sub Form_Unload(Cancel As Integer)

    Unload frmText
    
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)

    If KeyCode = vbKeyEscape Then
        Unload Me
    End If
    
End Sub

Private Sub Label1_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)

    OldX = X
    OldY = Y
    MoveIt = True

End Sub

Private Sub Label1_MouseMove(Button As Integer, Shift As Integer, X As Single, Y As Single)

    If MoveIt Then
        Me.Left = Me.Left + X - OldX
        Me.Top = Me.Top + Y - OldY
        frmText.Top = frmMain.Top + frmMain.Height + 80
        frmText.Left = (frmMain.Width * 0.5) + frmMain.Left - (frmText.Width * 0.5)
    End If

End Sub

Private Sub Label1_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)

    MoveIt = False

End Sub
