VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   4890
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   5610
   LinkTopic       =   "Form1"
   ScaleHeight     =   4890
   ScaleWidth      =   5610
   StartUpPosition =   3  'Windows Default
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'This is the most simple way to use the class, no command line and no custom ID for our application.

Private WithEvents f_cPI As cPrevInstance
Attribute f_cPI.VB_VarHelpID = -1

Private Sub Form_Load()
    Set f_cPI = New cPrevInstance

    If f_cPI.PrevInstance Then 'Is a previous instance running?
        Unload Me 'Unload App
    Else
        '++++++++++++++
        'Your Code here
        '++++++++++++++
    
        f_cPI.Ready = True 'Start parser
        'In this example is not really needed to set Ready = true
        'but is less resource consuming since we gonna discard all
        'the command line messages instead of collecting them.
    End If

End Sub

Private Sub f_cPI_PrevInstance( _
       ByVal sCommand As String, _
       ByVal bReady As Boolean, _
       ByVal Files As Collection, _
       ByVal Folders As Collection, _
       ByVal Parameters As Collection)
       
    f_cPI.ShowForm Me 'Bring our window to front
    
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set f_cPI = Nothing
End Sub
