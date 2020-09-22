VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   4860
   ClientLeft      =   60
   ClientTop       =   450
   ClientWidth     =   6825
   LinkTopic       =   "Form2"
   ScaleHeight     =   4860
   ScaleWidth      =   6825
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox List1 
      Height          =   4155
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   6315
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Another simple example, this time using the command line parameters

Private WithEvents f_cPI As cPrevInstance
Attribute f_cPI.VB_VarHelpID = -1

Private Sub Form_Load()
    Set f_cPI = New cPrevInstance

    If f_cPI.PrevInstance Then 'Is a previous instance running?
    
        'NOTE: the command line (if exists) will be sended automatically
        'to the first instance when we call PrevInstance.
        'To send custom data to the first instance use the optional parameter
        'sCommand of the PrevInstance function.
        
        Unload Me 'Unload App
    Else
        'By default the Ready property is set to false, that means that
        'the class is gonna collect all the command line parameters
        'until we are ready. When you set Ready to true the class will
        'parse the command line collection one by one and raise the PrevInstance
        'event passig all the data.
        
        'Why the class is doing this? Well, supouse that your application need some
        'kind of user input or authentication in order to start runing or processing
        'files so if you have the first instance of it, waiting for that user input or
        'authentication (not ready to process any command line parameter) and the user
        'opens a new instance of it, we gonna receive a new command line parameter,
        'since we are not ready yet the class will save the new command line in a collection
        'Then when we are ready we set the Ready property to true and the class will start
        'parsing and raising then PrevInstance event one by one.
        
        'IMPORTANT: If Ready is False the PrevInstance event will be raised anyways, but
        'with empty paramters and bReady = False, why? ok as I meantioned before if
        'you are waiting for some user input or authentication, you can bring the input/authentication
        'window to the top, everytime the event is raised and bReady = False
        
        '++++++++++++++
        'Your Code here
        '++++++++++++++
        
        f_cPI.Ready = True 'Start the parser
        
        'If this is the first instance and we have a command line parameter
        'the PrevInstance event will be raised and the command line
        'will be passed as sCommand, and the parsed data in their respective
        'collections.
    End If
End Sub

Private Sub f_cPI_PrevInstance( _
       ByVal sCommand As String, _
       ByVal bReady As Boolean, _
       ByVal Files As Collection, _
       ByVal Folders As Collection, _
       ByVal Parameters As Collection)
       
    f_cPI.ShowForm Me
    
    If bReady Then
    'If bReady is false sCommand, Files, Folders, Parameters will be empty
    'so we gonna ignore all this
        Dim vItem As Variant
    
        For Each vItem In Files
            List1.AddItem "File: " & vItem
        Next
    
        For Each vItem In Folders
            List1.AddItem "Folder: " & vItem
        Next
    
        For Each vItem In Parameters
            List1.AddItem "Parameter: " & vItem
        Next
        Debug.Print sCommand
    End If
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Set f_cPI = Nothing
End Sub

