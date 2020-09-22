VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsTextValidation"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

'The important bit - note the WithEvents keyword.
Private WithEvents txtMain As TextBox
Attribute txtMain.VB_VarHelpID = -1

Public Property Set TextBoxToValidate(SetTextBox As TextBox)
    'Setting this property connects the control on the form to this instance of this class
    Set txtMain = SetTextBox
End Property

Private Sub txtMain_Change()
    'This is just an example - you could reformat the text, change it's position, or
    'trigger off a different event if you want to. It's a doddle...
    Dim lRandomColour As Long
    
    Randomize Timer
    
    lRandomColour = Int(Rnd * 14) + 1
    
    txtMain.BackColor = QBColor(lRandomColour)
End Sub
