VERSION 5.00
Begin VB.Form frmMain 
   Caption         =   "WithEvents Example"
   ClientHeight    =   3195
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   4680
   LinkTopic       =   "Form1"
   ScaleHeight     =   3195
   ScaleWidth      =   4680
   StartUpPosition =   3  'Windows Default
   Begin VB.TextBox Text4 
      Height          =   345
      Left            =   330
      TabIndex        =   3
      Text            =   "Won't Change Colour"
      Top             =   1245
      Width           =   1785
   End
   Begin VB.TextBox Text3 
      Height          =   345
      Left            =   630
      TabIndex        =   2
      Tag             =   "Validate"
      Text            =   "Will Change Colour"
      Top             =   2010
      Width           =   2400
   End
   Begin VB.TextBox Text2 
      Height          =   1035
      Left            =   2775
      TabIndex        =   1
      Tag             =   "Validate"
      Text            =   "Will Change Colour"
      Top             =   705
      Width           =   1530
   End
   Begin VB.TextBox Text1 
      Height          =   465
      Left            =   540
      TabIndex        =   0
      Tag             =   "Validate"
      Text            =   "Will Change Colour"
      Top             =   330
      Width           =   1515
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'String to add to TAG property of textbox
Private Const VALIDATE As String = "Validate"

'Collection of validation classes
Private mValidatedTextBoxes As Collection

Private Sub Form_Load()
    Dim clsValidateTMP As clsTextValidation 'Temporary class object to be added to collection
    Dim ctlLoopTMP As Control 'Temporary control object to use when looping through controls
    
    Set mValidatedTextBoxes = New Collection 'Create a new, empty collection
    
    For Each ctlLoopTMP In Me.Controls 'Loop through all controls on form
        
        If TypeOf ctlLoopTMP Is TextBox Then 'If the current control is a textbox...
            
            If ctlLoopTMP.Tag = VALIDATE Then '...and it has "Validate" in it's tag property...
                
                'Create a new instance of the validation class
                Set clsValidateTMP = New clsTextValidation
                
                'Connect it to the current control
                Set clsValidateTMP.TextBoxToValidate = ctlLoopTMP
                
                'Add it to the collection
                mValidatedTextBoxes.Add clsValidateTMP
                
            End If
            
        End If
        
    Next
    
    'Tidy up object references
    Set clsValidateTMP = Nothing
    Set ctlLoopTMP = Nothing
End Sub
