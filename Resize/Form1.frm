VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   6090
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   7485
   LinkTopic       =   "Form1"
   MDIChild        =   -1  'True
   ScaleHeight     =   6090
   ScaleWidth      =   7485
   WindowState     =   2  'Maximized
   Begin VB.Frame Frame1 
      Caption         =   "Frame1"
      Height          =   675
      Left            =   120
      TabIndex        =   2
      Top             =   90
      Width           =   7215
      Begin VB.TextBox Text1 
         Height          =   315
         Left            =   240
         TabIndex        =   3
         Text            =   "Text1"
         Top             =   240
         Width           =   6735
      End
   End
   Begin VB.TextBox Text3 
      Height          =   5055
      Left            =   2760
      TabIndex        =   1
      Text            =   "Text3"
      Top             =   870
      Width           =   4575
   End
   Begin VB.TextBox Text2 
      Height          =   5055
      Left            =   120
      TabIndex        =   0
      Text            =   "Text2"
      Top             =   870
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'If you find any bugs or improvements,
'Please let me know : gould0711@yahoo.com.

'For resizing, you should implement two event procedures.
'Form_Resize() and Form_Unload(), like these.
Private Sub Form_Resize()
    'Frame1 is parent of Text1
    'So, if CtlToResize is not provided to Frame1 explicitly,
    'Frame1 goes with Text1's definition(its child control).
    Call EvtFormResize(Me, CtlToResize(Text1, 0, 1, 0, 0), _
                           CtlToResize(Text2, 0, 0, 0, 1), _
                           CtlToResize(Text3, 0, 1, 0, 1))
                           
'or like this
'    Call EvtFormResize(Me, CtlToResize(Text1, 0, 1, 0, 0.2), _
                           CtlToResize(Text2, 0, 0, 0.2, 0.8), _
                           CtlToResize(Text3, 0, 1, 0.2, 0.8))

'cFormDesignedWidth and cFormDesignedHeight should be the same value
'of this form designed. If not, you can give that values like this:
'    Call EvtFormResize(Me, 7600, 6400, _
                           CtlToResize(Text1, 0, 1, 0, 0), _
                           CtlToResize(Text2, 0, 0, 0, 1), _
                           CtlToResize(Text3, 0, 1, 0, 1))

'If you are using ListBox or DBList control,
'you should redefine its Height here
'because their heights are tend to shrink.
End Sub

Private Sub Form_Unload(Cancel As Integer)
    Call EvtFormUnload(Me, Cancel)
End Sub
