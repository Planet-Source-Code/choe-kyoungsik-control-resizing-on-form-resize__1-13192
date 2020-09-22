VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "CtlDefinition"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'
' CtlDefinition.cls
' - Used in EvtFormResize
' Written by CHOE KyoungSik
' 11/07/2000
'

Option Explicit

Public ctlControl As Control
Public LeftDiff As Double
Public TopDiff As Double
Public WidthDiff As Double
Public HeightDiff As Double

