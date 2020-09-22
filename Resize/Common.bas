Attribute VB_Name = "Common"
'
' Common.bas
' - Common Client Common module
' Written by CHOE KyoungSik
' 11/03/2000
'

Option Explicit

Const cModuleName = "Common"

'Private Constants
Const cFormDesignedWidth = 7600
Const cFormDesignedHeight = 6400

' EvtFormResize Related
Private Type FormSizeType
    strForm As String
    vntWidth As Variant
    vntHeight As Variant
    vntDesWidth As Integer      ' Designed Width
    vntDesHeight As Integer     ' Designed Height
    blnCustomSize As Boolean    ' Designed size is not standard
    blnHasControlToResize As Boolean
End Type
    
Private aFormSize() As FormSizeType
Private blnFormSize As Boolean

'==========================================================================
'           ********** Form Event Procedures **********
'==========================================================================
Public Sub EvtMDIFormLoad(frmForm As Form)
    frmForm.Height = cFormDesignedHeight + 660
    frmForm.Width = cFormDesignedWidth + 60
    frmForm.Left = (Screen.Width - frmForm.Width) / 2
    frmForm.Top = (Screen.Height - frmForm.Height) / 2
End Sub

'MDI Child Resize func
'adefResizable :
'   CtlDefinition's object array
'   If frmForm's designed size is not equal to cChildDesignedXXXX,
'     give the width and height as adefResizable's second & third argument.
Public Sub EvtFormResize(frmForm As Form, ParamArray adefResizable())
    Dim sftTemp As FormSizeType
    Dim intForm As Integer
    Dim blnCustomSize As Boolean
    Dim defTemp As Variant
    Dim ctlTemp As Variant
    Dim i As Integer
    Dim colToBeResized As Collection
    Dim ctlResizable As Control
    Dim defResizable As CtlDefinition
    Dim vntCurHeight, vntCurWidth, vntPreHeight, vntPreWidth
    Dim vntTop, vntLeft, vntWidth, vntHeight
    
    If blnFormSize = True Then
        For intForm = 1 To UBound(aFormSize)
            If aFormSize(intForm).strForm = frmForm.Caption Then
                Exit For
            End If
        Next
        If intForm <= UBound(aFormSize) Then GoTo SkipAdd
    End If
    
    'Resize is functional only in Maximized mode
    If frmForm.WindowState <> vbMaximized Then
        Exit Sub
    End If
    
    If blnFormSize = False Then
        blnFormSize = True
        intForm = 1
    End If
    ReDim Preserve aFormSize(1 To intForm)
    aFormSize(intForm).strForm = frmForm.Caption
    aFormSize(intForm).blnHasControlToResize = True
    
    If TypeName(adefResizable(0)) = "Integer" Then
        aFormSize(intForm).vntDesWidth = adefResizable(0)
        aFormSize(intForm).vntDesHeight = adefResizable(1)
        aFormSize(intForm).blnCustomSize = True
    Else
        aFormSize(intForm).vntDesWidth = cFormDesignedWidth
        aFormSize(intForm).vntDesHeight = cFormDesignedHeight
        aFormSize(intForm).blnCustomSize = False
    End If
    
    aFormSize(intForm).vntWidth = aFormSize(intForm).vntDesWidth
    aFormSize(intForm).vntHeight = aFormSize(intForm).vntDesHeight
    
SkipAdd:
    blnCustomSize = aFormSize(intForm).blnCustomSize

    If blnCustomSize Then
        'Child Form has custom size
        If frmForm.Height < aFormSize(intForm).vntDesHeight Or _
           frmForm.Width < aFormSize(intForm).vntDesWidth Then
            Exit Sub
        End If
    Else
        'Child Form has default size
        If frmForm.Height < (cFormDesignedHeight * 0.7) Or _
           frmForm.Width < (cFormDesignedWidth * 0.7) Then
            'Form is too small to resize
            Exit Sub
        End If
    End If
    
    ' Exit if there is no control to resize
    If aFormSize(intForm).blnHasControlToResize = False Then Exit Sub
    
    vntCurHeight = frmForm.Height
    vntCurWidth = frmForm.Width
    vntPreHeight = aFormSize(intForm).vntHeight
    vntPreWidth = aFormSize(intForm).vntWidth
    
    Set colToBeResized = New Collection
    For i = IIf(blnCustomSize, 2, 0) To UBound(adefResizable)
        Set defTemp = adefResizable(i)
        colToBeResized.Add defTemp, CStr(defTemp.ctlControl.TabIndex)
    Next
    
    'Resize Loop
    i = 0
    Do Until i = colToBeResized.Count
        i = i + 1
        Set defResizable = colToBeResized.Item(i)
        Set ctlResizable = defResizable.ctlControl
        
        Set ctlTemp = defResizable.ctlControl
        vntLeft = 0
        vntTop = 0
        Do Until ctlTemp.Name = frmForm.Name
            vntLeft = vntLeft + ctlTemp.Left
            vntTop = vntTop + ctlTemp.Top
            Set ctlTemp = ctlTemp.Container
        Loop
        
        vntWidth = vntCurWidth - vntPreWidth
        vntHeight = vntCurHeight - vntPreHeight
        
        On Error Resume Next    'Some ctl's width and height are Read-only
        ctlResizable.Left = ctlResizable.Left + vntWidth * defResizable.LeftDiff
        ctlResizable.Width = ctlResizable.Width + vntWidth * defResizable.WidthDiff
        ctlResizable.Top = ctlResizable.Top + vntHeight * defResizable.TopDiff
        ctlResizable.Height = ctlResizable.Height + vntHeight * defResizable.HeightDiff
        On Error GoTo 0
            
        If ctlResizable.Container.Name <> frmForm.Name Then
            Set defTemp = New CtlDefinition
            Set defTemp.ctlControl = ctlResizable.Container
            defTemp.LeftDiff = defResizable.LeftDiff
            defTemp.WidthDiff = defResizable.WidthDiff
            defTemp.TopDiff = defResizable.TopDiff
            defTemp.HeightDiff = defResizable.HeightDiff
            On Error Resume Next
            'If added before, don't add again. And no error.
            colToBeResized.Add defTemp, CStr(ctlResizable.Container.TabIndex)
            On Error GoTo 0
        End If
    Loop
    
PreserveCurrent:
    aFormSize(intForm).vntHeight = frmForm.Height
    aFormSize(intForm).vntWidth = frmForm.Width
End Sub

'Something to help EvtFormResize
Public Sub EvtFormUnload(frmForm As Form, Cancel As Integer)
    Dim intForm As Integer
    
    If blnFormSize = False Then Exit Sub
    
    For intForm = 1 To UBound(aFormSize)
        If aFormSize(intForm).strForm = frmForm.Caption Then
            Exit For
        End If
    Next
    If intForm <= UBound(aFormSize) Then
        aFormSize(intForm).vntWidth = cFormDesignedWidth
        aFormSize(intForm).vntHeight = cFormDesignedHeight
    End If
End Sub

'Used in EvtFormResize
'to define control's position and size
Public Function CtlToResize(ctlResizable As Control, _
                            LeftDiff As Double, _
                            WidthDiff As Double, _
                            TopDiff As Double, _
                            HeightDiff As Double) As CtlDefinition
    Set CtlToResize = New CtlDefinition
    Set CtlToResize.ctlControl = ctlResizable
    CtlToResize.LeftDiff = LeftDiff
    CtlToResize.WidthDiff = WidthDiff
    CtlToResize.TopDiff = TopDiff
    CtlToResize.HeightDiff = HeightDiff
End Function

