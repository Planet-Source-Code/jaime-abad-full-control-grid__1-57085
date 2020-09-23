VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "xssEditFlexGrid"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Ext_KEY = "SavedWithClassBuilder6" ,"Yes"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
Option Explicit
Option Compare Text
'****************************************************************
'*  Copyright (C) Jaime Abad 2004 Todos los derechos reservados *
'*                                                              *
'*  Archivo:  xssEditFlexGrid.cls                               *
'*                                                              *
'*  Description:                                                *
'*      Class Module to implement Copy, Paste, Edit,Clear Cells *                                            *
'*                                                              *
'*  History:                                                    *
'*      1 Nov 2004    Jaime Abad      Start Code                *
'*      4 Nov 2004    Jaime Abad      Upload Version 1.0 PSC    *
'*      4 Nov 2004    Jaime Abad      Fix Bug Paste             *
'*                                    Implement Select All      *
'****************************************************************

Const ERR_NOT_ENOUGH_ROW = 400      'Not enough Row to complete the paste
Const ERR_NOT_ENOUGH_COL = 401      'Not enough Col to complete the paste
Const ERR_NOT_SET_TEXTBOX = 402     'First use SetTextBox
Const ERR_NOT_SET_MSFLEXGRID = 403  'First use SetFlexGrid

Private WithEvents txtTextBox As TextBox
Attribute txtTextBox.VB_VarHelpID = -1
Private WithEvents mfxFlexGrid As MSFlexGrid
Attribute mfxFlexGrid.VB_VarHelpID = -1

Event BeforeEdit(OldValue As String)
Event AfterEdit(NewValue As String)

Public AllowEditCells As Boolean
Public AllowKeyTab As Boolean
Public AllowKeyEnter As Boolean
Public AllowKeyUp As Boolean
Public AllowKeyDown As Boolean
Public AllowKeyLeft As Boolean
Public AllowKeyRight As Boolean
Public AllowMoveCols As Boolean
Public AllowKeyEsc As Boolean
Public AllowKeyDel As Boolean

Public Enum EnumSlide
    RightSlide = 0
    DownSlide = 1
    RightDownSlide = 2
End Enum

Public Enum EnumAutocomplete
    ByCol = 0
    ByRow = 1
    None = 2
End Enum

Private intLastRow As Integer
Private intLastCol As Integer
Private intLastColPosition As Integer
Private strLastText As String
Private strAutocompleteText As String

Private Sub Class_Initialize()
    AllowKeyTab = True
    AllowKeyEnter = True
    AllowKeyUp = True
    AllowKeyDown = True
    AllowKeyLeft = True
    AllowKeyRight = True
    AllowMoveCols = True
    AllowKeyEsc = True
    AllowEditCells = True
    AllowKeyDel = True
    intLastColPosition = -1
    Set mfxFlexGrid = Nothing
    Set txtTextBox = Nothing
End Sub

Public Sub SetFlexGrid(MSFlexGridX As Object)
    Set mfxFlexGrid = MSFlexGridX
End Sub

Public Sub SetTextBox(TextBoxX As Object)
    Set txtTextBox = TextBoxX
    With txtTextBox
        .TabStop = False
        .Visible = False
    End With
End Sub

Private Sub Class_Terminate()
    Set txtTextBox = Nothing
    Set mfxFlexGrid = Nothing
End Sub

Private Sub mfxFlexGrid_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    If AllowMoveCols Then
        With mfxFlexGrid
            intLastColPosition = -1
            If .MouseCol < .FixedCols Or .MouseRow >= .FixedRows Then Exit Sub
            intLastColPosition = .MouseCol
            .Drag 1
        End With
    End If
End Sub

Private Sub mfxFlexGrid_DragDrop(Source As Control, x As Single, y As Single)
    With mfxFlexGrid
        If intLastColPosition = -1 Or .MouseCol < .FixedCols Then Exit Sub
        .Redraw = False
        .ColPosition(intLastColPosition) = .MouseCol
        .Redraw = True
    End With
End Sub

Private Sub mfxFlexGrid_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERRHANDLER
    With mfxFlexGrid
        Select Case KeyCode
            Case 9 'Tab
                If AllowKeyTab Then
                    .SetFocus
                    If .Col = .Cols - 1 Then
                        SendKeys ("{Home}")
                        SendKeys ("{Down}")
                    Else
                        SendKeys ("{Right}")
                    End If
                End If
            Case 13 'Enter
                If AllowKeyEnter Then
                    .SetFocus
                    If .Row = .Rows - 1 Then
                        
                    Else
                        SendKeys ("{Down}")
                    End If
                End If
            Case 46 'Del
                If AllowKeyDel Then
                    ClearCells
                End If
            Case 113 'F2
                With mfxFlexGrid
                    If AllowEditCells Then
                        RaiseEvent BeforeEdit(.Text)
                        intLastRow = .Row
                        intLastCol = .Col
                        Call PrepareEdit
                        txtTextBox.Text = .Text
                        txtTextBox.SelStart = Len(.Text)
                        txtTextBox.Visible = True
                        txtTextBox.SetFocus
                    End If
                End With
        End Select
    End With
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_TEXTBOX + vbObjectError, "xssEditFlexGrid", "First use SetTextBox "
    End If
End Sub

Private Sub mfxFlexGrid_KeyPress(KeyAscii As Integer)
    On Error GoTo ERRHANDLER
    Select Case KeyAscii
        Case Is = 9, 13, 38, 40, 27, 37, 39, 46
            'Nothing
        Case Else
            With mfxFlexGrid
                If AllowEditCells Then
                    strLastText = .Text
                    RaiseEvent BeforeEdit(.Text)
                    intLastRow = .Row
                    intLastCol = .Col
                    Call PrepareEdit
                    txtTextBox.Text = Chr(KeyAscii)
                    txtTextBox.SelStart = 1
                    txtTextBox.Visible = True
                    txtTextBox.SetFocus
                End If
            End With
    End Select
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_TEXTBOX + vbObjectError, "xssEditFlexGrid", "First use SetTextBox "
    End If
End Sub

Private Sub mfxFlexGrid_Scroll()
    mfxFlexGrid.SetFocus
End Sub

Private Sub txtTextBox_KeyDown(KeyCode As Integer, Shift As Integer)
    On Error GoTo ERRHANDLER
    With mfxFlexGrid
        Select Case KeyCode
            Case 9 'Tab
                If AllowKeyTab Then
                    .SetFocus
                    SendKeys ("{Right}")
                End If
            Case 39 'Right
                If AllowKeyRight Then
                    .SetFocus
                    SendKeys ("{Right}")
                End If
            Case 13 'Enter
                If AllowKeyEnter Then
                    .SetFocus
                    SendKeys ("{Down}")
                End If
            Case 40 'Down
                If AllowKeyDown Then
                    .SetFocus
                    SendKeys ("{Down}")
                End If
            Case 38 'Up
                If AllowKeyUp Then
                    .SetFocus
                    SendKeys ("{Up}")
                End If
            Case 37 'Left
                If AllowKeyLeft Then
                    .SetFocus
                    SendKeys ("{Left}")
                End If
            Case 27 'Esc
                If AllowKeyEsc Then
                    txtTextBox.Text = strLastText
                    .SetFocus
                End If
        End Select
    End With
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_MSFLEXGRID + vbObjectError, "xssEditFlexGrid", "First use SetFlexGrid "
    End If
End Sub

Private Sub txtTextBox_LostFocus()
    On Error GoTo ERRHANDLER
    With mfxFlexGrid
        .TextMatrix(intLastRow, intLastCol) = txtTextBox.Text
        .SetFocus
        txtTextBox.Visible = False
        RaiseEvent AfterEdit(.Text)
    End With
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_MSFLEXGRID + vbObjectError, "xssEditFlexGrid", "First use SetFlexGrid "
    End If
End Sub

Public Sub ClearCells()
    On Error GoTo ERRHANDLER
    Dim Row As Integer, Col As Integer, intStep1 As Integer, intStep2 As Integer
    With mfxFlexGrid
        intStep1 = IIf(.Row <= .RowSel, 1, -1)
        .Redraw = False
        For Row = .Row To .RowSel Step intStep1
            intStep2 = IIf(.Col <= .ColSel, 1, -1)
            For Col = .Col To .ColSel Step intStep2
                .TextMatrix(Row, Col) = ""
                DoEvents
            Next
            DoEvents
        Next
        .Redraw = True
    End With
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_MSFLEXGRID + vbObjectError, "xssEditFlexGrid", "First use SetFlexGrid "
    End If
End Sub

Public Sub StarEdit()
    On Error GoTo ERRHANDLER
    With mfxFlexGrid
        If AllowEditCells Then
            RaiseEvent BeforeEdit(.Text)
            intLastRow = .Row
            intLastCol = .Col
            Call PrepareEdit
            txtTextBox.Text = .Text
            txtTextBox.SelStart = Len(.Text)
            txtTextBox.Visible = True
            txtTextBox.SetFocus
        End If
    End With
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_TEXTBOX + vbObjectError, "xssEditFlexGrid (StarEdit)", "First use SetTextBox "
    End If
End Sub

Private Sub PrepareEdit()
    With mfxFlexGrid
        Select Case .ColAlignment(mfxFlexGrid.Col)
            Case Is = 0, 1, 2, 9
                txtTextBox.Alignment = 0
            Case Is = 3, 4, 5
                txtTextBox.Alignment = 2
            Case Is = 6, 7, 8
                txtTextBox.Alignment = 1
        End Select
        txtTextBox.FontBold = .CellFontBold
        txtTextBox.FontItalic = .CellFontItalic
        txtTextBox.FontName = .CellFontName
        txtTextBox.FontSize = .CellFontSize
        txtTextBox.FontStrikethru = .CellFontStrikeThrough
        txtTextBox.FontUnderline = .CellFontUnderline
        txtTextBox.ForeColor = .CellForeColor
        txtTextBox.Move .Left + .CellLeft, .Top + .CellTop, .CellWidth - 2, .CellHeight - 2
    End With
End Sub

'This method insert to clipboard
Public Sub Copy()
    On Error GoTo ERRHANDLER
    Dim Row As Integer, Col As Integer
    Dim StartRow As Integer, EndRow As Integer
    Dim StartCol As Integer, EndCol As Integer
    Dim strCopy As String
    strCopy = ""
    With mfxFlexGrid
        StartRow = IIf(.Row <= .RowSel, .Row, .RowSel)
        EndRow = IIf(.Row <= .RowSel, .RowSel, .Row)
        StartCol = IIf(.Col <= .ColSel, .Col, .ColSel)
        EndCol = IIf(.Col <= .ColSel, .ColSel, .Col)
        .Redraw = False
        For Row = StartRow To EndRow
            If Row <> StartRow Then
                strCopy = strCopy & vbCrLf
            End If
            For Col = StartCol To EndCol
                If Col <> StartCol Then
                    strCopy = strCopy & vbTab
                End If
                strCopy = strCopy & .TextMatrix(Row, Col)
                DoEvents
            Next
            DoEvents
        Next
        .Redraw = True
    End With
    Clipboard.Clear
    Clipboard.SetText strCopy
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_MSFLEXGRID + vbObjectError, "xssEditFlexGrid", "First use SetFlexGrid "
    End If
End Sub

'This function Paste from Clipboard
Public Function Paste(Optional AllowError = False) As Boolean
    On Error GoTo ERRHANDLER
    Dim strPaste() As String, strData() As String
    Dim Row As Integer, Col As Integer
    Paste = False
    If Clipboard.GetFormat(vbCFText) Then
        strPaste = Split(Clipboard.GetText, vbCrLf)
        With mfxFlexGrid
            Row = 0
            .Redraw = False
            Do While Row <= UBound(strPaste)
                If (Row < .Rows - .Row) Then
                    strData = Split(strPaste(Row), vbTab)
                    Col = 0
                    Do While Col <= UBound(strData)
                        If (Col < .Cols - .Col) Then
                            .TextMatrix(Row + .Row, Col + .Col) = strData(Col)
                        Else
                            If AllowError Then
                                Err.Raise ERR_NOT_ENOUGH_COL + vbObjectError, "xssEditFlexGrid (Paste)", "Not Enough Col to complete the paste operation"
                            End If
                            Exit Do
                        End If
                        Col = Col + 1
                        DoEvents
                    Loop
                Else
                    If AllowError Then
                        Err.Raise ERR_NOT_ENOUGH_ROW + vbObjectError, "xssEditFlexGrid (Paste)", "Not Enough Row to complete the paste operation"
                    End If
                    Exit Do
                End If
                Row = Row + 1
                DoEvents
            Loop
            .Redraw = True
        End With
        Paste = True
    End If
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_MSFLEXGRID + vbObjectError, "xssEditFlexGrid", "First use SetFlexGrid "
    ' Star Fixed 10/4/2004
    ' now raise error, now can send an error if the cols or the rows are enough
    ElseIf Err.Number = ERR_NOT_ENOUGH_ROW + vbObjectError Then
        Err.Raise Err.Number, "xssEditFlexGrid (Paste)", "Not Enough Row to complete the paste operation"
    ElseIf Err.Number = ERR_NOT_ENOUGH_COL + vbObjectError Then
        Err.Raise Err.Number, "xssEditFlexGrid (Paste)", "Not Enough Col to complete the paste operation"
    'end of fixed
    End If
End Function

Public Function Insert(Optional Slide As EnumSlide = 2) As Boolean
    On Error GoTo ERRHANDLER
    Dim intRowBlank As Integer, intRowsClip As Integer, intSetRow As Integer
    Dim intColBlank As Integer, intColsClip As Integer, intSetCol As Integer
    Insert = False
    If Clipboard.GetFormat(vbCFText) Then
        intRowsClip = RowsInClipboard
        intColsClip = ColsInClipboard
        With mfxFlexGrid
            .Redraw = False
            Select Case Slide
                Case 0 'Right
                    If .Rows < intRowsClip + .Row Then
                        .Rows = intRowsClip + .Row
                    End If
                    For intSetRow = 0 To intRowsClip - 1
                        SlideRight intColsClip, .Row + intSetRow
                        DoEvents
                    Next
                    Paste
                Case 1 'Down
                    If .Cols < intColsClip + .Col Then
                        .Cols = intColsClip + .Col
                    End If
                    For intSetCol = 0 To intColsClip - 1
                        SlideDown intRowsClip, .Col + intSetCol
                        DoEvents
                    Next
                    Paste
                Case 2 'Both
                    If .Rows < intRowsClip + .Row Then
                        .Rows = intRowsClip + .Row
                    End If
                    If .Cols < intColsClip + .Col Then
                        .Cols = intColsClip + .Col
                    End If
                    For intSetRow = 0 To intRowsClip - 1
                        SlideRight intColsClip, .Row + intSetRow
                        DoEvents
                    Next
                    For intSetCol = 0 To intColsClip - 1
                        SlideDown intRowsClip, .Col + intSetCol
                        DoEvents
                    Next
                    Paste
            End Select
            .Redraw = True
        End With
        Insert = True
    End If
ERRHANDLER:
    If Err.Number = 91 Then
        Err.Raise ERR_NOT_SET_MSFLEXGRID + vbObjectError, "xssEditFlexGrid", "First use SetFlexGrid "
    End If
End Function

Public Function RowsInClipboard() As Integer
    Dim strPaste() As String
    If Clipboard.GetFormat(vbCFText) Then
        strPaste = Split(Clipboard.GetText, vbCrLf)
        RowsInClipboard = UBound(strPaste) + 1
    Else
        RowsInClipboard = 0
    End If
End Function

Public Function ColsInClipboard() As Integer
    Dim strPaste() As String, strData() As String
    If Clipboard.GetFormat(vbCFText) Then
        strPaste = Split(Clipboard.GetText, vbCrLf)
        strData = Split(strPaste(0), vbTab)
        ColsInClipboard = UBound(strData) + 1
    Else
        ColsInClipboard = 0
    End If
End Function

Private Sub SlideRight(Cols As Integer, Optional Row)
    Dim intColBlank As Integer, intColsClip As Integer, intSetCol As Integer
    With mfxFlexGrid
        If IsMissing(Row) Then
            Row = .Row
        End If
        intSetCol = .Cols - 1
        intColBlank = 0
        intColsClip = Cols
        .Redraw = False
        Do While intSetCol >= .Col And intColBlank <= intColsClip
            If .TextMatrix(Row, intSetCol) <> "" Then
                Exit Do
            End If
            intSetCol = intSetCol - 1
            intColBlank = intColBlank + 1
            DoEvents
        Loop
        .Cols = .Cols + IIf(intColBlank >= intColsClip, 0, intColsClip - intColBlank)
        For intSetCol = .Cols - 1 To .Col + intColsClip Step -1
            .TextMatrix(Row, intSetCol) = .TextMatrix(Row, intSetCol - intColsClip)
        Next
        .Redraw = True
    End With
End Sub

Private Sub SlideDown(Rows As Integer, Optional Col)
    Dim intRowBlank As Integer, intRowsClip As Integer, intSetRow As Integer
    With mfxFlexGrid
        If IsMissing(Col) Then
            Col = .Col
        End If
        intSetRow = .Rows - 1
        intRowBlank = 0
        intRowsClip = Rows
        .Redraw = False
        Do While intSetRow >= .Row And intRowBlank <= intRowsClip
            If .TextMatrix(intSetRow, Col) <> "" Then
                Exit Do
            End If
            intSetRow = intSetRow - 1
            intRowBlank = intRowBlank + 1
            DoEvents
        Loop
        .Rows = .Rows + IIf(intRowBlank >= intRowsClip, 0, intRowsClip - intRowBlank)
        For intSetRow = .Rows - 1 To .Row + intRowsClip Step -1
            .TextMatrix(intSetRow, Col) = .TextMatrix(intSetRow - intRowsClip, Col)
        Next
        .Redraw = True
    End With
End Sub

Public Sub SelectAll()
    With mfxFlexGrid
        .Row = .FixedRows
        .Col = .FixedCols
        .RowSel = .Rows - 1
        .ColSel = .Cols - 1
    End With
End Sub
