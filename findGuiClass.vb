Imports System.Security

Public Class findGuiClass

  Private _Name As String
  Private _findObj As findObj
  Private _hlColor As Color
  Private _occColor As Color
  Private _disableColor As Color

  Private _enabled As Boolean
  Private _caseSensitive As Boolean
  Private _wholeWord As Boolean

  Public fgcLabel As ToolStripLabel
  Public WithEvents fgcTextBox As ToolStripTextBox
  Public WithEvents fgcBtnFind As ToolStripButton
  Public WithEvents fgcBtnWholeWord As ToolStripButton
  Public WithEvents fgcBtnCaseSensitve As ToolStripButton
  Public WithEvents fgcBtnFindUp As ToolStripButton
  Public WithEvents fgcBtnFindDwn As ToolStripButton
  Public WithEvents fgcRTBX As RichTextBox

  '================================================================================================
  Public Property caseSensitive() As Boolean
    Get
      Return _caseSensitive
    End Get

    Set
      _caseSensitive = Value
    End Set
  End Property

  '----------------------------------------------------------------------------
  Public Sub New(nName As String,
                 nLabel As ToolStripLabel,
                 nTBX As ToolStripTextBox,
                 nFndBtn As ToolStripButton,
                 nWholeBtn As ToolStripButton,
                 nCsBtn As ToolStripButton,
                 nUpBtn As ToolStripButton,
                 nDwnBtn As ToolStripButton,
                 nRtb As RichTextBox,
                 nHlColor As Color,
                  nColor As Color,
                 nDisbleColor As Color)

    fgcInit(nName,
            nLabel,
            nTBX,
            nFndBtn,
            nWholeBtn,
            nCsBtn,
            nUpBtn,
            nDwnBtn,
            nRtb,
            nHlColor,
            nColor,
            nDisbleColor)
  End Sub

  '----------------------------------------------------------------------------
  Public Sub fgcInit(nName As String,
                     nLabel As ToolStripLabel,
                     nTBX As ToolStripTextBox,
                     nFndBtn As ToolStripButton,
                     nWholeBtn As ToolStripButton,
                     nCsBtn As ToolStripButton,
                     nUpBtn As ToolStripButton,
                     nDwnBtn As ToolStripButton,
                     nRtb As RichTextBox,
                     nHlColor As Color,
                      nColor As Color,
                      nDisbleColor As Color)
    _Name = nName
    _findObj = New findObj()
    _hlColor = nHlColor
    _occColor = nColor
    _disableColor = nDisbleColor
    _enabled = False
    fgcLabel = nLabel
    fgcTextBox = nTBX
    fgcTextBox.BackColor = _occColor


    fgcBtnFind = nFndBtn
    fgcBtnWholeWord = nWholeBtn
    fgcBtnCaseSensitve = nCsBtn
    fgcBtnFindUp = nUpBtn
    fgcBtnFindDwn = nDwnBtn
    fgcRTBX = nRtb
    'fgcBtnCaseSensitve.BackColor = Color.Yellow


  End Sub
  Public Sub findEnable(nEn As Boolean)
    _enabled = nEn
    If _enabled Then
      fgcTextBox.Enabled = True
      'fgcLabel.Enabled = True
      fgcBtnFind.Enabled = True
      'fgcBtnFindUp.Enabled = True
      'fgcBtnFindDwn.Enabled = True
      fgcBtnCaseSensitve.Enabled = True
      fgcBtnWholeWord.Enabled = True
      fgcTextBox.BackColor = _occColor

    Else
      fgcTextBox.Enabled = False
      fgcLabel.Enabled = False
      fgcBtnFind.Enabled = False
      fgcBtnFindUp.Enabled = False
      fgcBtnFindDwn.Enabled = False
      fgcBtnCaseSensitve.Enabled = False
      fgcBtnWholeWord.Enabled = False
      fgcTextBox.BackColor = _disableColor

    End If
  End Sub

  '============================================================================
  Private Sub updateGUI()
    If _findObj.count = 0 Then
      fgcLabel.Enabled = False
      fgcBtnFindUp.Enabled = False
      fgcBtnFindDwn.Enabled = False

    Else
      If _findObj.count = 1 Then
        fgcLabel.Enabled = True
        fgcBtnFindUp.Enabled = False
        fgcBtnFindDwn.Enabled = False

      Else
        fgcLabel.Enabled = True
        fgcBtnFindUp.Enabled = True
        fgcBtnFindDwn.Enabled = True
      End If

      fgcRTBX.[Select](0, fgcRTBX.TextLength)
      fgcRTBX.SelectionFont = New Font(fgcRTBX.Font, FontStyle.Regular)
      fgcRTBX.SelectionBackColor = SystemColors.Window

      '---Select and highlight each occurrance
      For j As Integer = 0 To _findObj.count - 1
        markOccurance(j, True)
        If j = _findObj.occurance Then
          highlightOccurance(j, True)
        End If
      Next

    End If
    _findObj.fndGenOccurranceLabel()

  End Sub

  '============================================================================
  Private Sub tsbtFind_Click(sender As Object, e As KeyEventArgs) Handles fgcTextBox.KeyDown
    If e.KeyCode = Keys.Enter Then
      find()
    End If

  End Sub


  '============================================================================
  Private Sub tsbtFind_Click(sender As Object, e As EventArgs) Handles fgcBtnFind.Click
    find()
  End Sub

  '============================================================================
  Private Sub tsbtFindUp_Click(sender As Object, e As EventArgs) Handles fgcBtnFindUp.Click
    _findObj.fndIncDec(-1)
    updateGUI()
  End Sub

  '============================================================================
  Private Sub tsbtFindDwn_Click(sender As Object, e As EventArgs) Handles fgcBtnFindDwn.Click
    _findObj.fndIncDec(1)
    updateGUI()
  End Sub

  '============================================================================
  Private Sub tsbtCaseSensitve_Click(sender As Object, e As EventArgs) Handles fgcBtnCaseSensitve.Click
    _caseSensitive = fgcBtnCaseSensitve.Checked
    If _caseSensitive Then
      fgcBtnCaseSensitve.BackColor = Color.Yellow

    Else
      fgcBtnCaseSensitve.BackColor = System.Drawing.Color.Transparent

    End If

    find()

  End Sub


  '============================================================================
  Public Sub find()

    Dim searchStr As String

    Dim tempFindObj As New findObj

    searchStr = fgcTextBox.Text
    _findObj.fndInit(searchStr, fgcLabel, _occColor, _hlColor, _disableColor)



    '----Loop through all lines in the results richtextbox (RTB) for a match
    '    if a match is found, add it's absolute pos to the current find obj's occurrance list, 
    '    calculate the string length, set the ocurrance to 0, and generate the occurrance indicator string
    If _findObj.searchStr <> "" And fgcRTBX.Text <> "" Then
      Dim idx As Integer
      While idx <> -1
        'Look for the 1st occurrance of the search string in the text body
        If _caseSensitive Then
          idx = fgcRTBX.Find(_findObj.searchStr, idx, RichTextBoxFinds.None)
        Else
          idx = fgcRTBX.Find(_findObj.searchStr, idx, RichTextBoxFinds.MatchCase)
        End If

        If idx <> -1 Then
          'If the word is found add the index to the list
          '---Since we move the start position of the search to position right 
          '   after the previous occurrance (index of prev occurrance + its length)
          '   we need to calculate the occurrance absolute position in the searched string

          _findObj.fndAddOccurrance(idx)
          idx += tempFindObj.searchStrLen

        End If
      End While

      ''---Select and highlight each occurrance
      'For j As Integer = 0 To _findObj.count - 1

      '  markOccurance(j, True)
      'Next
    End If
    updateGUI()
  End Sub

  Private Sub highlightOccurance(nIdx As Integer, nOn As Boolean)
    fgcRTBX.[Select](_findObj.idxList(nIdx), _findObj.searchStrLen)
    If nOn Then
      fgcRTBX.SelectionFont = New Font(fgcRTBX.Font, FontStyle.Bold)
      fgcRTBX.SelectionBackColor = _findObj.boxColor
      fgcRTBX.ScrollToCaret()
    Else
      fgcRTBX.SelectionFont = New Font(fgcRTBX.Font, FontStyle.Regular)
      fgcRTBX.SelectionBackColor = _findObj.boxColor
      'fgcRTBX.SelectionBackColor = SystemColors.Window

    End If

  End Sub
  Private Sub markOccurance(nIdx As Integer, nOn As Boolean)
    Dim tmpIdx As Integer = _findObj.idxList(nIdx)
    fgcRTBX.[Select](tmpIdx, _findObj.searchStrLen)
    If nOn Then
      fgcRTBX.SelectionFont = New Font(fgcRTBX.Font, FontStyle.Regular)
      fgcRTBX.SelectionBackColor = _findObj.boxColor
    Else
      fgcRTBX.SelectionFont = New Font(fgcRTBX.Font, FontStyle.Regular)
      fgcRTBX.SelectionBackColor = SystemColors.Window

    End If

  End Sub

End Class
