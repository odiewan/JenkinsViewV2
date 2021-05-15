Public Class findObj
  Private _fndSearchStr As String
  Private _fndSearchLabel As ToolStripLabel
  Private _fndSearchStrLen As Integer
  Private _fndOccurIdxList As New List(Of Integer)()
  Private _fndOccurIdx As Integer
  Private _fndOccurance As Integer
  Private _fndOccurCnt As Integer
  Private _fndDir As Integer
  Private _fndBoxColor As Color = Color.Orange
  Private _fndBoxDisabledColor As Color = Color.Peru

  Private _fndHLColor As Color = Color.Orange

  '================================================================================================
  Public Property searchStr() As String
    Get
      Return _fndSearchStr
    End Get

    Set
      _fndSearchStr = Value
      _fndSearchStrLen = _fndSearchStr.Length
    End Set
  End Property

  '================================================================================================
  Public Property searchLabel() As ToolStripLabel
    Get
      Return _fndSearchLabel
    End Get

    Set
      _fndSearchLabel = Value
    End Set
  End Property

  '================================================================================================
  Public Property searchStrLen() As Integer
    Get
      Return _fndSearchStrLen
    End Get

    Set
      _fndSearchStrLen = Value
    End Set
  End Property

  '================================================================================================
  Public Property idxList() As List(Of Integer)
    Get
      Return _fndOccurIdxList
    End Get
    Set
      _fndOccurIdxList = Value
      _fndOccurCnt = _fndOccurIdxList.Count
    End Set
  End Property

  '================================================================================================
  Public Property idx() As Integer
    Get
      Return _fndOccurIdx
    End Get
    Set
      _fndOccurIdx = Value
    End Set
  End Property

  '================================================================================================
  Public ReadOnly Property occurance()
    Get
      Return _fndOccurance
    End Get
    'Set
    '  _fndOccurance = Value
    'End Set
  End Property

  '================================================================================================
  Public Property count() As Integer
    Get
      Return _fndOccurCnt
    End Get
    Set
      _fndOccurCnt = Value
    End Set
  End Property

  '================================================================================================
  Public Property boxColor() As Color

    Get
      Return _fndBoxColor
    End Get
    Set
      _fndBoxColor = Value
    End Set
  End Property

  '================================================================================================
  Public Property hlColor() As Color
    Get
      Return _fndHLColor
    End Get
    Set
      _fndHLColor = Value
    End Set
  End Property



  '================================================================================================
  Public Sub New(ByRef nSearchStr As String, ByRef nSearchLabel As ToolStripLabel, nColor As Color, nHLColor As Color, nDisableColor As Color)
    fndInit(nSearchStr, nSearchLabel, nColor, nHLColor, nDisableColor)
  End Sub

  '================================================================================================
  Public Sub New()
    Dim tempTSL As New ToolStripLabel
    tempTSL.Text = "Default Label"
    fndInit("", tempTSL, Color.Orange, Color.Orange, Color.Orange)
  End Sub

  '================================================================================================
  Public Sub updateGUI()
    fndGenOccurranceLabel()
  End Sub

  '================================================================================================
  Public Sub fndInit(ByRef nSearchStr As String, ByRef nSearchLabel As ToolStripLabel, nColor As Color, nHLColor As Color, nDisableColor As Color)
    _fndSearchStr = nSearchStr.Trim
    _fndSearchLabel = nSearchLabel
    _fndSearchStrLen = _fndSearchStr.Length + 1
    _fndOccurIdxList = New List(Of Integer)
    _fndOccurIdx = 0
    _fndOccurance = 0
    _fndOccurCnt = 0
    _fndBoxColor = nColor
    _fndHLColor = nHLColor
    _fndBoxDisabledColor = nDisableColor

  End Sub

  '================================================================================================
  Public Sub fndIncDec(nDir As Integer)

    _fndOccurance += nDir

    If _fndOccurance < 0 Then
      _fndOccurance = 0
    ElseIf _fndOccurance >= _fndOccurCnt Then
      _fndOccurance = _fndOccurCnt - 1
    End If
    _fndOccurIdx = _fndOccurIdxList.Item(_fndOccurance)


    updateGUI()
  End Sub

  '================================================================================================
  Public Sub fndAddOccurrance(nIdx As Integer)
    _fndOccurIdxList.Add(nIdx)
    _fndOccurIdx = nIdx
    _fndOccurCnt = _fndOccurIdxList.Count
    _fndSearchStrLen = _fndSearchStr.Length
    updateGUI()

  End Sub

  '================================================================================================
  Sub fndGenOccurranceLabel()
    If _fndOccurCnt > 0 Then
      _fndSearchLabel.Text = (_fndOccurance + 1).ToString + " of " + _fndOccurCnt.ToString
      '_fndSearchLabel.Text = (_fndOccurance + 1).ToString + " of " + _fndOccurCnt.ToString + "<" + _fndOccurIdx.ToString + ">"
    Else
      _fndSearchLabel.Text = "Not found"
    End If


  End Sub
End Class
