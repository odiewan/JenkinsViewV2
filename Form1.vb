Public Class Form1

  Dim frmInputSize As Integer
  Dim frmOutputSize As Integer


  Dim fndObj00 As findGuiClass
  Dim fndObj01 As findGuiClass
  Dim fndObj02 As findGuiClass


  Dim hlColor00 As Color = Color.Red
  Dim hlColor01 As Color = Color.Green
  Dim hlColor02 As Color = Color.Blue
  Dim Color00 As Color = Color.Salmon
  Dim Color01 As Color = Color.LightGreen
  Dim Color02 As Color = Color.LightBlue
  Dim disableColor00 As Color = Color.Peru
  Dim disableColor01 As Color = Color.Olive
  Dim disableColor02 As Color = Color.Firebrick



  '============================================================================
  Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
    frmInputSize = 0
    frmOutputSize = 0

    fndObj00 = New findGuiClass("find00",
                                tslbFind00,
                                tstbFind00,
                                tsbtFind00,
                                tsbtWholeWord00,
                                tsbtCaseSensitive00,
                                tsbtFind00Up,
                                tsbtFind00Dwn,
                                rtbxFind,
                                hlColor00,
                                Color00,
disableColor00)

    fndObj01 = New findGuiClass("find01",
                                tslbFind01,
                                tstbFind01,
                                tsbtFind01,
                                tsbtWholeWord01,
                                tsbtCaseSensitive01,
                                tsbtFind01Up,
                                tsbtFind01Dwn,
                                rtbxFind,
                                hlColor01,
                                Color01,
                                disableColor01)

    fndObj02 = New findGuiClass("find02",
                                tslbFind02,
                                tstbFind02,
                                tsbtFind02,
                                tsbtWholeWord02,
                                tsbtCaseSensitive02,
                                tsbtFind02Up,
                                tsbtFind02Dwn,
                                rtbxFind,
                                hlColor02,
                                Color02,
                                disableColor02)


    updateGUI()
  End Sub

  '============================================================================
  Private Sub updateGUI()

    fndObj00.findEnable(lbxOutput.Items.Count <> 0)
    fndObj01.findEnable(lbxOutput.Items.Count <> 0)
    fndObj02.findEnable(lbxOutput.Items.Count <> 0)

    If lbxOutput.Items.Count = 0 Then
      cmsOutput.Enabled = False
      lbxOutput.Enabled = False
    Else
      cmsOutput.Enabled = True
      lbxOutput.Enabled = True
    End If


    tsslInputStatus.Text = "Input size:" + frmInputSize.ToString


  End Sub

  '============================================================================
  Private Sub tbxInput_DragEnter(sender As Object, e As DragEventArgs) Handles tbxElementIn.DragEnter, tbxInput.DragEnter
    Debug.WriteLine("tbxInput_DragEnter")

    If e.Data.GetDataPresent(DataFormats.Text) Then
      e.Effect = DragDropEffects.Move
      tbxInput.Clear()
      tbxInput.BackColor = Color.Orange
    ElseIf e.Data.GetDataPresent(DataFormats.FileDrop) Then
      e.Effect = DragDropEffects.Copy
    Else
      e.Effect = DragDropEffects.None
    End If
    'updateGUI()
  End Sub


  '============================================================================
  Private Sub tbxInput_DragDrop(sender As Object, e As DragEventArgs) Handles tbxElementIn.DragDrop, tbxInput.DragDrop
    Debug.WriteLine("tbxSearchTerm_DragDrop")
    Debug.WriteLine("AllowedEffect:" & e.AllowedEffect)

    parseInputTbx(e.Data.GetData(DataFormats.Text))

  End Sub


  '============================================================================
  Private Function insertNewLineTabAfter(nHaystack As String, nNeedle As String) As String

    Return nHaystack.Replace(nNeedle, nNeedle + vbNewLine + vbTab)

  End Function
  '============================================================================
  Private Function insertNewLineTabBefore(nHaystack As String, nNeedle As String)
    Return nHaystack.Replace(nNeedle, vbNewLine + vbTab + nNeedle)

  End Function

  '============================================================================
  Private Sub parseInputTbx(ByRef nInputStr As String)


    tbxInput.Clear()
    tbxInput.Text = nInputStr
    tbxInput.BackColor = SystemColors.Window

    lbxOutput.Items.Clear()
    lbxOutput.BackColor = SystemColors.Window

    rtbxFind.Clear()
    rtbxFind.BackColor = SystemColors.Window

    frmInputSize = nInputStr.Length()

    Dim tmpStr As String() = Split(nInputStr, "\n")
    'tmpStr = Split(tmpStr, "")

    lbxOutput.Items.Clear()
    lbxOutput.BackColor = Color.Orange

    rtbxFind.Clear()
    rtbxFind.BackColor = Color.Orange

    frmOutputSize = tmpStr.Length()

    For Each line As String In tmpStr
      line = line.Replace("\\\n", "")
      line = line.Replace("\\n", "")
      line = line.Replace(vbCrLf, "")
      line = line.Replace("\r", "")
      line = line.Replace("\'", "'")
      'line = line.Replace("{'", vbNewLine + vbTab + "{'")
      line = insertNewLineTabAfter(line, "{'")
      line = insertNewLineTabAfter(line, "},")
      line = line.Replace(", '", "," + vbNewLine + vbTab + "'")
      line = line.Replace(",'", "," + vbNewLine + vbTab + "'")


      line = line.Replace("{\'", vbNewLine + vbTab + "{'")
      line = line.Replace("\', \'", "'," + vbNewLine + vbTab + "'")


      line = line.Replace("', {'", "'," + vbNewLine + vbTab + "{'")
      line = line.Replace("': [{'", "':" + vbNewLine + vbTab + "[{'")


      'line = line.Replace("{\'", vbNewLine + vbTab + "{\'")
      'line = line.Replace("\', \'", "\'," + vbNewLine + vbTab + "\'")
      lbxOutput.Items.Add(line)
      rtbxFind.AppendText(line + vbNewLine)

    Next
    lbxOutput.BackColor = SystemColors.Window
    rtbxFind.BackColor = SystemColors.Window
    updateGUI()

  End Sub





  '============================================================================
  Private Sub ClearToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearToolStripMenuItem.Click
    tbxInput.Clear()
    updateGUI()
  End Sub

  '============================================================================
  Private Sub ClearOutputToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ClearOutputToolStripMenuItem.Click
    lbxOutput.Items.Clear()
    updateGUI()
  End Sub

  '============================================================================
  Private Sub RefreshToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles RefreshToolStripMenuItem.Click
    parseInputTbx(tbxInput.Text)

  End Sub


  '============================================================================
  Private Sub ToTopToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToTopToolStripMenuItem.Click
    lbxOutput.SelectedIndex = lbxOutput.TopIndex
  End Sub

  '============================================================================
  Private Sub ToBottomToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles ToBottomToolStripMenuItem.Click
    lbxOutput.SelectedIndex = lbxOutput.Items.Count - 1
  End Sub


  Private Sub tbxInput_DoubleClick(sender As Object, e As EventArgs)
    parseInputTbx(tbxInput.Text)
  End Sub

  Private Sub tsbtParse_Click_1(sender As Object, e As EventArgs) Handles tsbtParse.Click
    parseInputTbx(tbxInput.Text)
  End Sub

  Private Sub tsbtClear_Click(sender As Object, e As EventArgs) Handles tsbtClear.Click
    tbxInput.Clear()
    updateGUI()
  End Sub

  Private Sub tsbtCaseSensitive00_Click(sender As Object, e As EventArgs) Handles tsbtCaseSensitive00.Click

  End Sub

  Private Sub tbxInput_TextChanged(sender As Object, e As EventArgs) Handles tbxInput.TextChanged

  End Sub
End Class
