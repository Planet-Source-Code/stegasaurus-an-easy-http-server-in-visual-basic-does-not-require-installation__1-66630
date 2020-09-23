Attribute VB_Name = "Functions"
'//Code by michael Billington. This work is in the public domain, do with it as you wish
Private Declare Function BitBlt Lib "gdi32" (ByVal hDestDC As Long, ByVal x As Long, ByVal y As Long, ByVal nWidth As Long, ByVal nHeight As Long, ByVal hSrcDC As Long, ByVal xSrc As Long, ByVal ySrc As Long, ByVal dwRop As Long) As Long
'//Used for the output thing, very useful, moves a box of pixels up by 16
Sub Scrollup(Pic As PictureBox)
BitBlt Pic.hDC, 0, 0, Pic.Width, Pic.Height, Pic.hDC, 0, 16, vbSrcCopy
Pic.Line (0, Pic.Height - 16)-(Pic.Width, Pic.Height), vbWhite, BF
End Sub

'//grabs the content of a file
Function LoadText(FileName As String) As String
  If FileExists(FileName) Then Open FileName For Binary As #1: LoadText = Space(LOF(1)): Get #1, , LoadText: Close #1
End Function

'//Saves text to a file
Sub SaveText(ByVal text As String, FileName As String)
   On Error Resume Next: Open FileName For Output As #1: Print #1, text: Close #1
End Sub
Sub AppendText(ByVal text As String, FileName As String)
   On Error Resume Next: Open FileName For Append As #1: Print #1, text: Close #1
End Sub

'//function to check that a file exists
Function FileExists(FileName As String) As Boolean
   On Error GoTo errorhandle
   If FileLen(FileName) >= 0 Then: FileExists = True: Exit Function
errorhandle:
   FileExists = False
End Function

'very nifty four functions to help with string processing
Function getL(ByVal str As String, ByVal str2 As String) As String
  getL = Left(str, InStr(str, str2) - 1)
  End Function
  Function getR(ByVal str As String, ByVal str2 As String) As String
  getR = Right(str, Len(str) - InStr(str, str2) - Len(str2) + 1)
  End Function
  Function getLrev(ByVal str As String, ByVal str2 As String) As String
  getLrev = Left(str, InStrRev(str, str2) - 1)
  End Function
  Function getRrev(ByVal str As String, ByVal str2 As String) As String
  getRrev = Right(str, Len(str) - InStrRev(str, str2) - Len(str2) + 1)
End Function
Sub Outp(text As String, Optional col As Byte)
FrmMain.Pic.ForeColor = QBColor(col)
Scrollup FrmMain.Pic
FrmMain.Pic.CurrentX = 0
FrmMain.Pic.CurrentY = FrmMain.Pic.Height - 16
FrmMain.Pic.Print text
FrmMain.Pic.Refresh
End Sub

