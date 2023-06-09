Option Explicit

'[参照HP]
'https://www.feedsoft.net/access/tips/tips11.html

'lpszLongPathのフォルダ名を短い形式のフォルダ名に変換します。
'lpszLongPathのフォルダは、長い形式のフォルダ名でなくてもOK。
'実際に存在しているフォルダでなくてはいけません。
Public Declare Function GetShortPathName Lib "kernel32" Alias "GetShortPathNameA" _
    (ByVal lpszLongPath As String, ByVal lpszShortPath As String, ByVal cchBuffer As Long) As Long

'短いファイル名から長いファイル名を取得する関数
Private Declare Function GetLongPathName Lib "kernel32" Alias "GetLongPathNameA" _
    (ByVal lpszShortPath As String, ByVal lpszLongPath As String, ByVal cchBuffer As Long) As Long

'パス名を短い形式のパス名に変換します。
Public Function MyGetShortPath(sFilePath As String) As String
     '２５６文字確保
    Dim sBuff As String * 256

    MyGetShortPath = Left$(sBuff, GetShortPathName(sFilePath, sBuff, 256))
End Function

'パス名を長い形式のパス名に変換します。
Public Function MyGetLongPath(sFilePath As String) As String
    Dim sBuff As String * 256

    MyGetLongPath = Left$(sBuff, GetLongPathName(sFilePath, sBuff, 256))
End Function


'====
Sub Sample()
    Dim myPath  As String
    Dim FilePath As String
    
    myPath = "C:\Users\ato5f\由恵フォルダ\003_PC\Office\Excel\マクロ"
    
    
    FilePath = MyGetShortPath(myPath)   'NULL除去
    Debug.Print Left(FilePath, InStr(FilePath, vbNullChar) - 1)
    
    FilePath = MyGetLongPath(myPath)   'NULL除去
    Debug.Print Left(FilePath, InStr(FilePath, vbNullChar) - 1)
End Sub
