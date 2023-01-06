'定数
Public Const EXCEL_EXTENSION = "xlsx"
Public Const PROCCESS_FAILED = "処理に失敗しました。"
Public Const CONFIRM = "確認"


'【概要】特定のディレクトリの特定の拡張子のファイル群を取得する
Public Function GetFilePaths(ByVal strDirectoryPath As String, _
                        ByVal strExtensionName As String) As Variant
On Error GoTo GetFilePaths_Err
    
    Dim lngArrIdx As Long
    Dim arrRet() As Variant
    Dim objFso As FileSystemObject
    Dim objFile As File
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    With objFso
        'ファイルの数だけ繰り返す
        For Each objFile In .GetFolder(strDirectoryPath).Files
            '拡張子が指定の者だった場合
            If .GetExtensionName(objFile.Name) = strExtensionName Then
                '配列再宣言
                ReDim Preserve arrRet(lngArrIdx)
                '配列に格納
                arrRet(lngArrIdx) = objFile.Path
                '配列の要素番号を1つ進める
                lngArrIdx = lngArrIdx + 1
            End If
        Next objFile
    End With
        
    GetFilePaths = arrRet
    
GetFilePaths_Err:

GetFilePaths_Exit:
    Set objFso = Nothing
    Set objFile = Nothing
End Function

