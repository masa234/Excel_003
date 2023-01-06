    

'【概要】Excelファイル（複数）をCSVとして出力する
Public Function ExcelFilesToCSV(ByVal arrExcelFilePaths, _
                            ByVal strCSVFilePath As String) As Boolean
On Error GoTo ExcelFilesToCSV_Err

    ExcelFilesToCSV = False
    
    Dim lngFreeFile As Long
    Dim lngLastRow As Long
    Dim lngCurrentRow As Long
    Dim strValue As String
    Dim objWb As Excel.Workbook
    Dim objWs As Excel.Worksheet
    
    'フリーファイル
    lngFreeFile = FreeFile
    
    'Fsoを呼び出す
    Set objFso = New FileSystemObject
    
    'CSVファイルを開く（書き込みモード）
    Open strCSVFilePath For Output As #1

    '配列の最初から終端まで繰り返す
    For lngArrIdx = 0 To UBound(arrExcelFilePaths)
        'ファイルパス
        strFilePath = arrExcelFilePaths(lngArrIdx)
        'Excelファイルを開く
        Workbooks.Open strFilePath
        Set objWb = ActiveWorkbook
        'シートの数だけ繰り返す
        For Each objWs In objWb.Worksheets
            '最終行を取得する
            lngLastRow = objWb.Worksheets(objWs.Name).Cells(1, 1).End(xlDown).Row
            '最終行まで繰り返す
            For lngCurrentRow = 1 To lngLastRows
                'Excelの値
                strValue = objWb.Worksheets(objWs.Name).Cells(lngCurrentRow, 1).Value
                'CSVに1行出力
                Print #lngFreeFile, strValue
                '行カウントを1つ進める
                lngCurrentRow = lngCurrentRow + 1
            Next lngCurrentRow
        Next objWs
    Next lngArrIdx
            
    ExcelFilesToCSV = True
    
ExcelFilesToCSV_Err:

ExcelFilesToCSV_Exit:
    'CSVファイルを閉じる
    Close #lngFreeFile
    Set objWb = Nothing
    Set objWs = Nothing
End Function
