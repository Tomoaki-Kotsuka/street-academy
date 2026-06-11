Attribute VB_Name = "Module_Aggregate"
' ============================================================
' 47 都道府県分の sales_*.xlsx を読み込み、
' 商品別の売上合計を集計して経過時間を表示する VBA マクロ
'
' 使い方:
'   1. 任意の場所に新しい Excel ブックを作成 → 拡張子 .xlsm で保存
'   2. Alt + F11 で VBE を開く
'   3. メニュー: ファイル → ファイルのインポート → このファイル
'   4. F5 でマクロ実行
'   5. ダイアログに sales_2026-01 フォルダのフルパスを入力
' ============================================================
Option Explicit

Public Sub AggregateMonthlySales()
    Dim startTime As Double
    startTime = Timer

    ' ----- 対象フォルダの取得 -----
    Dim targetFolder As String
    targetFolder = InputBox( _
        "売上 Excel が入っているフォルダの絶対パスを入れてください" & vbCrLf & vbCrLf & _
        "例 (Windows): C:\Users\YourName\Desktop\python-data-basics-local\data\capstone\sales_2026-01" & vbCrLf & _
        "例 (Mac)    : /Users/yourname/python-data-basics-local/data/capstone/sales_2026-01", _
        "VBA 速度ベンチマーク")
    If targetFolder = "" Then Exit Sub

    Dim sep As String
    sep = Application.PathSeparator   ' Windows: \ , Mac: /
    If Right(targetFolder, 1) <> sep Then targetFolder = targetFolder & sep

    ' ----- ファイル一覧の取得 -----
    Dim fileName As String
    fileName = Dir(targetFolder & "sales_*.xlsx")
    If fileName = "" Then
        MsgBox "対象ファイルが見つかりませんでした:" & vbCrLf & targetFolder, vbExclamation
        Exit Sub
    End If

    ' ----- 集計用 Dictionary -----
    Dim totals As Object
    Set totals = CreateObject("Scripting.Dictionary")

    Dim fileCount As Long: fileCount = 0
    Dim rowCount As Long: rowCount = 0

    ' ----- 高速化のおまじない (公平に「最適化済み VBA」として実行) -----
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.DisplayAlerts = False
    Application.EnableEvents = False

    ' ----- 各ファイルを開いて集計 -----
    Do While fileName <> ""
        Dim wb As Workbook
        Set wb = Workbooks.Open(targetFolder & fileName, ReadOnly:=True, UpdateLinks:=False)

        Dim ws As Worksheet
        Set ws = wb.Worksheets("売上")

        Dim lastRow As Long
        lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

        ' 1 回でまとめて配列に読み込んでループ (個別 Cells アクセスより十分速い)
        Dim arr As Variant
        arr = ws.Range("A2:F" & lastRow).Value

        Dim i As Long
        For i = 1 To UBound(arr, 1)
            Dim productName As String
            Dim amount As Double
            productName = CStr(arr(i, 3))   ' C 列: product_name
            amount = CDbl(arr(i, 6))         ' F 列: amount

            If totals.Exists(productName) Then
                totals(productName) = totals(productName) + amount
            Else
                totals.Add productName, amount
            End If
            rowCount = rowCount + 1
        Next i

        wb.Close SaveChanges:=False
        Set wb = Nothing

        fileCount = fileCount + 1
        fileName = Dir
    Loop

    ' ----- 結果を ThisWorkbook の "VBA結果" シートに書き出す -----
    Dim resultSheet As Worksheet
    On Error Resume Next
    Set resultSheet = ThisWorkbook.Worksheets("VBA結果")
    On Error GoTo 0
    If resultSheet Is Nothing Then
        Set resultSheet = ThisWorkbook.Worksheets.Add
        resultSheet.Name = "VBA結果"
    End If
    resultSheet.Cells.Clear
    resultSheet.Cells(1, 1).Value = "商品名"
    resultSheet.Cells(1, 2).Value = "売上合計"

    Dim r As Long: r = 2
    Dim key As Variant
    For Each key In totals.Keys
        resultSheet.Cells(r, 1).Value = key
        resultSheet.Cells(r, 2).Value = totals(key)
        r = r + 1
    Next key

    ' ----- おまじない解除 -----
    Application.ScreenUpdating = True
    Application.Calculation = xlCalculationAutomatic
    Application.DisplayAlerts = True
    Application.EnableEvents = True

    ' ----- 結果表示 -----
    Dim elapsed As Double
    elapsed = Timer - startTime
    MsgBox _
        "VBA 集計完了" & vbCrLf & vbCrLf & _
        "ファイル数 : " & fileCount & vbCrLf & _
        "合計行数   : " & rowCount & vbCrLf & _
        "経過時間   : " & Format(elapsed, "0.00") & " 秒", _
        vbInformation, "VBA 速度ベンチマーク"
End Sub
