Attribute VB_Name = "Module_箱数計算転記"
' ============================================================================
'  作業予定表印刷(豚) → 計算フォーマット(豚) 転記マクロ
'
'  何のため:
'    現場から届く作業予定表を加工した「作業予定表印刷(豚)」シートを読んで、
'    今まで手書きしていた「計算フォーマット」を丸ごと自動で作る。
'
'  実行するもの: 計算フォーマットを作る （このモジュールの入口）
'
'  判定条件は「箱計算マスタ」「略称マスタ」シートに外出ししてある。
'  条件を変えたいときはマクロではなくシートを直す。
' ============================================================================
Option Explicit
' Option Compare は既定（Binary）のまま使う。
' 文字比較を「ロケール依存のテキスト比較」に頼ると PC の設定で結果が変わるため、
' 全角／半角・大小文字の吸収は NormKey() だけの責任にする。

' ---- 入力（作業予定表印刷(豚)）-------------------------------------------
Private Const SRC_WB_NAME    As String = "作業表加工用マクロ_レイテク作成_2.xlsm"
Private Const SRC_WS_NAME    As String = "作業予定表印刷(豚)"
Private Const SRC_ROW_DATE   As Long = 3      ' B3 に日付文字列
Private Const SRC_ROW_HEADER As Long = 5      ' 5行目がヘッダ
Private Const SRC_ROW_FIRST  As Long = 6      ' 6行目からデータ
Private Const SC_ORDER As Long = 2            ' B 順番
Private Const SC_NO    As Long = 3            ' C 作業№
Private Const SC_NAME  As Long = 4            ' D 作業名称
Private Const SC_HEAD  As Long = 5            ' E 頭数
Private Const SC_HALF  As Long = 8            ' H 半箱＋加工条件（マクロ生成文字列）
Private Const SC_CUT   As Long = 9            ' I カット内容
Private Const SC_FARM  As Long = 10           ' J 生産農場
Private Const SC_MAX   As Long = 13           ' M まで読む

' ---- 出力（このブック）---------------------------------------------------
Private Const FMT_WS  As String = "計算フォーマット(豚)"
Private Const MST_WS  As String = "箱計算マスタ"
Private Const ABBR_WS As String = "略称マスタ"
Private Const LOG_WS  As String = "判定ログ"

Private Const ROW_SUMMARY As Long = 1         ' 無地半／黒半・日付
Private Const ROW_TOTAL   As Long = 2         ' 総頭数
Private Const ROW_ITEM    As Long = 3         ' 項目（略称）＝表の1行目
Private Const ROW_COUNT   As Long = 11        ' 項目～No.9 までの行数
Private Const ROW_LAST    As Long = 13        ' ROW_ITEM + ROW_COUNT - 1
Private Const ROW_NOTE    As Long = 15        ' 印刷範囲外に出す警告サマリ

' 表の中の相対位置（ROW_ITEM からのオフセット。0 起点）
Private Const OFS_ITEM  As Long = 0
Private Const OFS_HEAD  As Long = 1
Private Const OFS_BOX   As Long = 2
Private Const OFS_UDE   As Long = 3
Private Const OFS_MOMO  As Long = 4
Private Const OFS_ROSU  As Long = 5
Private Const OFS_BARA  As Long = 6
Private Const OFS_KATA  As Long = 7
Private Const OFS_HIRE  As Long = 8
Private Const OFS_CHIMA As Long = 9
Private Const OFS_NO9   As Long = 10

Private Const FIRST_DATA_COL As Long = 2      ' B 列から作業ごとの列を作る
Private Const SUMMARY_LAST_COL As Long = 10   ' 冒頭集計欄は J 列まで
Private Const MAX_WORK_COLS  As Long = 200    ' 異常件数の上限（これを超えたら止める）
Private Const MAX_HEADS      As Long = 100000 ' 頭数として受け付ける上限（桁あふれ防止）
Private Const CLEAR_COLS     As Long = 220    ' 毎回クリアする列数

Private Const EXCLUDED_MARK As String = "/"   ' コンテナで箱を作らない部位の印

' No.6箱の判定キーワード（2026-09-27 REQ-007）
' 「美星黒豚」の「単品」だけが No.6。「会社単品ﾁﾙﾄﾞ(黒豚)」等の黒豚単品は通常のヒレ箱にする。
' マスタ化はせず定数のまま（要件台帳 REQ-009 の方針どおり、今回はここまで）。
Private Const KW_NO6_1 As String = "美星黒豚"
Private Const KW_NO6_2 As String = "単品"

' ---- マスタの1行 --------------------------------------------------------
Private Type TRule
    RuleNo  As String
    Keyword As String        ' 正規化済みキーワード
    RawKey  As String        ' 表示用（元のまま）
    Target  As String        ' 作業名称 / カット内容 / 生産農場 / 両方
    Value   As String        ' 除外＝処理区分、全農＝区分、略称＝略称
    Memo    As String
End Type

' ---- 「■ 1頭セット条件」マスタの1行（2026-09-28 REQ-010）-----------------
' 判定キーワードはマスタに出す方針（コードに直書きしない。要件台帳 REQ-009）。
' 対象の箱で「黒」は書けない（黒豚を巻き込まない。CheckRules1Set で弾く）。
Private Type TRule1
    RuleNo   As String
    Keyword  As String        ' 正規化済みキーワード
    RawKey   As String        ' 表示用（元のまま）
    Target   As String        ' 作業名称 / カット内容 / 生産農場 / 両方
    BoxScope As String        ' 対象の箱: 全 / 自 / 全・自
    Ude      As String        ' 部位ごとの計算方法: 頭数 / 半分切捨 / 半分切上 / 斜線 / 空欄（空セルも空欄）
    Momo     As String
    Rosu     As String
    Bara     As String
    Kata     As String
    Hire     As String
    Memo     As String
End Type

' ---- 「■ シートの印」マスタの1行（REQ-014）--------------------------------
' 区分（シートなし／全農のシートあり／会社のシートあり）ごとに、数字の前へ付ける印を持つ。
' 固定で3区分だけを想定する（1頭セット条件のような「上から順に評価」ではない）。
Private Type TMarkRow
    RowNo As Long             ' 不備メッセージ・4行目以降チェック用の実際の行番号
    Label As String           ' 区分（Trim 済み）
    Mark  As String           ' 印（Trim 済み。空欄可＝印なし）
    Memo  As String
End Type

' ---- 1作業ぶんの判定結果 ------------------------------------------------
Private Type TWork
    SrcRow      As Long
    OrderNo     As String
    WorkNo      As String
    WorkName    As String
    HalfRaw     As String
    CutName     As String
    FarmName    As String
    Heads       As Long
    HeadsOK     As Boolean

    Result      As String    ' 列を作成 / 半頭箱 / 除外 / エラー
    Reason      As String
    BoxKind     As String    ' 自 / 全 / 黒
    BoxReason   As String
    Zennou      As String    ' 全農 / 自社
    ZennouWhy   As String
    SheetYesNo  As String    ' シート あり / なし（REQ-014。列を作成した行だけ意味を持つ）
    SheetWhy    As String    ' シート判定の根拠（H列明示／作業名称チルド／カット内容チルド／既定あり）
    SheetMark   As String    ' 数字の前に付ける印（●・○・空欄。マスタ「■ シートの印」で決まる）
    ExcludeParts As String
    Abbr        As String
    AbbrWhy     As String
    OutCol      As Long
    NeedCheck   As String
    Warn        As String

    ' 部位ごとの出力値（Variant。数値または "/" または空）
    Ude As Variant
    Momo As Variant
    Rosu As Variant
    Bara As Variant
    Kata As Variant
    Hire As Variant
End Type

' ---- Application 設定の退避 ---------------------------------------------
Private mSaveScreen As Boolean
Private mSaveCalc   As XlCalculation
Private mSaveEvents As Boolean
Private mSaved      As Boolean
Private mPrintWarn  As String     ' 印刷設定に失敗したときの警告
Private mMasterNote As String     ' マスタの表（自動追加できるもの）の自動追加・追加失敗を知らせる
                                   ' （REQ-010「1頭セット条件」・REQ-014「シートの印」。NoteMaster で追記する）

' 開始時の設定を覚えておく（正常終了・エラー終了のどちらでも元に戻すため）
Private Sub SaveAppState()
    If mSaved Then Exit Sub
    mSaveScreen = Application.ScreenUpdating
    mSaveCalc = Application.Calculation
    mSaveEvents = Application.EnableEvents
    mSaved = True
End Sub

' 復帰処理そのものが失敗しても元のエラーを消さないよう、ここだけは握って進む
Private Sub RestoreAppState()
    If Not mSaved Then Exit Sub
    On Error Resume Next
    Application.Calculation = mSaveCalc
    Application.EnableEvents = mSaveEvents
    Application.ScreenUpdating = mSaveScreen
    On Error GoTo 0
    mSaved = False
End Sub

' ===========================================================================
'  入口
' ===========================================================================
Public Sub 計算フォーマットを作る()
    Dim srcWs As Worksheet
    Dim rulesExclude() As TRule, nExclude As Long
    Dim rulesZennou() As TRule, nZennou As Long
    Dim rulesHalfSkip() As TRule, nHalfSkip As Long
    Dim rulesAbbr() As TRule, nAbbr As Long
    Dim rules1Set() As TRule1, n1Set As Long
    Dim markSheetNone As String, markSheetZen As String, markSheetJisha As String
    Dim halfBoxPerHead As Long, colsPerPage As Long, fallbackLen As Long
    Dim works() As TWork, nWorks As Long
    Dim sumWhiteYes As Long, sumWhiteNo As Long
    Dim sumBlackYes As Long, sumBlackNo As Long
    Dim totalHeads As Long, sheetTotal As Long, hasSheetTotal As Boolean
    Dim dateText As String
    Dim warnings As String, warnCount As Long
    Dim msg As String

    On Error GoTo ErrHandler
    mMasterNote = ""

    ' --- 1. 入力シートを掴む（見つからなければここで終わり）---
    Set srcWs = FindSourceSheet()
    If srcWs Is Nothing Then Exit Sub
    If Not ValidateSourceHeader(srcWs) Then Exit Sub
    If srcWs.Parent Is ThisWorkbook Then
        MsgBox "このマクロは入力ブックとは別のブック（計算フォーマット）に入れて実行してください。" & vbCrLf & _
               "今は入力ブック自身の中で動こうとしています。", vbCritical
        Exit Sub
    End If

    ' --- 2. マスタを読む（不備があればここで止める）---
    If Not LoadMasters(rulesExclude, nExclude, rulesZennou, nZennou, _
                       rulesHalfSkip, nHalfSkip, rulesAbbr, nAbbr, _
                       rules1Set, n1Set, _
                       markSheetNone, markSheetZen, markSheetJisha, _
                       halfBoxPerHead, colsPerPage, fallbackLen) Then
        Exit Sub
    End If

    SaveAppState
    Application.ScreenUpdating = False
    Application.Calculation = xlCalculationManual
    Application.EnableEvents = False

    ' --- 3. 全行を判定する ---
    If Not JudgeAllRows(srcWs, rulesExclude, nExclude, rulesZennou, nZennou, _
                        rulesHalfSkip, nHalfSkip, rulesAbbr, nAbbr, _
                        rules1Set, n1Set, _
                        markSheetNone, markSheetZen, markSheetJisha, _
                        halfBoxPerHead, fallbackLen, _
                        works, nWorks, _
                        sumWhiteYes, sumWhiteNo, sumBlackYes, sumBlackNo, _
                        totalHeads, sheetTotal, hasSheetTotal, dateText) Then
        RestoreAppState
        Exit Sub
    End If

    ' --- 4. 検算 ---
    If hasSheetTotal And sheetTotal <> totalHeads Then
        AppendWarn warnings, warnCount, _
            "総頭数が合いません。予定表の合計行=" & sheetTotal & " / 明細の合計=" & totalHeads
    End If

    ' --- 5. 書き出す ---
    mPrintWarn = ""
    WriteFormatSheet works, nWorks, sumWhiteYes, sumWhiteNo, sumBlackYes, sumBlackNo, _
                     totalHeads, dateText, colsPerPage
    If Len(mPrintWarn) > 0 Then AppendWarn warnings, warnCount, mPrintWarn
    WriteLogSheet works, nWorks, warnings

    ' --- 6. 警告をまとめて知らせる ---
    CollectRowWarnings works, nWorks, warnCount
    WriteNoteOnFormat warnCount, warnings
    RestoreAppState

    msg = "計算フォーマットを作りました。" & vbCrLf & vbCrLf & _
          "日付　　　： " & dateText & vbCrLf & _
          "総頭数　　： " & totalHeads & " 頭" & vbCrLf & _
          "列を作った： " & CountResult(works, nWorks, "列を作成") & " 作業" & vbCrLf & _
          "半頭箱　　： " & CountResult(works, nWorks, "半頭箱") & " 作業" & vbCrLf & _
          "除外　　　： " & CountResult(works, nWorks, "除外") & " 作業" & vbCrLf & vbCrLf & _
          "無地半　あり " & sumWhiteYes & " ／ なし " & sumWhiteNo & vbCrLf & _
          "黒半　　あり " & sumBlackYes & " ／ なし " & sumBlackNo
    If Len(mMasterNote) > 0 Then
        msg = msg & vbCrLf & vbCrLf & mMasterNote
    End If
    If warnCount > 0 Then
        msg = msg & vbCrLf & vbCrLf & "▲ 確認してほしいことが " & warnCount & " 件あります。" & vbCrLf & _
              "「判定ログ」シートの「警告」の列（T列）を見てください。"
        MsgBox msg, vbExclamation
    Else
        MsgBox msg, vbInformation
    End If
    Exit Sub

ErrHandler:
    Dim errNo As Long, errText As String
    errNo = Err.Number
    errText = Err.Description
    RestoreAppState
    MsgBox AppendMasterNote("エラーが発生したため中断しました。" & vbCrLf & _
           "エラー番号: " & errNo & vbCrLf & "内容: " & errText), vbCritical
End Sub

' ===========================================================================
'  入力ブック・シートの特定と検査
' ===========================================================================

' 開いているブックの中から入力シートを探す。同名ブックが複数あるときは知らせる。
Private Function FindSourceSheet() As Worksheet
    Dim wb As Workbook, ws As Worksheet
    Dim hit As Worksheet, hitCount As Long
    Dim names As String

    For Each wb In Application.Workbooks
        If Not wb Is ThisWorkbook Then
            For Each ws In wb.Worksheets
                If ws.Name = SRC_WS_NAME Then
                    hitCount = hitCount + 1
                    If hit Is Nothing Then Set hit = ws
                    names = names & vbCrLf & "　・" & wb.FullName
                End If
            Next ws
        End If
    Next wb

    If hitCount = 0 Then
        MsgBox "「" & SRC_WS_NAME & "」シートが見つかりません。" & vbCrLf & vbCrLf & _
               "「" & SRC_WB_NAME & "」を開いて、作業予定表の加工まで済ませてから" & vbCrLf & _
               "もう一度このボタンを押してください。", vbExclamation
        Exit Function
    ElseIf hitCount > 1 Then
        MsgBox "「" & SRC_WS_NAME & "」シートを持つブックが " & hitCount & " 個開いています。" & names & vbCrLf & vbCrLf & _
               "どれを使えばよいか決められないので、使わないほうを閉じてから実行してください。", vbExclamation
        Exit Function
    End If

    Set FindSourceSheet = hit
End Function

' 5行目のヘッダを照合して、列がずれていないか確かめる。
Private Function ValidateSourceHeader(ByVal ws As Worksheet) As Boolean
    Dim checks(1 To 4, 1 To 2) As String
    Dim i As Long, actual As String

    checks(1, 1) = CStr(SC_NO):   checks(1, 2) = "作業"
    checks(2, 1) = CStr(SC_NAME): checks(2, 2) = "作業名称"
    checks(3, 1) = CStr(SC_HEAD): checks(3, 2) = "頭数"
    checks(4, 1) = CStr(SC_FARM): checks(4, 2) = "生産農場"

    For i = 1 To 4
        actual = NormKey(ws.Cells(SRC_ROW_HEADER, CLng(checks(i, 1))).Value)
        If InStr(1, actual, NormKey(checks(i, 2))) = 0 Then
            MsgBox "「" & SRC_WS_NAME & "」の列の並びが想定と違います。" & vbCrLf & vbCrLf & _
                   SRC_ROW_HEADER & "行目 " & ColLetter(CLng(checks(i, 1))) & "列 は「" & checks(i, 2) & "」のはずですが、" & vbCrLf & _
                   "実際は「" & actual & "」でした。" & vbCrLf & vbCrLf & _
                   "列が挿入・削除されていないか確認してください。誤った計算をしないためここで止めます。", vbCritical
            Exit Function
        End If
    Next i
    ValidateSourceHeader = True
End Function

' ===========================================================================
'  マスタ読み込み
' ===========================================================================
Private Function LoadMasters(ByRef rExc() As TRule, ByRef nExc As Long, _
                             ByRef rZen() As TRule, ByRef nZen As Long, _
                             ByRef rSkip() As TRule, ByRef nSkip As Long, _
                             ByRef rAbbr() As TRule, ByRef nAbbr As Long, _
                             ByRef r1Set() As TRule1, ByRef n1Set As Long, _
                             ByRef markSheetNone As String, ByRef markSheetZen As String, ByRef markSheetJisha As String, _
                             ByRef halfPerHead As Long, ByRef colsPerPage As Long, _
                             ByRef fallbackLen As Long) As Boolean
    Dim wsM As Worksheet, wsA As Worksheet
    Dim err1 As String, structErr1 As String
    Dim markRows() As TMarkRow, nMarkRows As Long, structErrMark As String, validateMark As Boolean

    Set wsM = GetSheet(MST_WS)
    Set wsA = GetSheet(ABBR_WS)
    If wsM Is Nothing Or wsA Is Nothing Then
        MsgBox "「" & MST_WS & "」または「" & ABBR_WS & "」シートが見つかりません。" & vbCrLf & _
               "このブックが壊れている可能性があります。", vbCritical
        Exit Function
    End If

    nExc = ReadRuleTable(wsM, "■ 除外条件", rExc)
    nZen = ReadRuleTable(wsM, "■ 全農／自社 条件", rZen)
    nSkip = ReadRuleTable(wsM, "■ 半頭箱の集計から除く条件", rSkip)
    nAbbr = ReadRuleTable(wsA, "■ 略称マスタ", rAbbr)

    ' 見出しが見つからないときは -1 が返る。表そのものが消えている状態なので止める
    If nExc < 0 Or nZen < 0 Or nSkip < 0 Or nAbbr < 0 Then
        MsgBox "マスタの見出し行（「■ …」で始まる行）が見つかりません。" & vbCrLf & _
               "「" & MST_WS & "」「" & ABBR_WS & "」シートの見出しを消していないか確認してください。" & vbCrLf & vbCrLf & _
               "除外条件=" & nExc & " / 全農条件=" & nZen & " / 半頭集計除外=" & nSkip & " / 略称=" & nAbbr, vbCritical
        Exit Function
    End If

    ' 「1頭セット条件」は先方がマクロを触れない前提のマスタ方式（要件台帳 REQ-010）。
    ' 見出しが無い版のブック（9/11版など）でも止めず、初回だけ初期値で自動追加する。
    n1Set = ReadRule1Table(wsM, "■ 1頭セット条件", r1Set, structErr1)
    If n1Set = -1 Then
        If AddInitial1SetTable(wsM, "■ 1頭セット条件") Then
            n1Set = ReadRule1Table(wsM, "■ 1頭セット条件", r1Set, structErr1)
        Else
            n1Set = 0
            ReDim r1Set(1 To 1)
        End If
    End If

    ' 「シートの印」も同じくマスタ方式（要件台帳 REQ-014）。見出しが無ければ初回だけ
    ' 既定値（シートなし＝●／全農のシートあり＝○／会社のシートあり＝空欄）で自動追加する。
    ' 自動追加に失敗したときは、表の中身は検査せずこの既定値のまま続行する（validateMark=False）。
    validateMark = True
    nMarkRows = ReadSheetMarkTable(wsM, "■ シートの印", markRows, structErrMark)
    If nMarkRows = -1 Then
        If AddInitialSheetMarkTable(wsM, "■ シートの印") Then
            nMarkRows = ReadSheetMarkTable(wsM, "■ シートの印", markRows, structErrMark)
        Else
            markSheetNone = "●": markSheetZen = "○": markSheetJisha = ""
            validateMark = False
        End If
    End If

    halfPerHead = ReadSetting(wsM, "半頭箱 1頭あたりの箱数", 2)
    colsPerPage = ReadSetting(wsM, "1ページあたりの列数", 15)
    fallbackLen = ReadSetting(wsM, "略称が未登録のとき使う文字数", 6)

    ' --- 中身の検査（先に止めたほうが安全なもの）---
    err1 = CheckRules(rExc, nExc, "除外条件", "全量除外|部位のみ除外", True)
    If Len(err1) = 0 Then err1 = CheckRules(rZen, nZen, "全農／自社 条件", "全農|自社", True)
    If Len(err1) = 0 Then err1 = CheckRules(rSkip, nSkip, "半頭箱の集計から除く条件", "", True)
    ' 略称マスタは「No / キーワード / 略称 / メモ」の4列で、3列目が略称なので対象列の検査はしない
    If Len(err1) = 0 Then err1 = CheckRules(rAbbr, nAbbr, "略称マスタ", "", False)
    If Len(err1) = 0 Then err1 = structErr1
    If Len(err1) = 0 And n1Set > 0 Then err1 = CheckRules1Set(r1Set, n1Set)
    If Len(err1) = 0 Then err1 = structErrMark
    If Len(err1) = 0 And validateMark Then
        err1 = CheckSheetMarkRows(markRows, nMarkRows, markSheetNone, markSheetZen, markSheetJisha)
    End If
    If Len(err1) > 0 Then
        MsgBox AppendMasterNote("マスタの書き方に不備があります。直してから実行してください。" & vbCrLf & vbCrLf & err1), vbCritical
        Exit Function
    End If
    If nZen = 0 Then
        MsgBox AppendMasterNote("「" & MST_WS & "」の全農／自社 条件が1件もありません。すべて「自」になってしまうため止めます。"), vbCritical
        Exit Function
    End If
    If halfPerHead < 1 Then halfPerHead = 2
    If colsPerPage < 1 Then colsPerPage = 15
    If fallbackLen < 1 Then fallbackLen = 6

    LoadMasters = True
End Function

' 見出し行（"■ …"）を探し、その2行下から次の見出しまでを1つの表として読む。
' 途中に空行やメモ行があっても読み飛ばすだけで打ち切らない。
Private Function ReadRuleTable(ByVal ws As Worksheet, ByVal sectionTitle As String, _
                               ByRef rules() As TRule) As Long
    Dim lastRow As Long, r As Long, startRow As Long
    Dim n As Long, kw As String
    Dim a1 As String

    ReDim rules(1 To 1)
    ReadRuleTable = -1                  ' 見出しが無いことを表す
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 500 Then lastRow = 500

    ' 見出しを探す
    For r = 1 To lastRow
        a1 = SafeStr(ws.Cells(r, 1).Value)
        If Left$(a1, Len(sectionTitle)) = sectionTitle Then
            startRow = r + 2          ' 見出し行 + ヘッダ行 の次
            Exit For
        End If
    Next r
    If startRow = 0 Then Exit Function
    ReadRuleTable = 0
    If startRow > lastRow Then Exit Function        ' 見出しだけで中身が無い

    ReDim rules(1 To lastRow - startRow + 2)
    For r = startRow To lastRow
        a1 = SafeStr(ws.Cells(r, 1).Value)
        If Left$(a1, 1) = "■" Then Exit For        ' 次の見出しに来たら終わり
        kw = Trim$(SafeStr(ws.Cells(r, 2).Value))
        If Len(kw) > 0 Then
            n = n + 1
            rules(n).RuleNo = a1
            rules(n).RawKey = kw
            rules(n).Keyword = NormKey(kw)
            rules(n).Target = Trim$(SafeStr(ws.Cells(r, 3).Value))
            rules(n).Value = Trim$(SafeStr(ws.Cells(r, 4).Value))
            rules(n).Memo = Trim$(SafeStr(ws.Cells(r, 5).Value))
            If Len(rules(n).Target) = 0 Then rules(n).Target = "作業名称"
        End If
    Next r
    ReadRuleTable = n
End Function

' マスタの書き方を検査してエラー文を返す（空文字なら問題なし）
Private Function CheckRules(ByRef rules() As TRule, ByVal n As Long, _
                            ByVal tableName As String, ByVal allowed As String, _
                            ByVal checkTarget As Boolean) As String
    Dim i As Long, j As Long, msg As String
    Dim t As String

    For i = 1 To n
        t = rules(i).Target
        If checkTarget Then
            If t <> "作業名称" And t <> "カット内容" And t <> "生産農場" And t <> "両方" Then
                msg = msg & "・[" & tableName & "] No." & rules(i).RuleNo & " の対象列「" & t & _
                      "」は使えません。作業名称／カット内容／生産農場／両方 のどれかにしてください。" & vbCrLf
            End If
        Else
            If Len(t) = 0 Then
                msg = msg & "・[" & tableName & "] No." & rules(i).RuleNo & "「" & rules(i).RawKey & _
                      "」の略称が空欄です。" & vbCrLf
            End If
        End If
        If InStr(1, rules(i).RawKey, "[") > 0 Or InStr(1, rules(i).RawKey, "]") > 0 _
           Or InStr(1, rules(i).RawKey, "?") > 0 Or InStr(1, rules(i).RawKey, "#") > 0 Then
            msg = msg & "・[" & tableName & "] No." & rules(i).RuleNo & " のキーワードに ［ ］ ？ ＃ は使えません：" & _
                  rules(i).RawKey & vbCrLf
        End If
        ' 「*」だけのキーワードは全件に一致してしまうので止める
        If Len(Replace(rules(i).Keyword, "*", "")) = 0 Then
            msg = msg & "・[" & tableName & "] No." & rules(i).RuleNo & " のキーワードが「" & rules(i).RawKey & _
                  "」です。＊だけだと全部の作業に一致してしまうため使えません。" & vbCrLf
        End If
        If Len(allowed) > 0 Then
            If Not MatchesList(rules(i).Value, allowed) Then
                msg = msg & "・[" & tableName & "] No." & rules(i).RuleNo & " の「" & rules(i).Value & _
                      "」は使えません。" & Replace(allowed, "|", " または ") & " にしてください。" & vbCrLf
            End If
        End If
        For j = i + 1 To n
            If rules(i).Keyword = rules(j).Keyword And rules(i).Target = rules(j).Target Then
                msg = msg & "・[" & tableName & "] No." & rules(i).RuleNo & " と No." & rules(j).RuleNo & _
                      " が同じキーワード「" & rules(i).RawKey & "」です。下の行は使われません。" & vbCrLf
                Exit For
            End If
        Next j
    Next i
    CheckRules = msg
End Function

Private Function MatchesList(ByVal v As String, ByVal pipeList As String) As Boolean
    Dim parts() As String, i As Long
    parts = Split(pipeList, "|")
    For i = LBound(parts) To UBound(parts)
        If v = parts(i) Then MatchesList = True: Exit Function
    Next i
End Function

Private Function ReadSetting(ByVal ws As Worksheet, ByVal label As String, ByVal defVal As Long) As Long
    Dim lastRow As Long, r As Long
    ReadSetting = defVal
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 500 Then lastRow = 500
    For r = 1 To lastRow
        If NormKey(ws.Cells(r, 1).Value) = NormKey(label) Then
            If IsNumeric(ws.Cells(r, 2).Value) Then ReadSetting = CLng(ws.Cells(r, 2).Value)
            Exit Function
        End If
    Next r
End Function

' ===========================================================================
'  「1頭セット条件」マスタ（REQ-010・2026-09-28）
'
'  黒豚以外の全農「1頭ｾｯﾄ」は、通常の全箱（バラ÷2切捨）と違い部位ごとに
'  異なる計算方法（頭数そのまま／斜線 等）を使う。今後も同じ形の作業が増える
'  可能性があるため、他の判定と同じくコードに直書きせずマスタで持つ
'  （要件台帳 REQ-009 の方針）。先方はマクロを触れないので、初回だけ
'  見つからなければ自動で追加する（下の AddInitial1SetTable）。
' ===========================================================================

' 「■ 1頭セット条件」の11列表を読む。ReadRuleTable と同じ規則（見出し探索・500行打ち切り）に加え、
' ヘッダの並び・キーワード空欄なのに他列だけ入力・エラー値のセルを検査する（R6）。
' 見出しが無ければ -1（自動追加するかは呼び側の判断）。見出しはあるが中身が0件なら 0。
' structErr に不備があれば理由を積んで返す（空文字なら問題なし）。
Private Function ReadRule1Table(ByVal ws As Worksheet, ByVal sectionTitle As String, _
                                ByRef rules() As TRule1, ByRef structErr As String) As Long
    Dim lastRow As Long, r As Long, startRow As Long, headRow As Long, titleRow As Long
    Dim n As Long, kw As String, a1 As String
    Dim expectHead As Variant, c As Long, actual As String
    Dim dupCount As Long, dupRows As String

    structErr = ""
    ReDim rules(1 To 1)
    ReadRule1Table = -1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 500 Then lastRow = 500

    ' 見出しは1つだけのはず。2つ以上あれば表の重複（コピー貼り付け等）を疑い不備にする（REQ-014・C5）。
    ' 読込そのものは最初に見つかった見出しで続ける
    For r = 1 To lastRow
        a1 = SafeStr(ws.Cells(r, 1).Value)
        If Left$(a1, Len(sectionTitle)) = sectionTitle Then
            dupCount = dupCount + 1
            If dupCount = 1 Then
                titleRow = r
                headRow = r + 1
                startRow = r + 2
            Else
                If Len(dupRows) > 0 Then dupRows = dupRows & "・"
                dupRows = dupRows & r & "行目"
            End If
        End If
    Next r
    If startRow = 0 Then Exit Function
    ReadRule1Table = 0
    If dupCount > 1 Then
        structErr = structErr & "・[1頭セット条件] 見出し「" & sectionTitle & "」が2つ以上あります（" & _
                    titleRow & "行目のほかに " & dupRows & "）。" & vbCrLf
    End If

    ' --- ヘッダ（A～K）の並びを検査（R6）。中身が0件の空表でも必ず検査する ---
    expectHead = Array("No", "キーワード", "対象列", "対象の箱", "ウデ", "モモ", "ロース", "バラ", "カタ", "ヒレ", "メモ")
    For c = LBound(expectHead) To UBound(expectHead)
        actual = Trim$(SafeStr(ws.Cells(headRow, c + 1).Value))
        If actual <> expectHead(c) Then
            structErr = structErr & "・[1頭セット条件] ヘッダ行（" & headRow & "行目 " & ColLetter(c + 1) & _
                        "列）は「" & expectHead(c) & "」のはずですが「" & actual & "」でした。" & vbCrLf
        End If
    Next c

    If startRow > lastRow Then Exit Function        ' ヘッダ検査は済んだ。見出しとヘッダだけで中身が無い

    ReDim rules(1 To lastRow - startRow + 2)
    For r = startRow To lastRow
        a1 = SafeStr(ws.Cells(r, 1).Value)
        If Left$(a1, 1) = "■" Then Exit For        ' 次の見出しに来たら終わり

        ' --- エラー値のセルは不備（R6）---
        For c = 1 To 11
            If IsError(ws.Cells(r, c).Value) Then
                structErr = structErr & "・[1頭セット条件] " & r & "行目 " & ColLetter(c) & "列がエラー値です。" & vbCrLf
                Exit For
            End If
        Next c

        kw = Trim$(SafeStr(ws.Cells(r, 2).Value))
        If Len(kw) = 0 Then
            ' --- キーワード（B）が空なのに C～K（メモ含む）に入力がある行は不備（R6）。
            '     No（A列）だけが入っている行（半頭除外の表などにある「No.1 だけの空行」と同じ運用）は
            '     読み飛ばしてよいので、A列は検査に含めない ---
            For c = 3 To 11
                If Len(Trim$(SafeStr(ws.Cells(r, c).Value))) > 0 Then
                    structErr = structErr & "・[1頭セット条件] " & r & "行目はキーワード（B列）が空欄なのに" & _
                                ColLetter(c) & "列に入力があります。行を空にするかキーワードを入れてください。" & vbCrLf
                    Exit For
                End If
            Next c
        Else
            n = n + 1
            rules(n).RuleNo = a1
            rules(n).RawKey = kw
            rules(n).Keyword = NormKey(kw)
            rules(n).Target = Trim$(SafeStr(ws.Cells(r, 3).Value))
            rules(n).BoxScope = Trim$(SafeStr(ws.Cells(r, 4).Value))
            rules(n).Ude = Trim$(SafeStr(ws.Cells(r, 5).Value))
            rules(n).Momo = Trim$(SafeStr(ws.Cells(r, 6).Value))
            rules(n).Rosu = Trim$(SafeStr(ws.Cells(r, 7).Value))
            rules(n).Bara = Trim$(SafeStr(ws.Cells(r, 8).Value))
            rules(n).Kata = Trim$(SafeStr(ws.Cells(r, 9).Value))
            rules(n).Hire = Trim$(SafeStr(ws.Cells(r, 10).Value))
            rules(n).Memo = Trim$(SafeStr(ws.Cells(r, 11).Value))
            If Len(rules(n).Target) = 0 Then rules(n).Target = "作業名称"
        End If
    Next r
    ReadRule1Table = n
End Function

' 「1頭セット条件」マスタの書き方を検査してエラー文を返す（空文字なら問題なし）。
' 既存の CheckRules と違い、重複・到達不能は「正規化キーワード＋対象列＋対象の箱」で見る（R7）。
Private Function CheckRules1Set(ByRef rules() As TRule1, ByVal n As Long) As String
    Dim i As Long, j As Long, msg As String, t As String, bx As String

    For i = 1 To n
        t = rules(i).Target
        If t <> "作業名称" And t <> "カット内容" And t <> "生産農場" And t <> "両方" Then
            msg = msg & "・[1頭セット条件] No." & rules(i).RuleNo & " の対象列「" & t & _
                  "」は使えません。作業名称／カット内容／生産農場／両方 のどれかにしてください。" & vbCrLf
        End If

        bx = rules(i).BoxScope
        If bx <> "全" And bx <> "自" And bx <> "全・自" Then
            msg = msg & "・[1頭セット条件] No." & rules(i).RuleNo & " の対象の箱「" & bx & _
                  "」は使えません。全／自／全・自 のどれかにしてください（黒は指定できません）。" & vbCrLf
        End If

        ' 既存 CheckRules と同じ流儀（［ ］？＃禁止・＊だけ禁止）
        If InStr(1, rules(i).RawKey, "[") > 0 Or InStr(1, rules(i).RawKey, "]") > 0 _
           Or InStr(1, rules(i).RawKey, "?") > 0 Or InStr(1, rules(i).RawKey, "#") > 0 Then
            msg = msg & "・[1頭セット条件] No." & rules(i).RuleNo & " のキーワードに ［ ］ ？ ＃ は使えません：" & _
                  rules(i).RawKey & vbCrLf
        End If
        ' 「*」だけのキーワードは全件に一致してしまうので止める
        If Len(Replace(rules(i).Keyword, "*", "")) = 0 Then
            msg = msg & "・[1頭セット条件] No." & rules(i).RuleNo & " のキーワードが「" & rules(i).RawKey & _
                  "」です。＊だけだと全部の作業に一致してしまうため使えません。" & vbCrLf
        End If

        msg = msg & CheckPartMethod("ウデ", rules(i).RuleNo, rules(i).Ude)
        msg = msg & CheckPartMethod("モモ", rules(i).RuleNo, rules(i).Momo)
        msg = msg & CheckPartMethod("ロース", rules(i).RuleNo, rules(i).Rosu)
        msg = msg & CheckPartMethod("バラ", rules(i).RuleNo, rules(i).Bara)
        msg = msg & CheckPartMethod("カタ", rules(i).RuleNo, rules(i).Kata)
        msg = msg & CheckPartMethod("ヒレ", rules(i).RuleNo, rules(i).Hire)

        For j = i + 1 To n
            If rules(i).Keyword = rules(j).Keyword And rules(i).Target = rules(j).Target Then
                If rules(i).BoxScope = rules(j).BoxScope Then
                    msg = msg & "・[1頭セット条件] No." & rules(i).RuleNo & " と No." & rules(j).RuleNo & _
                          " が同じキーワード「" & rules(i).RawKey & "」・同じ対象の箱です。下の行は使われません。" & vbCrLf
                    Exit For
                ElseIf rules(i).BoxScope = "全・自" Then
                    ' 先に「全・自」があると、後の同キーワードの「全」「自」は絶対に届かない
                    msg = msg & "・[1頭セット条件] No." & rules(j).RuleNo & " は No." & rules(i).RuleNo & _
                          "「" & rules(i).RawKey & "」(対象の箱=全・自) より後にあり、同じキーワードでは到達できません。" & vbCrLf
                End If
            End If
        Next j
    Next i
    CheckRules1Set = msg
End Function

Private Function CheckPartMethod(ByVal label As String, ByVal ruleNo As String, ByVal v As String) As String
    Select Case v
        Case "", "頭数", "半分切捨", "半分切上", "斜線", "空欄"
            CheckPartMethod = ""
        Case Else
            CheckPartMethod = "・[1頭セット条件] No." & ruleNo & " の" & label & "「" & v & _
                "」は使えません。頭数／半分切捨／半分切上／斜線／空欄 のどれかにしてください（空欄可）。" & vbCrLf
    End Select
End Function

' 「■ 1頭セット条件」が見つからないときに初期値1行で自動追加する（R2～R4）。
' 書く順は「初期値行→ヘッダ→見出し」（見出しを最後に書く）。こうしておけば、
' 途中で失敗して元に戻せなくても「見出しが無い」＝未追加のままなので、次回また
' 安全にやり直せる（見出しが中途半端に生きたまま中身が壊れる、という状態を作らない）。
Private Function AddInitial1SetTable(ByVal ws As Worksheet, ByVal sectionTitle As String) As Boolean
    Const NEEDCOLS As Long = 11
    Dim lastCellRow As Long, titleRow As Long, headRow As Long, dataRow As Long
    Dim c As Range
    Dim dataArr(1 To 1, 1 To NEEDCOLS) As Variant
    Dim headArr(1 To 1, 1 To NEEDCOLS) As Variant

    AddInitial1SetTable = False

    ' R3: 見出しの位置は「A列」ではなく「シート全体の使用済み最終行」の2行下
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If c Is Nothing Then lastCellRow = 0 Else lastCellRow = c.Row

    titleRow = lastCellRow + 2
    headRow = lastCellRow + 3
    dataRow = lastCellRow + 4

    ' R4: 読込の上限（500行）を超えて書いても次回読めないので、書かずに0件で続行する
    If dataRow > 500 Then
        NoteMaster "「箱計算マスタ」がいっぱいで「■ 1頭セット条件」の表を自動追加できませんでした" & _
                   "（追加先が500行を超えます）。空いている行を作るか、手動で表を追加してください。" & _
                   "1頭ｾｯﾄ条件は今回0件で続行します。"
        Exit Function
    End If
    If ws.ProtectContents Then
        NoteMaster "「" & ws.Name & "」シートが保護されているため「■ 1頭セット条件」の表を自動追加できませんでした。" & _
                   "保護を解除して再実行するか、手動で表を追加してください。1頭ｾｯﾄ条件は今回0件で続行します。"
        Exit Function
    End If

    On Error GoTo Failed

    ' 初期値1行（要件台帳 REQ-010・Q-04。カタ・ヒレは仮決めの空欄）
    dataArr(1, 1) = "1": dataArr(1, 2) = "1頭ｾｯﾄ": dataArr(1, 3) = "カット内容": dataArr(1, 4) = "全"
    dataArr(1, 5) = "頭数": dataArr(1, 6) = "頭数": dataArr(1, 7) = "斜線": dataArr(1, 8) = "頭数"
    dataArr(1, 9) = "": dataArr(1, 10) = "": dataArr(1, 11) = "全農の1頭ｾｯﾄ（2026/09/11）。カタ・ヒレは仮決め"

    headArr(1, 1) = "No": headArr(1, 2) = "キーワード": headArr(1, 3) = "対象列": headArr(1, 4) = "対象の箱"
    headArr(1, 5) = "ウデ": headArr(1, 6) = "モモ": headArr(1, 7) = "ロース": headArr(1, 8) = "バラ"
    headArr(1, 9) = "カタ": headArr(1, 10) = "ヒレ": headArr(1, 11) = "メモ"

    ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow, NEEDCOLS)).Value = dataArr
    ws.Range(ws.Cells(headRow, 1), ws.Cells(headRow, NEEDCOLS)).Value = headArr
    ws.Cells(titleRow, 1).Value = sectionTitle & _
        "　ここに書いた作業は、部位ごとの箱数をこの表の計算方法で出します（上から順に評価し、最初に合致した行で確定）"

    AddInitial1SetTable = True
    NoteMaster "箱計算マスタに「■ 1頭セット条件」の表を追加しました。ブックを保存してください。"
    Exit Function

Failed:
    ' Err はここで即座に退避する（この後 On Error 文・Resume を書くと Err がクリアされて
    ' Err.Description が空になるため。既存 ErrHandler の errNo/errText と同じ考え方）
    Dim errNo As Long, errText As String
    errNo = Err.Number
    errText = Err.Description
    ' エラー処理中のまま .Clear を試みると、そこで新たなエラーが起きたときに
    ' On Error Resume Next が効かず呼び出し元へ伝わってしまう。Resume で後始末用の
    ' ラベルへ移り、エラー処理中の状態を抜けてから安全に後始末する
    Resume CleanUp1Set

CleanUp1Set:
    On Error Resume Next
    ws.Range(ws.Cells(titleRow, 1), ws.Cells(dataRow, NEEDCOLS)).Clear
    On Error GoTo 0
    NoteMaster "「■ 1頭セット条件」の表の自動追加に失敗しました（" & errText & _
               "）。1頭ｾｯﾄ条件は今回0件で続行します。"
    AddInitial1SetTable = False
End Function

' ===========================================================================
'  「シートの印」マスタ（REQ-014・2026-09-28）
'
'  紙の計算フォーマットでは各列の数字が赤丸・青丸で囲まれ、これがシートの有無を
'  表す先方の一番の要望（要件台帳 REQ-014）。図形の丸は使わず、数字の前に記号を
'  付ける方式にした。記号は「1頭セット条件」と同じくマスタで持ち、
'  先方がマクロを触らずに変えられるようにする。
' ===========================================================================

' 「■ シートの印」の3行表（区分／印／メモ）を読む。ReadRuleTable と同じ規則
' （見出し探索・500行打ち切り）に加え、ヘッダの並び・エラー値のセルを検査する。
' 見出しが無ければ -1（自動追加するかは呼び側の判断）。見出しはあるが中身が0件なら 0。
' rows() には見つかった行を順に積む（3行を超えても打ち切らない。件数の検査は
' CheckSheetMarkRows 側で行う＝「4行目以降にデータがある」を検出できるようにするため）。
Private Function ReadSheetMarkTable(ByVal ws As Worksheet, ByVal sectionTitle As String, _
                                    ByRef rows() As TMarkRow, ByRef structErr As String) As Long
    Dim lastRow As Long, r As Long, startRow As Long, headRow As Long, titleRow As Long
    Dim n As Long, a1 As String, lab As String
    Dim expectHead As Variant, c As Long, actual As String
    Dim dupCount As Long, dupRows As String

    structErr = ""
    ReDim rows(1 To 1)
    ReadSheetMarkTable = -1
    lastRow = ws.Cells(ws.Rows.Count, 1).End(xlUp).Row
    If lastRow > 500 Then lastRow = 500

    ' 見出しは1つだけのはず。2つ以上あれば表の重複（コピー貼り付け等）を疑い不備にする。
    ' 読込そのものは最初に見つかった見出しで続ける
    For r = 1 To lastRow
        a1 = SafeStr(ws.Cells(r, 1).Value)
        If Left$(a1, Len(sectionTitle)) = sectionTitle Then
            dupCount = dupCount + 1
            If dupCount = 1 Then
                titleRow = r
                headRow = r + 1
                startRow = r + 2
            Else
                If Len(dupRows) > 0 Then dupRows = dupRows & "・"
                dupRows = dupRows & r & "行目"
            End If
        End If
    Next r
    If startRow = 0 Then Exit Function
    ReadSheetMarkTable = 0
    If dupCount > 1 Then
        structErr = structErr & "・[シートの印] 見出し「" & sectionTitle & "」が2つ以上あります（" & _
                    titleRow & "行目のほかに " & dupRows & "）。" & vbCrLf
    End If

    ' --- ヘッダ（A～C）の並びを検査。中身が0件の空表でも必ず検査する ---
    expectHead = Array("区分", "印", "メモ")
    For c = LBound(expectHead) To UBound(expectHead)
        actual = Trim$(SafeStr(ws.Cells(headRow, c + 1).Value))
        If actual <> expectHead(c) Then
            structErr = structErr & "・[シートの印] ヘッダ行（" & headRow & "行目 " & ColLetter(c + 1) & _
                        "列）は「" & expectHead(c) & "」のはずですが「" & actual & "」でした。" & vbCrLf
        End If
    Next c

    If startRow > lastRow Then Exit Function        ' ヘッダ検査は済んだ。見出しとヘッダだけで中身が無い

    ReDim rows(1 To lastRow - startRow + 2)
    For r = startRow To lastRow
        a1 = SafeStr(ws.Cells(r, 1).Value)
        If Left$(a1, 1) = "■" Then Exit For        ' 次の見出しに来たら終わり

        ' --- エラー値のセルは不備 ---
        For c = 1 To 3
            If IsError(ws.Cells(r, c).Value) Then
                structErr = structErr & "・[シートの印] " & r & "行目 " & ColLetter(c) & "列がエラー値です。" & vbCrLf
                Exit For
            End If
        Next c

        lab = Trim$(a1)
        If Len(lab) = 0 Then
            ' 区分（A列）が空なのに 印／メモ に入力がある行は不備（他のマスタの
            ' 「キーワード空欄なのに他列に入力」と同じ考え方）
            If Len(Trim$(SafeStr(ws.Cells(r, 2).Value))) > 0 Or Len(Trim$(SafeStr(ws.Cells(r, 3).Value))) > 0 Then
                structErr = structErr & "・[シートの印] " & r & "行目は区分（A列）が空欄なのに印／メモに入力があります。" & vbCrLf
            End If
        Else
            n = n + 1
            rows(n).RowNo = r
            rows(n).Label = lab
            rows(n).Mark = TrimWide(SafeStr(ws.Cells(r, 2).Value))
            rows(n).Memo = Trim$(SafeStr(ws.Cells(r, 3).Value))
        End If
    Next r
    ReadSheetMarkTable = n
End Function

' 「シートの印」マスタの書き方を検査してエラー文を返す（空文字なら問題なし）。
' 3区分（シートなし／全農のシートあり／会社のシートあり）が各1回だけ・4行目以降に
' データが無いことを見る。読めた印は markNone/markZen/markJisha に詰めて返す。
Private Function CheckSheetMarkRows(ByRef rows() As TMarkRow, ByVal n As Long, _
                                    ByRef markNone As String, ByRef markZen As String, ByRef markJisha As String) As String
    Const LBL_NONE As String = "シートなし"
    Const LBL_ZEN As String = "全農のシートあり"
    Const LBL_JISHA As String = "会社のシートあり"
    Dim i As Long, msg As String
    Dim cntNone As Long, cntZen As Long, cntJisha As Long

    markNone = "": markZen = "": markJisha = ""

    ' 「4行目以降」は表の物理的な行位置（1行目＝最初に見つかった行）で判定する。
    ' 空行やメモだけの行は rows() に入らないため、単純に i（何件目に見つけたか）で
    ' 判定すると「空行を挟んで離れた3件目」を見逃す（Codexレビュー指摘）。
    ' rows(1).RowNo を表の1行目とみなし、そこから3行（行番号差2）を超えた位置に
    ' データがあれば不備にする
    For i = 1 To n
        If n > 0 Then
            If rows(i).RowNo - rows(1).RowNo >= 3 Then
                If InStr(1, msg, "4行目以降にデータ") = 0 Then
                    msg = msg & "・[シートの印] 表の4行目以降にデータがあります。区分は「" & LBL_NONE & "」「" & _
                          LBL_ZEN & "」「" & LBL_JISHA & "」の3行だけにしてください（" & rows(i).RowNo & "行目）。" & vbCrLf
                End If
            End If
        End If

        Select Case rows(i).Label
            Case LBL_NONE
                cntNone = cntNone + 1
                markNone = rows(i).Mark
            Case LBL_ZEN
                cntZen = cntZen + 1
                markZen = rows(i).Mark
            Case LBL_JISHA
                cntJisha = cntJisha + 1
                markJisha = rows(i).Mark
            Case Else
                msg = msg & "・[シートの印] " & rows(i).RowNo & "行目の区分「" & rows(i).Label & _
                      "」は使えません。「" & LBL_NONE & "」「" & LBL_ZEN & "」「" & LBL_JISHA & "」のどれかにしてください。" & vbCrLf
        End Select

        msg = msg & CheckMarkValue(rows(i).RowNo, rows(i).Label, rows(i).Mark)
    Next i

    If cntNone = 0 Then msg = msg & "・[シートの印] 区分「" & LBL_NONE & "」がありません。" & vbCrLf
    If cntNone > 1 Then msg = msg & "・[シートの印] 区分「" & LBL_NONE & "」が2つ以上あります。" & vbCrLf
    If cntZen = 0 Then msg = msg & "・[シートの印] 区分「" & LBL_ZEN & "」がありません。" & vbCrLf
    If cntZen > 1 Then msg = msg & "・[シートの印] 区分「" & LBL_ZEN & "」が2つ以上あります。" & vbCrLf
    If cntJisha = 0 Then msg = msg & "・[シートの印] 区分「" & LBL_JISHA & "」がありません。" & vbCrLf
    If cntJisha > 1 Then msg = msg & "・[シートの印] 区分「" & LBL_JISHA & "」が2つ以上あります。" & vbCrLf

    CheckSheetMarkRows = msg
End Function

' 印1件ぶんの書き方を検査する（Trim 後の値で判定。空欄＝印なしで可）。
' 改行・数字（全角含む）・「/」（全角含む）・3文字を超える長さは不備にする
' （数字の前に付く印なので、数字や「/」が混ざると数の読み違いのもとになる）。
Private Function CheckMarkValue(ByVal rowNo As Long, ByVal label As String, ByVal mk As String) As String
    Dim i As Long, ch As String

    If Len(mk) = 0 Then Exit Function      ' 印なしは可

    If InStr(1, "=+-'@", Left$(mk, 1)) > 0 Then
        CheckMarkValue = "・[シートの印] " & rowNo & "行目「" & label & "」の印がセルの先頭に使えない文字（= + - ' @）で" & _
            "始まっています（Excelが数式などと解釈します）：" & mk & vbCrLf
        Exit Function
    End If
    If InStr(1, mk, vbLf) > 0 Or InStr(1, mk, vbCr) > 0 Then
        CheckMarkValue = "・[シートの印] " & rowNo & "行目「" & label & "」の印に改行は使えません：" & mk & vbCrLf
        Exit Function
    End If
    If InStr(1, mk, "/") > 0 Or InStr(1, mk, ChrW(&HFF0F)) > 0 Then     ' 半角／全角の「/」
        CheckMarkValue = "・[シートの印] " & rowNo & "行目「" & label & "」の印に「/」は使えません（部位が無い印と紛れます）：" & mk & vbCrLf
        Exit Function
    End If
    For i = 1 To Len(mk)
        ch = Mid$(mk, i, 1)
        If (ch >= "0" And ch <= "9") Or (ch >= ChrW(&HFF10) And ch <= ChrW(&HFF19)) Then   ' 半角0-9／全角０-９
            CheckMarkValue = "・[シートの印] " & rowNo & "行目「" & label & "」の印に数字は使えません（数字の前に付く印なので紛らわしくなります）：" & mk & vbCrLf
            Exit Function
        End If
    Next i
    If Len(mk) > 3 Then
        CheckMarkValue = "・[シートの印] " & rowNo & "行目「" & label & "」の印が長すぎます（" & Len(mk) & "文字）。1～3文字にしてください：" & mk & vbCrLf
        Exit Function
    End If
End Function

' 「■ シートの印」が見つからないときに既定の3行（シートなし＝●／全農のシートあり＝○／
' 会社のシートあり＝空欄）で自動追加する。AddInitial1SetTable と同じ作法（書く順は
' 「データ→ヘッダ→見出し」、位置はシート全体の使用済み最終行の2行下、500行上限、
' 保護時は書かずに既定値で続行、失敗時は消して戻す）。
Private Function AddInitialSheetMarkTable(ByVal ws As Worksheet, ByVal sectionTitle As String) As Boolean
    Const NEEDCOLS As Long = 3
    Dim lastCellRow As Long, titleRow As Long, headRow As Long, dataRow As Long
    Dim c As Range
    Dim dataArr(1 To 3, 1 To NEEDCOLS) As Variant
    Dim headArr(1 To 1, 1 To NEEDCOLS) As Variant

    AddInitialSheetMarkTable = False

    ' 見出しの位置は「シート全体の使用済み最終行」の2行下（1頭セット条件と同じ。
    ' 1頭セット条件がこの実行で足されていれば、その最終行から求まる）
    Set c = ws.Cells.Find(What:="*", After:=ws.Cells(1, 1), LookIn:=xlFormulas, _
                          LookAt:=xlPart, SearchOrder:=xlByRows, SearchDirection:=xlPrevious, MatchCase:=False)
    If c Is Nothing Then lastCellRow = 0 Else lastCellRow = c.Row

    titleRow = lastCellRow + 2
    headRow = lastCellRow + 3
    dataRow = lastCellRow + 4

    ' 3行使うので、最後に書く行（dataRow+2）が読込の上限（500行）に収まるか見る
    If dataRow + 2 > 500 Then
        NoteMaster "「箱計算マスタ」がいっぱいで「■ シートの印」の表を自動追加できませんでした" & _
                   "（追加先が500行を超えます）。空いている行を作るか、手動で表を追加してください。" & _
                   "シートの印は既定（シートなし＝●／全農のシートあり＝○／会社のシートあり＝空欄）で続行します。"
        Exit Function
    End If
    If ws.ProtectContents Then
        NoteMaster "「" & ws.Name & "」シートが保護されているため「■ シートの印」の表を自動追加できませんでした。" & _
                   "保護を解除して再実行するか、手動で表を追加してください。シートの印は既定で続行します。"
        Exit Function
    End If

    On Error GoTo Failed

    ' 既定の3行（要件台帳 REQ-014・2026-09-28）
    dataArr(1, 1) = "シートなし": dataArr(1, 2) = "●": dataArr(1, 3) = "紙の赤丸。会社・全農・黒豚とも"
    dataArr(2, 1) = "全農のシートあり": dataArr(2, 2) = "○": dataArr(2, 3) = "紙の青丸"
    dataArr(3, 1) = "会社のシートあり": dataArr(3, 2) = "": dataArr(3, 3) = "紙では印なし。黒豚のシートありもここに従う（仮決め）"

    headArr(1, 1) = "区分": headArr(1, 2) = "印": headArr(1, 3) = "メモ"

    ws.Range(ws.Cells(dataRow, 1), ws.Cells(dataRow + 2, NEEDCOLS)).Value = dataArr
    ws.Range(ws.Cells(headRow, 1), ws.Cells(headRow, NEEDCOLS)).Value = headArr
    ws.Cells(titleRow, 1).Value = sectionTitle & _
        "　各列の数字の前に付ける印です（先方要望の赤丸・青丸の代わり）。印を空欄にすると印を付けません"

    AddInitialSheetMarkTable = True
    NoteMaster "箱計算マスタに「■ シートの印」の表を追加しました。ブックを保存してください。"
    Exit Function

Failed:
    Dim errNo As Long, errText As String
    errNo = Err.Number
    errText = Err.Description
    Resume CleanUpMark

CleanUpMark:
    On Error Resume Next
    ws.Range(ws.Cells(titleRow, 1), ws.Cells(dataRow + 2, NEEDCOLS)).Clear
    On Error GoTo 0
    NoteMaster "「■ シートの印」の表の自動追加に失敗しました（" & errText & _
               "）。シートの印は既定（シートなし＝●／全農のシートあり＝○／会社のシートあり＝空欄）で続行します。"
    AddInitialSheetMarkTable = False
End Function

' ===========================================================================
'  判定の本体
' ===========================================================================
Private Function JudgeAllRows(ByVal srcWs As Worksheet, _
        ByRef rExc() As TRule, ByVal nExc As Long, _
        ByRef rZen() As TRule, ByVal nZen As Long, _
        ByRef rSkip() As TRule, ByVal nSkip As Long, _
        ByRef rAbbr() As TRule, ByVal nAbbr As Long, _
        ByRef r1Set() As TRule1, ByVal n1Set As Long, _
        ByVal markSheetNone As String, ByVal markSheetZen As String, ByVal markSheetJisha As String, _
        ByVal halfPerHead As Long, ByVal fallbackLen As Long, _
        ByRef works() As TWork, ByRef nWorks As Long, _
        ByRef sumWhiteYes As Long, ByRef sumWhiteNo As Long, _
        ByRef sumBlackYes As Long, ByRef sumBlackNo As Long, _
        ByRef totalHeads As Long, ByRef sheetTotal As Long, ByRef hasSheetTotal As Boolean, _
        ByRef dateText As String) As Boolean

    Dim lastRow As Long, r As Long, i As Long
    Dim data As Variant
    Dim nameN As String, cutN As String, farmN As String
    Dim halfColor As String, halfBoxes As Long, isHalf As Boolean, halfSheet As String, halfSheetWhy As String
    Dim contState As Long, parts As String     ' contState: 0=なし 1=部位あり 2=解析失敗
    Dim w As TWork
    Dim outCol As Long
    Dim lastZennou As String, lastZennouRow As Long
    Dim hit As Long, hit1 As Long

    ' --- 日付 ---
    dateText = Trim$(SafeStr(srcWs.Cells(SRC_ROW_DATE, SC_ORDER).Value))

    ' --- データ最終行（作業名称のある最後の行）---
    lastRow = 0
    For r = srcWs.Cells(srcWs.Rows.Count, SC_NAME).End(xlUp).Row To SRC_ROW_FIRST Step -1
        If Len(Trim$(SafeStr(srcWs.Cells(r, SC_NAME).Value))) > 0 Then lastRow = r: Exit For
    Next r
    If lastRow < SRC_ROW_FIRST Then
        MsgBox AppendMasterNote("「" & SRC_WS_NAME & "」に作業データがありません（" & SRC_ROW_FIRST & "行目以降）。"), vbExclamation
        Exit Function
    End If
    If lastRow - SRC_ROW_FIRST + 1 > MAX_WORK_COLS Then
        MsgBox AppendMasterNote("作業件数が " & (lastRow - SRC_ROW_FIRST + 1) & " 件あります（上限 " & MAX_WORK_COLS & " 件）。" & vbCrLf & _
               "入力データがおかしくないか確認してください。"), vbCritical
        Exit Function
    End If

    ' --- 合計行（作業名称が空で頭数が数値の行）を探す ---
    For r = lastRow + 1 To lastRow + 3
        If r <= srcWs.Rows.Count Then
            If Len(Trim$(SafeStr(srcWs.Cells(r, SC_NAME).Value))) = 0 Then
                If IsNumeric(srcWs.Cells(r, SC_HEAD).Value) And Not IsEmpty(srcWs.Cells(r, SC_HEAD).Value) Then
                    sheetTotal = CLng(srcWs.Cells(r, SC_HEAD).Value)
                    hasSheetTotal = True
                    Exit For
                End If
            End If
        End If
    Next r

    ' --- 一括読み込み（B:M）---
    data = srcWs.Range(srcWs.Cells(SRC_ROW_FIRST, SC_ORDER), srcWs.Cells(lastRow, SC_MAX)).Value

    ReDim works(1 To lastRow - SRC_ROW_FIRST + 1)
    lastZennou = ""
    lastZennouRow = 0

    For i = 1 To UBound(data, 1)
        r = SRC_ROW_FIRST + i - 1
        Erase2 w

        w.SrcRow = r
        w.OrderNo = SafeStr(data(i, SC_ORDER - SC_ORDER + 1))
        w.WorkNo = SafeStr(data(i, SC_NO - SC_ORDER + 1))
        w.WorkName = SafeStr(data(i, SC_NAME - SC_ORDER + 1))
        w.HalfRaw = SafeStr(data(i, SC_HALF - SC_ORDER + 1))
        w.CutName = SafeStr(data(i, SC_CUT - SC_ORDER + 1))
        w.FarmName = SafeStr(data(i, SC_FARM - SC_ORDER + 1))

        If Len(Trim$(w.WorkName)) = 0 Then
            ' 途中の空行。読み飛ばすが記録は残す
            nWorks = nWorks + 1
            w.Result = "エラー"
            w.Reason = "作業名称が空欄のため読み飛ばしました"
            w.Warn = "作業名称が空欄の行があります（入力行 " & r & "）"
            works(nWorks) = w
            GoTo NextRow
        End If

        ' 頭数（1以上 MAX_HEADS 以下の整数だけを有効とする）
        ParseHeads data(i, SC_HEAD - SC_ORDER + 1), w.Heads, w.HeadsOK
        If Not w.HeadsOK Then
            w.Warn = AddSep(w.Warn, "頭数として使えません（" & SafeStr(data(i, SC_HEAD - SC_ORDER + 1)) & _
                     "）。1～" & MAX_HEADS & " の整数を入れてください。この列の箱数は空欄にしました")
        End If

        nameN = NormKey(w.WorkName)
        cutN = NormKey(w.CutName)
        farmN = NormKey(w.FarmName)

        ' ---------- 全農／自社の判定（すべての行で行う）----------
        hit = MatchRule(rZen, nZen, nameN, cutN, farmN)
        If hit > 0 Then
            w.Zennou = rZen(hit).Value
            w.ZennouWhy = "マスタ No." & rZen(hit).RuleNo & "「" & rZen(hit).RawKey & "」(" & rZen(hit).Target & ")"
            lastZennou = w.Zennou
            lastZennouRow = r
        Else
            If Len(lastZennou) > 0 Then
                w.Zennou = lastZennou
                w.ZennouWhy = "直前に確定した入力行 " & lastZennouRow & " から継承"
            Else
                w.Zennou = "自社"
                w.ZennouWhy = "継承元が無いため初期値（自社）"
            End If
            w.NeedCheck = "要目視確認"
            ' 警告に出すのは、この判定が箱種に効く行だけ（後で箱種が決まってから足す）
        End If

        ' ---------- STEP1 半頭箱 ----------
        ParseHalfBox w.HalfRaw, halfColor, halfBoxes, isHalf, halfSheet
        If isHalf Then
            w.Result = "半頭箱"
            ' シート有無（H列の明示 → 無ければチルド → 無ければあり）は列の判定（STEP4.5）と
            ' 同じ関数を使う（要件台帳 REQ-014）。halfSheetWhy はここでは使わない
            halfSheet = ResolveSheetYesNo(halfSheet, nameN, cutN, halfSheetWhy)
            If halfBoxes > 0 Then
                w.Warn = AddSep(w.Warn, "「" & Trim$(w.HalfRaw) & "」の指定により箱数を " & halfBoxes & " として集計しました（頭数×" & halfPerHead & " ではありません）")
            Else
                halfBoxes = w.Heads * halfPerHead
            End If
            w.Reason = "半箱欄が「" & halfColor & "」→ " & IIf(halfColor = "白", "無地半", "黒半") & _
                       "（シート" & halfSheet & "）に " & halfBoxes & " 箱を加算"

            If MatchRule(rSkip, nSkip, nameN, cutN, farmN) > 0 Then
                w.Reason = w.Reason & " ※集計除外マスタに一致したため加算しません"
                w.Warn = AddSep(w.Warn, "半頭箱の集計除外マスタに一致したため、箱数を集計に入れていません")
            ElseIf halfColor = "白" Then
                If halfSheet = "あり" Then sumWhiteYes = sumWhiteYes + halfBoxes Else sumWhiteNo = sumWhiteNo + halfBoxes
            Else
                If halfSheet = "あり" Then sumBlackYes = sumBlackYes + halfBoxes Else sumBlackNo = sumBlackNo + halfBoxes
            End If

            totalHeads = totalHeads + w.Heads
            nWorks = nWorks + 1
            works(nWorks) = w
            GoTo NextRow
        End If

        ' ---------- STEP2 部位コンテナ ----------
        ParseContainerParts nameN, contState, parts
        w.ExcludeParts = parts

        ' ---------- STEP3 除外 ----------
        hit = MatchRule(rExc, nExc, nameN, cutN, farmN)
        If hit > 0 Then
            If rExc(hit).Value = "部位のみ除外" And contState = 1 Then
                ' 部位指定つきコンテナ → 列は作る
                w.Reason = "マスタ No." & rExc(hit).RuleNo & "「" & rExc(hit).RawKey & "」に一致（部位のみ除外）：" & parts
            Else
                w.Result = "除外"
                w.Reason = "マスタ No." & rExc(hit).RuleNo & "「" & rExc(hit).RawKey & "」(" & rExc(hit).Target & ") に一致 → " & rExc(hit).Value
                If contState = 2 Then
                    w.Reason = w.Reason & "（作業名称の「コンテナ」に部位の指定は無し＝全量）"
                End If
                totalHeads = totalHeads + w.Heads
                nWorks = nWorks + 1
                works(nWorks) = w
                GoTo NextRow
            End If
        ElseIf contState = 2 Then
            ' 除外マスタに当たらないが解析に失敗 → 通常計算に流さない
            w.Result = "除外"
            w.Reason = "作業名称に「コンテナ」があるが部位を読み取れなかったため全量除外（安全側）"
            w.Warn = AddSep(w.Warn, "作業名称に「コンテナ」があるのに部位を読み取れませんでした。安全側で全量除外にしています：" & w.WorkName)
            totalHeads = totalHeads + w.Heads
            nWorks = nWorks + 1
            works(nWorks) = w
            GoTo NextRow
        End If

        ' ---------- STEP4 箱種 ----------
        If InStr(1, nameN, NormKey("黒豚")) > 0 Then
            w.BoxKind = "黒"
            w.BoxReason = "作業名称に「黒豚」"
        ElseIf w.Zennou = "全農" Then
            w.BoxKind = "全"
            w.BoxReason = "全農判定 → 全農"
        Else
            w.BoxKind = "自"
            w.BoxReason = "全農判定 → 自社"
        End If

        ' 継承で判定した行のうち、箱種が「全／自」に効くものだけ警告する
        ' （黒箱・除外・半頭箱は全農判定が出力に出ないので黙っておく）
        If w.NeedCheck = "要目視確認" And w.BoxKind <> "黒" Then
            w.Warn = AddSep(w.Warn, "全農／自社をマスタで判定できず「" & w.Zennou & "」として継承したので、箱が「" & w.BoxKind & "」になっています（" & w.ZennouWhy & "）")
        End If

        ' ---------- STEP4.5 シートの印（REQ-014）----------
        ' 半頭箱（STEP1）と同じ決まり（H列の明示 → 無ければチルド → 無ければあり）を列にも当てる。
        ' 印は「シートなし→マスタの印(●)」「全＋シートあり→マスタの印(○)」「自・黒＋シートあり→マスタの印(空欄)」
        ' （黒のシートありは紙・受領データに例が無いため自社と同じ扱いの仮決め。要件台帳 REQ-014）。
        w.SheetYesNo = ResolveSheetYesNo(SheetFromHalfText(w.HalfRaw), nameN, cutN, w.SheetWhy)
        If w.SheetYesNo = "なし" Then
            w.SheetMark = markSheetNone
        ElseIf w.BoxKind = "全" Then
            w.SheetMark = markSheetZen
        Else
            w.SheetMark = markSheetJisha
        End If

        ' ---------- 1頭セット条件（REQ-010）----------
        ' 箱種が決まった後に、その箱種を対象に含む行だけを候補にする（黒の列は評価しない）。
        ' 一致判定と根拠の記録はここで行い、数量への反映は STEP5（頭数が読めた行だけ）で行う。
        hit1 = Match1SetRule(r1Set, n1Set, w.BoxKind, nameN, cutN, farmN)
        If hit1 > 0 Then
            w.Reason = AddSep(w.Reason, Describe1SetMatch(r1Set(hit1)))
        End If

        ' ---------- STEP5 部位数量 ----------
        FillParts w, parts, r1Set, hit1

        ' ---------- STEP6 略称 ----------
        hit = MatchRuleName(rAbbr, nAbbr, nameN)
        If hit > 0 Then
            w.Abbr = rAbbr(hit).Target      ' 略称マスタは3列目が略称
            w.AbbrWhy = "マスタ No." & rAbbr(hit).RuleNo & "「" & rAbbr(hit).RawKey & "」"
        Else
            w.Abbr = Left$(w.WorkName, fallbackLen)
            w.AbbrWhy = "略称マスタに無いため作業名称の先頭" & fallbackLen & "文字（仮）"
            w.Warn = AddSep(w.Warn, "略称マスタに登録がありません。仮に「" & w.Abbr & "」としました：" & w.WorkName)
        End If

        w.Result = "列を作成"
        If Len(w.Reason) = 0 Then w.Reason = "通常の箱計算対象"
        outCol = outCol + 1
        w.OutCol = FIRST_DATA_COL + outCol - 1
        totalHeads = totalHeads + w.Heads
        nWorks = nWorks + 1
        works(nWorks) = w

NextRow:
    Next i

    JudgeAllRows = True
End Function

' 頭数セルを解釈する。1～MAX_HEADS の整数以外は無効として扱う。
' （IsNumeric だけだと Long あふれ・小数・指数表記を通してしまう）
Private Sub ParseHeads(ByVal v As Variant, ByRef heads As Long, ByRef okFlag As Boolean)
    Dim d As Double
    heads = 0: okFlag = False
    If IsError(v) Then Exit Sub
    If IsNull(v) Then Exit Sub
    If IsEmpty(v) Then Exit Sub
    If Not IsNumeric(v) Then Exit Sub
    d = CDbl(v)
    If d < 1 Or d > MAX_HEADS Then Exit Sub
    If d <> Int(d) Then Exit Sub
    heads = CLng(d)
    okFlag = True
End Sub

' 箱種に応じて部位の箱数を入れる。コンテナ除外部位は "/" にする。
' r1Set/hit1 は1頭セット条件マスタとその一致行（0=不一致。REQ-010）。
Private Sub FillParts(ByRef w As TWork, ByVal excluded As String, _
                      ByRef r1Set() As TRule1, ByVal hit1 As Long)
    Dim n As Long
    n = w.Heads

    ' 頭数が読めなかった行は数字を出さない（0 を並べて誤解させないため）
    If Not w.HeadsOK Then
        w.Ude = Empty: w.Momo = Empty: w.Rosu = Empty
        w.Bara = Empty: w.Kata = Empty: w.Hire = Empty
        Exit Sub
    End If

    w.Ude = n
    w.Momo = n
    Select Case w.BoxKind
        Case "黒"
            w.Rosu = n
            w.Bara = n
            w.Kata = n
            If InStr(1, NormKey(w.WorkName), NormKey(KW_NO6_1)) > 0 And _
               InStr(1, NormKey(w.WorkName), NormKey(KW_NO6_2)) > 0 Then
                ' 出力先はヒレ行で確定（2026-09-28・要件台帳 Q-01）。警告は出さない
                w.Hire = "No.6" & vbLf & n
            Else
                w.Hire = n
            End If
        Case "全"
            w.Rosu = Int(n / 2)
            w.Bara = Int(n / 2)
            w.Kata = Empty
            w.Hire = Empty
        Case Else       ' 自
            w.Rosu = -Int(-n / 2)      ' 切り上げ
            w.Bara = -Int(-n / 2)
            w.Kata = Empty
            w.Hire = Empty
    End Select

    ' ---- 1頭セット条件（REQ-010）: 一致した行の計算方法で ウデ～ヒレ を決め直す ----
    If hit1 > 0 Then
        w.Ude = ComputePart1(r1Set(hit1).Ude, n)
        w.Momo = ComputePart1(r1Set(hit1).Momo, n)
        w.Rosu = ComputePart1(r1Set(hit1).Rosu, n)
        w.Bara = ComputePart1(r1Set(hit1).Bara, n)
        w.Kata = ComputePart1(r1Set(hit1).Kata, n)
        w.Hire = ComputePart1(r1Set(hit1).Hire, n)
    End If

    If Len(excluded) > 0 Then
        If InStr(1, excluded, "ウデ") > 0 Then w.Ude = EXCLUDED_MARK
        If InStr(1, excluded, "モモ") > 0 Then w.Momo = EXCLUDED_MARK
        If InStr(1, excluded, "ロース") > 0 Then w.Rosu = EXCLUDED_MARK
        If InStr(1, excluded, "バラ") > 0 Then w.Bara = EXCLUDED_MARK
        If InStr(1, excluded, "カタ") > 0 Then w.Kata = EXCLUDED_MARK
        If InStr(1, excluded, "ヒレ") > 0 Then w.Hire = EXCLUDED_MARK
    End If
End Sub

' 「1頭セット条件」マスタ1件が当たるか調べる。当たった行番号を返す（0=なし）。
' 箱種が決まった後に、その箱種を対象の箱に含む行だけを候補にする。黒の列は評価しない（REQ-010）。
Private Function Match1SetRule(ByRef r1Set() As TRule1, ByVal n1Set As Long, _
                               ByVal boxKind As String, _
                               ByVal nameN As String, ByVal cutN As String, ByVal farmN As String) As Long
    Dim i As Long
    If boxKind = "黒" Then Exit Function
    For i = 1 To n1Set
        If BoxScopeCovers(r1Set(i).BoxScope, boxKind) Then
            Select Case r1Set(i).Target
                Case "作業名称"
                    If MatchKeyword1(nameN, r1Set(i).Keyword) Then Match1SetRule = i: Exit Function
                Case "カット内容"
                    If MatchKeyword1(cutN, r1Set(i).Keyword) Then Match1SetRule = i: Exit Function
                Case "生産農場"
                    If MatchKeyword1(farmN, r1Set(i).Keyword) Then Match1SetRule = i: Exit Function
                Case Else       ' 両方
                    If MatchKeyword1(nameN, r1Set(i).Keyword) Or MatchKeyword1(cutN, r1Set(i).Keyword) _
                       Or MatchKeyword1(farmN, r1Set(i).Keyword) Then Match1SetRule = i: Exit Function
            End Select
        End If
    Next i
End Function

' 対象の箱（全／自／全・自）が今の箱種を含むか
Private Function BoxScopeCovers(ByVal scope As String, ByVal boxKind As String) As Boolean
    BoxScopeCovers = (scope = boxKind) Or (scope = "全・自")
End Function

' 「1頭セット条件」専用のキーワード一致。
' キーワードに「*」を含む場合は従来どおり target Like "*"&keyword&"*"。
' 含まず数字始まりの場合は、出現位置ごとに「直前の1文字が数字でない（または先頭）」ものが
' 1つでもあれば一致とする（「1頭ｾｯﾄ」が「11頭ｾｯﾄ」「21頭ｾｯﾄ」の部分に当たらないようにするため）。
Private Function MatchKeyword1(ByVal target As String, ByVal keyword As String) As Boolean
    Dim p As Long, prevCh As String

    If Len(keyword) = 0 Then Exit Function
    If InStr(1, keyword, "*") > 0 Then
        MatchKeyword1 = target Like "*" & keyword & "*"
        Exit Function
    End If
    If Not (Left$(keyword, 1) >= "0" And Left$(keyword, 1) <= "9") Then
        MatchKeyword1 = target Like "*" & keyword & "*"
        Exit Function
    End If

    p = InStr(1, target, keyword)
    Do While p > 0
        If p = 1 Then
            MatchKeyword1 = True: Exit Function
        End If
        prevCh = Mid$(target, p - 1, 1)
        If Not (prevCh >= "0" And prevCh <= "9") Then
            MatchKeyword1 = True: Exit Function
        End If
        p = InStr(p + 1, target, keyword)
    Loop
End Function

' 「1頭セット条件」マスタの計算方法1件ぶんの値を出す
Private Function ComputePart1(ByVal method As String, ByVal n As Long) As Variant
    Select Case method
        Case "頭数"
            ComputePart1 = n
        Case "半分切捨"
            ComputePart1 = Int(n / 2)
        Case "半分切上"
            ComputePart1 = -Int(-n / 2)
        Case "斜線"
            ComputePart1 = EXCLUDED_MARK
        Case Else       ' 空欄（空セルもここに来る）
            ComputePart1 = Empty
    End Select
End Function

' 判定ログの理由欄に足す「1頭セット条件に一致した」旨の文言
Private Function Describe1SetMatch(ByRef rule As TRule1) As String
    Describe1SetMatch = "1頭セット条件 No." & rule.RuleNo & "「" & rule.RawKey & "」(" & rule.Target & _
        ") に一致: ウデ=" & PartLabel(rule.Ude) & " モモ=" & PartLabel(rule.Momo) & _
        " ロース=" & PartLabel(rule.Rosu) & " バラ=" & PartLabel(rule.Bara) & _
        " カタ=" & PartLabel(rule.Kata) & " ヒレ=" & PartLabel(rule.Hire)
End Function

Private Function PartLabel(ByVal method As String) As String
    If Len(Trim$(method)) = 0 Then PartLabel = "空欄" Else PartLabel = method
End Function

' ===========================================================================
'  文字列の解析
' ===========================================================================

' H列の先頭トークンから半頭箱を読む。
' 区切り（全角空白）を消す前に先頭を切り出すのが肝心。
Private Sub ParseHalfBox(ByVal raw As String, ByRef color As String, ByRef boxes As Long, _
                         ByRef isHalf As Boolean, ByRef sheetYesNo As String)
    Dim s As String, token As String, rest As String
    Dim p As Long, k As Long, numTxt As String

    color = "": boxes = 0: isHalf = False: sheetYesNo = ""
    s = raw
    If Len(Trim$(s)) = 0 Then Exit Sub

    ' 区切りを空白1文字に寄せてから先頭トークンを取る
    s = Replace(s, ChrW(&H3000), " ")
    s = Replace(s, vbTab, " ")
    s = Trim$(s)
    p = InStr(1, s, " ")
    If p > 0 Then token = Left$(s, p - 1) Else token = s

    token = NarrowSafe(token)
    token = Trim$(token)
    If Len(token) = 0 Then Exit Sub

    If Left$(token, 1) = "白" Then
        color = "白"
    ElseIf Left$(token, 1) = "黒" Then
        color = "黒"
    Else
        Exit Sub
    End If

    rest = Mid$(token, 2)
    ' 「白×3」「白x3」「白*3」などの倍数指定
    Do While Len(rest) > 0
        Select Case Left$(rest, 1)
            Case ChrW(&HD7), "x", "X", "*", " ", ChrW(&H2715)   ' × x X * 空白 と U+2715
                rest = Mid$(rest, 2)
            Case Else
                Exit Do
        End Select
    Loop
    For k = 1 To Len(rest)
        If Len(numTxt) >= 6 Then Exit For           ' 桁あふれ防止
        If Mid$(rest, k, 1) >= "0" And Mid$(rest, k, 1) <= "9" Then
            numTxt = numTxt & Mid$(rest, k, 1)
        Else
            Exit For
        End If
    Next k
    If Len(numTxt) > 0 Then boxes = CLng(numTxt)
    rest = Mid$(rest, Len(numTxt) + 1)

    ' 「白」「黒」の後に残った文字が括弧書きでもなければ、半頭箱の表記ではないと見なす
    ' （「白豚」「黒豚」「白紙」などを半頭箱と誤認しないため）
    rest = Trim$(rest)
    If Len(rest) > 0 Then
        Select Case Left$(rest, 1)
            Case "(", "[", ChrW(&HFF08), ChrW(&HFF3B)   ' 半角( 半角[ 全角( 全角[
                ' 括弧書きの補足（例: 白×20(シート有)）は許す
            Case Else
                color = "": boxes = 0
                Exit Sub
        End Select
    End If
    isHalf = True

    ' シート有無は先頭トークンだけでなく H 列の全文から探す
    ' （「白×3　シート無」のように離れて書かれることがある。列のシート判定にも
    ' 使う関数を共用する。要件台帳 REQ-014・C1）
    sheetYesNo = SheetFromHalfText(raw)
End Sub

' H列（半箱欄）の全文だけを見て、シート有無の明示を読む（副作用なし。REQ-014）。
' ParseHalfBox（半頭箱の判定）と列のシート判定（ResolveSheetYesNo 経由）の両方から使う。
Private Function SheetFromHalfText(ByVal raw As String) As String
    Dim whole As String
    whole = NormKey(raw)
    If InStr(1, whole, NormKey("シート有")) > 0 Then
        SheetFromHalfText = "あり"
    ElseIf InStr(1, whole, NormKey("シート無")) > 0 Or InStr(1, whole, NormKey("シートなし")) > 0 Then
        SheetFromHalfText = "なし"
    End If
End Function

' シート有無を決める（H列の明示 → 無ければ作業名称のチルド → 無ければカット内容のチルド
' → 無ければ「あり」）。半頭箱（STEP1）と列（STEP4.5）の両方から使う（要件台帳 REQ-014）。
' whyText には判定の根拠を返す（列の判定ログに使う。半頭箱側は使わなくてよい）。
Private Function ResolveSheetYesNo(ByVal explicitYesNo As String, ByVal nameN As String, ByVal cutN As String, _
                                   ByRef whyText As String) As String
    If Len(explicitYesNo) > 0 Then
        ResolveSheetYesNo = explicitYesNo
        whyText = "H列の明示（" & IIf(explicitYesNo = "あり", "シート有", "シート無/シートなし") & "）"
    ElseIf InStr(1, nameN, NormKey("チルド")) > 0 Then
        ResolveSheetYesNo = "なし"
        whyText = "作業名称に「チルド」"
    ElseIf InStr(1, cutN, NormKey("チルド")) > 0 Then
        ResolveSheetYesNo = "なし"
        whyText = "カット内容に「チルド」"
    Else
        ResolveSheetYesNo = "あり"
        whyText = "チルドの記載なし（既定あり）"
    End If
End Function

' 作業名称の「(モモ･バラコンテナ)」から除外部位を読む。
'   state 0 = コンテナの記載なし
'   state 1 = 部位を読み取れた（parts に列挙）
'   state 2 = コンテナはあるが部位を読み取れない → 呼び側で全量除外にする
Private Sub ParseContainerParts(ByVal nameN As String, ByRef state As Long, ByRef parts As String)
    Dim partNames As Variant, partsNorm As Variant
    Dim p As Long, i As Long, tail As String
    Dim matched As Boolean
    Const SEPS As String = "･・,、/⇒>＞"

    state = 0: parts = ""
    p = InStr(1, nameN, NormKey("コンテナ"))
    If p = 0 Then Exit Sub

    ' 長いものから先に見る（ロース より先に ロース、が来ないよう順序を固定）
    partNames = Array("ロース", "カルビ", "チマキ", "ウデ", "モモ", "バラ", "カタ", "ヒレ")
    ReDim partsNorm(LBound(partNames) To UBound(partNames))
    For i = LBound(partNames) To UBound(partNames)
        partsNorm(i) = NormKey(partNames(i))
    Next i

    tail = Left$(nameN, p - 1)
    Do
        ' 「⇒コンテナ」「・コンテナ」のように区切り記号が直前にある形に備えて、
        ' 部位を探す前にも区切りを落としておく
        Do While Len(tail) > 0
            If InStr(1, SEPS, Right$(tail, 1)) > 0 Then
                tail = Left$(tail, Len(tail) - 1)
            Else
                Exit Do
            End If
        Loop
        matched = False
        For i = LBound(partNames) To UBound(partNames)
            If Len(tail) >= Len(partsNorm(i)) Then
                If Right$(tail, Len(partsNorm(i))) = partsNorm(i) Then
                    parts = partNames(i) & IIf(Len(parts) > 0, "・" & parts, "")
                    tail = Left$(tail, Len(tail) - Len(partsNorm(i)))
                    matched = True
                    Do While Len(tail) > 0
                        If InStr(1, SEPS, Right$(tail, 1)) > 0 Then
                            tail = Left$(tail, Len(tail) - 1)
                        Else
                            Exit Do
                        End If
                    Loop
                    Exit For
                End If
            End If
        Next i
    Loop While matched

    If Len(parts) > 0 Then state = 1 Else state = 2
End Sub

' マスタ1件が当たるか調べる。当たった行番号を返す（0=なし）。
Private Function MatchRule(ByRef rules() As TRule, ByVal n As Long, _
                           ByVal nameN As String, ByVal cutN As String, ByVal farmN As String) As Long
    Dim i As Long, pat As String
    For i = 1 To n
        pat = "*" & rules(i).Keyword & "*"
        Select Case rules(i).Target
            Case "作業名称"
                If nameN Like pat Then MatchRule = i: Exit Function
            Case "カット内容"
                If cutN Like pat Then MatchRule = i: Exit Function
            Case "生産農場"
                If farmN Like pat Then MatchRule = i: Exit Function
            Case Else       ' 両方
                If nameN Like pat Or cutN Like pat Or farmN Like pat Then MatchRule = i: Exit Function
        End Select
    Next i
End Function

' 略称マスタ用（作業名称のみを見る）
Private Function MatchRuleName(ByRef rules() As TRule, ByVal n As Long, ByVal nameN As String) As Long
    Dim i As Long
    For i = 1 To n
        If nameN Like "*" & rules(i).Keyword & "*" Then MatchRuleName = i: Exit Function
    Next i
End Function

' ===========================================================================
'  出力
' ===========================================================================

' 数値のセルにだけシートの印（マスタ「■ シートの印」の記号）を頭に付ける（REQ-014）。
' TWork の内部の値（Ude/Momo/…/Heads）は数値のまま持ち、書き出す直前にここで初めて
' 文字列に変換する（内部の値と表示文字列を分ける。C3）。
'   ・mark が空欄 → 印を付けない（元の値のまま）
'   ・Empty（部位を使わない箱種）・"/"（部位コンテナで除外）→ 何もしない
'   ・"No.6" & 改行 & 数字（REQ-007）→ 改行の後ろの数字にだけ印を付ける
'   ・それ以外の数値 → 印＋数値の文字列にする
Private Function MarkedValue(ByVal raw As Variant, ByVal mark As String) As Variant
    Const NO6_PREFIX As String = "No.6" & vbLf
    Dim numPart As String

    If Len(mark) = 0 Then
        MarkedValue = raw
        Exit Function
    End If
    If IsEmpty(raw) Then
        MarkedValue = raw
        Exit Function
    End If
    If VarType(raw) = vbString Then
        If raw = EXCLUDED_MARK Then
            MarkedValue = raw
        ElseIf Left$(CStr(raw), Len(NO6_PREFIX)) = NO6_PREFIX Then
            numPart = Mid$(CStr(raw), Len(NO6_PREFIX) + 1)
            If numPart = "0" Then
                MarkedValue = raw       ' 0 には印を付けない（2026-09-28）
            Else
                MarkedValue = NO6_PREFIX & mark & numPart
            End If
        Else
            MarkedValue = raw           ' 想定外の文字列はそのまま（安全側）
        End If
    ElseIf raw = 0 Then
        MarkedValue = raw               ' 0 には印を付けない（頭数・部位とも。2026-09-28）
    Else
        MarkedValue = mark & CStr(raw)
    End If
End Function

Private Sub WriteFormatSheet(ByRef works() As TWork, ByVal nWorks As Long, _
        ByVal wYes As Long, ByVal wNo As Long, ByVal bYes As Long, ByVal bNo As Long, _
        ByVal totalHeads As Long, ByVal dateText As String, ByVal colsPerPage As Long)

    Dim ws As Worksheet
    Dim nCols As Long, i As Long, c As Long
    Dim out As Variant
    Dim lastCol As Long
    Dim rng As Range

    Set ws = GetSheet(FMT_WS)
    If ws Is Nothing Then Err.Raise 5, , "「" & FMT_WS & "」シートがありません。"
    If ws.ProtectContents Then
        Err.Raise 5, , "「" & FMT_WS & "」シートが保護されています。保護を解除してから実行してください。"
    End If

    ' --- 帳票エリアだけを毎回まっさらにする（行1～15のみ）---
    With ws.Range(ws.Cells(1, 1), ws.Cells(ROW_NOTE, CLEAR_COLS))
        .UnMerge                                   ' 結合が残っていると一括書き込みが失敗する（無い場合も無害）
        .ClearContents
        .ClearFormats
        .RowHeight = 15
    End With

    ' --- 列数 ---
    For i = 1 To nWorks
        If works(i).Result = "列を作成" Then nCols = nCols + 1
    Next i
    lastCol = FIRST_DATA_COL + nCols - 1
    If lastCol < SUMMARY_LAST_COL Then lastCol = SUMMARY_LAST_COL   ' 冒頭集計欄は必ず入る
    ' 紙と同じく、1ページ分（既定15列）までは空でも枠を引いておく
    If lastCol < FIRST_DATA_COL + colsPerPage - 1 Then lastCol = FIRST_DATA_COL + colsPerPage - 1

    ' --- 冒頭集計欄 ---
    ws.Cells(ROW_SUMMARY, 1).Value = "無地半"
    ws.Cells(ROW_SUMMARY, 2).Value = "あり"
    ws.Cells(ROW_SUMMARY, 3).Value = wYes
    ws.Cells(ROW_SUMMARY, 4).Value = "なし"
    ws.Cells(ROW_SUMMARY, 5).Value = wNo
    ws.Cells(ROW_SUMMARY, 6).Value = "黒半"
    ws.Cells(ROW_SUMMARY, 7).Value = "あり"
    ws.Cells(ROW_SUMMARY, 8).Value = bYes
    ws.Cells(ROW_SUMMARY, 9).Value = "なし"
    ws.Cells(ROW_SUMMARY, 10).Value = bNo
    ws.Cells(ROW_SUMMARY, 12).Value = dateText
    ws.Cells(ROW_TOTAL, 12).Value = "総頭数　" & totalHeads & " 頭"

    ' --- 行ラベル ---
    ws.Cells(ROW_ITEM + OFS_ITEM, 1).Value = "項　目"
    ws.Cells(ROW_ITEM + OFS_HEAD, 1).Value = "頭　数"
    ws.Cells(ROW_ITEM + OFS_BOX, 1).Value = "箱"
    ws.Cells(ROW_ITEM + OFS_UDE, 1).Value = "ウデ"
    ws.Cells(ROW_ITEM + OFS_MOMO, 1).Value = "モモ"
    ws.Cells(ROW_ITEM + OFS_ROSU, 1).Value = "ロース"
    ws.Cells(ROW_ITEM + OFS_BARA, 1).Value = "バラ"
    ws.Cells(ROW_ITEM + OFS_KATA, 1).Value = "カタ"
    ws.Cells(ROW_ITEM + OFS_HIRE, 1).Value = "ヒレ"
    ws.Cells(ROW_ITEM + OFS_CHIMA, 1).Value = "チマキ"
    ws.Cells(ROW_ITEM + OFS_NO9, 1).Value = "No.9"

    ' --- データ列（0件でも落ちないように分岐）---
    If nCols > 0 Then
        ReDim out(1 To ROW_COUNT, 1 To nCols)
        c = 0
        For i = 1 To nWorks
            If works(i).Result = "列を作成" Then
                c = c + 1
                out(OFS_ITEM + 1, c) = works(i).Abbr
                ' 頭数・部位の実数値（数値）にだけシートの印を付ける（REQ-014）。
                ' 頭数が読めなかった行（HeadsOK=False）は 0 のまま印を付けない。
                ' 部位（ウデ～ヒレ）は Empty／"/" のときは MarkedValue が何もしないので、
                ' ここでの分岐は不要（内部の値と書き出し時の表示文字列を分ける。C2・C3）
                If works(i).HeadsOK Then
                    out(OFS_HEAD + 1, c) = MarkedValue(works(i).Heads, works(i).SheetMark)
                Else
                    out(OFS_HEAD + 1, c) = works(i).Heads
                End If
                out(OFS_BOX + 1, c) = works(i).BoxKind
                out(OFS_UDE + 1, c) = MarkedValue(works(i).Ude, works(i).SheetMark)
                out(OFS_MOMO + 1, c) = MarkedValue(works(i).Momo, works(i).SheetMark)
                out(OFS_ROSU + 1, c) = MarkedValue(works(i).Rosu, works(i).SheetMark)
                out(OFS_BARA + 1, c) = MarkedValue(works(i).Bara, works(i).SheetMark)
                out(OFS_KATA + 1, c) = MarkedValue(works(i).Kata, works(i).SheetMark)
                out(OFS_HIRE + 1, c) = MarkedValue(works(i).Hire, works(i).SheetMark)
                out(OFS_CHIMA + 1, c) = Empty      ' 現場記入枠
                out(OFS_NO9 + 1, c) = Empty        ' 現場記入枠
            End If
        Next i
        ws.Range(ws.Cells(ROW_ITEM, FIRST_DATA_COL), _
                 ws.Cells(ROW_LAST, FIRST_DATA_COL + nCols - 1)).Value = out
    End If

    ' --- 書式 ---
    With ws.Range(ws.Cells(ROW_SUMMARY, 1), ws.Cells(ROW_SUMMARY, SUMMARY_LAST_COL))
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 12
        .ShrinkToFit = True
    End With
    ws.Range(ws.Cells(ROW_SUMMARY, 1), ws.Cells(ROW_SUMMARY, 1)).Font.Bold = True
    ws.Range(ws.Cells(ROW_SUMMARY, 6), ws.Cells(ROW_SUMMARY, 6)).Font.Bold = True
    With ws.Range(ws.Cells(ROW_SUMMARY, 12), ws.Cells(ROW_TOTAL, 12))
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 13
        .HorizontalAlignment = xlLeft
        .VerticalAlignment = xlCenter
    End With

    Set rng = ws.Range(ws.Cells(ROW_ITEM, 1), ws.Cells(ROW_LAST, lastCol))
    With rng
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
        .VerticalAlignment = xlCenter
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 14
        .ShrinkToFit = True
        .NumberFormat = "General"
    End With
    With ws.Range(ws.Cells(ROW_ITEM, 1), ws.Cells(ROW_LAST, 1))
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
    End With
    ' ヒレ行は「No.6」＋改行を出すので折り返す
    With ws.Range(ws.Cells(ROW_ITEM + OFS_HIRE, FIRST_DATA_COL), ws.Cells(ROW_ITEM + OFS_HIRE, lastCol))
        .ShrinkToFit = False
        .WrapText = True
        .Font.Size = 11
    End With
    ' 現場記入枠は薄く色を敷いて「書く場所」とわかるようにする
    With ws.Range(ws.Cells(ROW_ITEM + OFS_CHIMA, FIRST_DATA_COL), ws.Cells(ROW_ITEM + OFS_NO9, lastCol))
        .Interior.Color = RGB(252, 252, 245)
    End With

    ' --- 寸法 ---
    ws.Columns(1).ColumnWidth = 9.5
    ws.Range(ws.Columns(FIRST_DATA_COL), ws.Columns(lastCol)).ColumnWidth = 7.3
    ws.Rows(ROW_SUMMARY).RowHeight = 26
    ws.Rows(ROW_TOTAL).RowHeight = 22
    ws.Rows(ROW_ITEM & ":" & ROW_LAST).RowHeight = 32

    ' --- 印刷設定は環境（プリンタ有無・ドライバ）に左右されるので切り離す。
    '     ここで失敗しても表そのものは残す ---
    mPrintWarn = ApplyPrintSetup(ws, lastCol, colsPerPage)
End Sub

' 印刷範囲・向き・改ページを設定する。失敗したら警告文を返す（表は壊さない）。
Private Function ApplyPrintSetup(ByVal ws As Worksheet, ByVal lastCol As Long, _
                                 ByVal colsPerPage As Long) As String
    Dim brk As Long
    On Error GoTo Failed

    ws.ResetAllPageBreaks
    With ws.PageSetup
        .PrintArea = ws.Range(ws.Cells(1, 1), ws.Cells(ROW_LAST, lastCol)).Address
        .Orientation = xlLandscape
        .PaperSize = xlPaperA4
        .Zoom = 100                 ' 手動改ページを効かせるため FitToPages は使わない
        .PrintTitleColumns = "$A:$A"
        .PrintTitleRows = ""
        .LeftMargin = Application.InchesToPoints(0.3)
        .RightMargin = Application.InchesToPoints(0.3)
        .TopMargin = Application.InchesToPoints(0.4)
        .BottomMargin = Application.InchesToPoints(0.4)
        .CenterHorizontally = False
    End With

    ' データ列を colsPerPage 件ずつに区切る（次ページ先頭列の左に入れる）
    brk = FIRST_DATA_COL + colsPerPage
    Do While brk <= lastCol
        ws.VPageBreaks.Add Before:=ws.Columns(brk)   ' 縦（列）方向の改ページ。名前は VPageBreaks（Vertical... ではない）
        brk = brk + colsPerPage
    Loop
    Exit Function

Failed:
    ApplyPrintSetup = "印刷設定（用紙・向き・改ページ）を設定できませんでした：" & Err.Description & _
                      "　表の中身は作れています。印刷前に印刷プレビューで確認してください。"
End Function

Private Sub WriteNoteOnFormat(ByVal warnCount As Long, ByVal extraMsg As String)
    Dim ws As Worksheet
    Set ws = GetSheet(FMT_WS)
    If ws Is Nothing Then Exit Sub
    With ws.Cells(ROW_NOTE, 1)
        If warnCount > 0 Then
            .Value = "▲ 確認してほしいことが " & warnCount & " 件あります。「判定ログ」シートの「警告」の列（T列）を見てください。" & _
                     IIf(Len(extraMsg) > 0, "　【" & extraMsg & "】", "")
            .Font.Color = RGB(192, 0, 0)
        Else
            .Value = "確認が必要な項目はありませんでした（" & Format$(Now, "yyyy/mm/dd hh:nn") & " 作成）"
            .Font.Color = RGB(89, 89, 89)
        End If
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 11
        .Font.Bold = (warnCount > 0)
    End With
End Sub

' 判定ログの見出し（1行目）。build_format.py の LOG_HEADERS と文言・順序を必ず一致させる
' （check_consistency.py が突き合わせる）。先方の今のブック（9/11版など）は20列の見出ししか
' 無いため、REQ-014 の U・V 列（シート（あり/なし）／シートの根拠）が欠けたままになる。
' 毎回この関数で書き直すことで、古いブックでも不足分の見出しが必ず入るようにする（REQ-014）。
Private Function LogHeaderLabels() As Variant
    LogHeaderLabels = Array("入力行", "順番", "作業№", "作業名称", "頭数", "半箱欄(H)", "カット内容", "生産農場", _
        "処理結果", "判定した理由", "箱種", "箱種の根拠", "全農判定", "全農判定の根拠", _
        "除外した部位", "略称", "略称のマッチ元", "出力列", "要目視確認", "警告", _
        "シート（あり/なし）", "シートの根拠")
End Function

Private Sub WriteLogSheet(ByRef works() As TWork, ByVal nWorks As Long, ByVal warnings As String)
    Dim ws As Worksheet
    Dim out As Variant
    Dim i As Long
    Dim headers As Variant, j As Long
    Const NCOL As Long = 22   ' 20列＋シート（あり/なし）・シートの根拠（REQ-014）

    Set ws = GetSheet(LOG_WS)
    If ws Is Nothing Then Exit Sub

    ' --- 見出し（1行目）は毎回書き直す（既存ブックに不足している列を必ず埋めるため）---
    headers = LogHeaderLabels()
    For j = LBound(headers) To UBound(headers)
        ws.Cells(1, j + 1).Value = headers(j)
    Next j
    With ws.Range(ws.Cells(1, 1), ws.Cells(1, NCOL))
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 11
        .Font.Bold = True
        .Interior.Color = RGB(242, 242, 242)
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlThin
        .HorizontalAlignment = xlCenter
    End With

    If ws.Cells(ws.Rows.Count, 1).End(xlUp).Row > 1 Then
        ws.Range(ws.Cells(2, 1), ws.Cells(ws.Cells(ws.Rows.Count, 1).End(xlUp).Row, NCOL)).Clear
    End If
    If nWorks = 0 Then Exit Sub

    ReDim out(1 To nWorks, 1 To NCOL)
    For i = 1 To nWorks
        out(i, 1) = works(i).SrcRow
        out(i, 2) = works(i).OrderNo
        out(i, 3) = works(i).WorkNo
        out(i, 4) = works(i).WorkName
        out(i, 5) = works(i).Heads
        out(i, 6) = works(i).HalfRaw
        out(i, 7) = works(i).CutName
        out(i, 8) = works(i).FarmName
        out(i, 9) = works(i).Result
        out(i, 10) = works(i).Reason
        out(i, 11) = works(i).BoxKind
        out(i, 12) = works(i).BoxReason
        out(i, 13) = works(i).Zennou
        out(i, 14) = works(i).ZennouWhy
        out(i, 15) = works(i).ExcludeParts
        out(i, 16) = works(i).Abbr
        out(i, 17) = works(i).AbbrWhy
        out(i, 18) = IIf(works(i).OutCol > 0, ColLetter(works(i).OutCol), "")
        out(i, 19) = works(i).NeedCheck
        out(i, 20) = works(i).Warn
        out(i, 21) = works(i).SheetYesNo
        out(i, 22) = works(i).SheetWhy
    Next i
    ws.Range(ws.Cells(2, 1), ws.Cells(1 + nWorks, NCOL)).Value = out

    With ws.Range(ws.Cells(2, 1), ws.Cells(1 + nWorks, NCOL))
        .Font.Name = "ＭＳ Ｐゴシック"
        .Font.Size = 10
        .VerticalAlignment = xlTop
        .Borders.LineStyle = xlContinuous
        .Borders.Weight = xlHairline
    End With
    With ws.Range(ws.Cells(2, 19), ws.Cells(1 + nWorks, 20))
        .Font.Color = RGB(192, 0, 0)
        .WrapText = True
    End With
End Sub

Private Sub CollectRowWarnings(ByRef works() As TWork, ByVal nWorks As Long, ByRef warnCount As Long)
    Dim i As Long
    For i = 1 To nWorks
        If Len(works(i).Warn) > 0 Then
            warnCount = warnCount + 1
        End If
    Next i
End Sub

Private Function CountResult(ByRef works() As TWork, ByVal nWorks As Long, ByVal kind As String) As Long
    Dim i As Long
    For i = 1 To nWorks
        If works(i).Result = kind Then CountResult = CountResult + 1
    Next i
End Function

' ===========================================================================
'  小道具
' ===========================================================================

' 比較用の文字列に整える。全角英数・全角カナを半角にし、空白を落とす。
' ※ H列の区切り解析には使わない（区切りが消えるため）
Private Function NormKey(ByVal v As Variant) As String
    Dim s As String
    NormKey = ""
    If IsError(v) Then Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    s = CStr(v)
    s = NarrowSafe(s)
    s = Replace(s, ChrW(&H3000), "")
    s = Replace(s, " ", "")
    s = Replace(s, vbTab, "")
    s = Replace(s, vbCr, "")
    s = Replace(s, vbLf, "")
    NormKey = UCase$(s)          ' 大小文字の差もここで吸収する（spf と SPF を同じに扱う）
End Function

' StrConv は環境によって失敗しうるので、失敗したら元の文字列を返す
Private Function NarrowSafe(ByVal s As String) As String
    Dim t As String
    On Error Resume Next
    t = StrConv(s, vbNarrow)
    If Err.Number <> 0 Then Err.Clear: t = s
    On Error GoTo 0
    If Len(t) = 0 And Len(s) > 0 Then t = s
    NarrowSafe = t
End Function

' Trim$ は半角スペース（&H20）しか削らないため、全角スペース（&H3000）も前後から削る版。
' 「シートの印」の印（マスタ B列）のように、全角スペースが混ざりやすい入力の検査に使う（REQ-014）
Private Function TrimWide(ByVal s As String) As String
    Dim t As String
    t = s
    Do While Len(t) > 0
        Select Case Left$(t, 1)
            Case " ", ChrW(&H3000)
                t = Mid$(t, 2)
            Case Else
                Exit Do
        End Select
    Loop
    Do While Len(t) > 0
        Select Case Right$(t, 1)
            Case " ", ChrW(&H3000)
                t = Left$(t, Len(t) - 1)
            Case Else
                Exit Do
        End Select
    Loop
    TrimWide = t
End Function

Private Function SafeStr(ByVal v As Variant) As String
    SafeStr = ""
    If IsError(v) Then SafeStr = "#エラー": Exit Function
    If IsNull(v) Then Exit Function
    If IsEmpty(v) Then Exit Function
    SafeStr = CStr(v)
End Function

Private Function AddSep(ByVal baseTxt As String, ByVal addTxt As String) As String
    If Len(baseTxt) = 0 Then AddSep = addTxt Else AddSep = baseTxt & " / " & addTxt
End Function

Private Sub AppendWarn(ByRef warnings As String, ByRef warnCount As Long, ByVal msg As String)
    warnings = AddSep(warnings, msg)
    warnCount = warnCount + 1
End Sub

' 「1頭セット条件」表の自動追加・追加失敗の通知（mMasterNote）を、途中で中断する
' メッセージにも必ず添える（マスタは既に追加できていたのに、後続のマスタ不備や
' 判定失敗で止まって利用者に伝わらない、を防ぐ。既に mMasterNote が空なら何もしない）
Private Function AppendMasterNote(ByVal baseMsg As String) As String
    If Len(mMasterNote) > 0 Then
        AppendMasterNote = baseMsg & vbCrLf & vbCrLf & mMasterNote
    Else
        AppendMasterNote = baseMsg
    End If
End Function

' mMasterNote に1行追記する（REQ-014）。「1頭セット条件」「シートの印」の両方の表を
' 1回の実行で追加した場合など、複数の通知を1本にまとめて出すため、必ずこれを使う
' （mMasterNote = "…" と直接代入すると、先に積んだ通知を消してしまう）
Private Sub NoteMaster(ByVal msg As String)
    If Len(mMasterNote) > 0 Then
        mMasterNote = mMasterNote & vbCrLf & msg
    Else
        mMasterNote = msg
    End If
End Sub

Private Function GetSheet(ByVal nm As String) As Worksheet
    Dim ws As Worksheet
    For Each ws In ThisWorkbook.Worksheets
        If ws.Name = nm Then Set GetSheet = ws: Exit Function
    Next ws
End Function

' 列番号を A, B, ... AA のような列名にする（ActiveSheet に依存させない）
Private Function ColLetter(ByVal c As Long) As String
    Dim n As Long, s As String, r As Long
    n = c
    Do While n > 0
        r = (n - 1) Mod 26
        s = Chr$(65 + r) & s
        n = (n - 1 - r) \ 26
    Loop
    ColLetter = s
End Function

' TWork を初期化する（Type は Erase できないので詰め替える）
Private Sub Erase2(ByRef w As TWork)
    Dim blank As TWork
    w = blank
End Sub
