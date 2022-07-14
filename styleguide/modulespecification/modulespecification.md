# 目次
- [メインモジュール(最上位モジュール)](#main-module)
- [クラスモジュール](#class-module)

***

# 各モジュールの仕様

- 最上位のメインモジュールでは、必ず Try-Catch(独自コード) を記述し、エラーハンドリングを行うこと。

<a name="main-module"></a>

## メインモジュール(最上位モジュール)

|モジュール名|説明|
|:-----|:-----|
|M_Main_WorkbookEvent|Workbook_OpenなどのThisWorkbookイベント用モジュール。<br />ブックのイベント処理の記述を、この標準モジュールに統一させる。<br />なお、ThisWorkbookモジュールへは既定のコードを記述するのみでよい。(※1)|
|M_Main_WorksheetEvent|Worksheet_ChangeなどのWorksheetイベント用モジュール。<br />各シートのイベント処理の記述を、この標準モジュールに統一させる。<br />なお、Worksheetモジュールへは既定のコードを記述するのみでよい。(※2)|
|M_Main_ButtonModule|フォームコントロールボタンに登録することでのみ使用できるメインモジュール。<br />登録したボタンオブジェクト名を取得し、分岐処理を行う。|

- (※1) ThisWorkbookモジュールへの記述用コード (Workbook_Openでの例)
	```vbnet
	Private Sub Workbook_Open()
		Main_Workbook_Open
	End Sub
	```

- (※2) Worksheetモジュールへの記述用コード (Worksheet_Changeでの例)
	```vbnet
	Private Sub Worksheet_Change(ByVal target As Range)
		Main_Worksheet_Change Me, target
	End Sub
	```

<a name="class-module"></a>

## クラスモジュール

- 以下のクラスを用いてプロジェクト用に独自に定めたい場合には、<u>必ず "**C_Pj_**" とprefixした別途のモジュールを作成</u>し、そこに記述すること。
"**C_**" とのみprefixされているモジュールのコードを、直接改変することは禁止とする。

	**[example]**<br>
	C_Pj_Button<br>

|クラス名|説明|
|:-----|:-----|
|C_Button|ボタン(フォームコントロール)を生成|
|C_Color|書式に使用する色の設定|
|C_Display|画面表示に関する設定(表示倍率、シート選択バーの表示・非表示等)|
|C_Hidden|行列の表示・非表示|
|C_Initialize|各処理の開始時及び終了時に行う、動作高速化等の処理を集約したもの|
|C_Lock|シートのロック(保護)・解除|
|C_Pj_CalcWs|各処理の中で一時的に使用するシートを設定(処理中以外はユーザへ不可視状態)|
|C_Pj_VisibleWs|ユーザへの可視状態を許可するシートを設定|
|C_Position|初期化時の選択位置(シート、セル等)の設定|
|C_UserForm_Msg|ユーザフォーム "UF_UserForm_Msg" に関する全ての設定|
|C_Format|通常の書式及び条件付き書式の設定|
|C_Fomrat_Conditions|主に、条件付き書式用の条件関数を生成し取得する関数を集約したもの|
|C_Size|行列等のサイズの設定|
|C_Validation|セルの入力規則の設定|