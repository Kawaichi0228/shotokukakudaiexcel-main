# 目次

- [インデント](#indentation)
- [コメント](#comment)
- [Public及びPrivate](#Public-and-Private)
- [Public及びPrivate - Sub,Function](#PubPri-SubFunction)
- [Public及びPrivate - Variable,Const,Enum](#PubPri-VarConsEnum)
- [宣言](#declaration)
- [宣言 - Const](#declaration-Const)
- [宣言 - Class - インスタンス化](#declaration-Class-new)
- [宣言 - Class - インスタンスの破棄](#declaration-Class-nothing)
- [宣言 - Variant](#declaration-Variant)
- [宣言 - 論理値](#declaration-boolean)
- [宣言 - 数値](#declaration-number)
- [宣言 - 配列](#declaration-array)
- [グローバル変数](#global-variables)
- [ルーチンの作成方法](#how-to-routin)
- [ルーチンの呼出](#call-routin)
- [制御構造](#control-structure)
- [制御構造 - 論理値の判定](#boolean-expression)
- [制御構造 - 論理値以外の判定](#notboolean-expression)
- [エラーハンドリング](#error-handling)
- [行結合(行の改行)](#state-join)

***

# コーディング規約

<a name="indentation"></a>

## インデント

- 1レベルのインデントに4つの半角スペースを使用する。
- Try文 (`On Error` ステートを含む行) はインデントゼロとする。
- 式の途中で改行する場合は、2行目以降を1行目より1段深くインデントする。

    ```vbnet
    'example1
    Dim filePath As Variant
    filePath = Application.GetSaveAsFilename( _
      FileFilter:=extensionName & "ファイル" & "," & "*" & extensionName & "," & "全てのファイル,*.*", _
      FilterIndex:=1, _
      InitialFileName:=defaultBaseName & extensionName, _
      Title:="名前を付けて保存")

    'example2
    Private Function CreateButton_GotoWs_給与等データ入力(toTargetWs As Worksheet, _
      Optional Margin_left As Long)
    ```

<a name="comment"></a>

## コメント

- 分割線としてコメントを用いたい場合は、以下の2種類のみを使用するものとする。<br>

  ```vbnet
  'Short
  '/******************************************************
  '*******************************************************

  '*******************************************************
  '******************************************************/

  'Long
  '/**********************************************************************************************
  '***********************************************************************************************

  '***********************************************************************************************
  '**********************************************************************************************/
  ```

- なお、行を増やすためのスラッシュを含まないコメントについては省略することができる。

  ```vbnet
  'example (行を増やすためのスラッシュを含まないコメントは省略可)
  '/******************************************************
  'example用のルーチン
  Sub FooRoutin()
    
  End Sub
  '******************************************************/
  ```

<a name="Public-and-Private"></a>

## Public及びPrivate

<a name="PubPri-SubFunction"></a>

### Sub,Function
- プロジェクトレベルの `Sub` 及び `Function` は、`Public` を省略する。

  **[理由]** 文字数の差によるコードの横位置関係により、視覚的に `Public` と `Private` を区別できるため

  ```vbnet
  'good
  Sub FooRoutin_Parent()
      
  End Sub

  Private Sub FooRoutin_Child()
      
  End Sub

  'bad
  Public Sub FooRoutin_Parent()
      
  End Sub

  Private Sub FooRoutin_Child()
      
  End Sub
  ```

<a name="PubPri-VarConsEnum"></a>

### Variable,Const,Enum
- モジュールレベルの宣言フィールド において、モジュールレベル変数 及び 定数 及び 列挙体 を宣言する場合、グローバルの場合は `Public` 及び `Public Const` 及び `Public Enum` を使用し、プライベートの場合は、`Private` 及び `Private Const` 及び `Private Enum` を使用する。

  **[理由]** プロシージャレベル変数 及び 定数との曖昧さ回避のため

  ```vbnet
  'good
  Public fooVar_Public As Long
  Private fooVar_Private As Long
  Public Const bar_Public = "string"
  Private Const bar_Private = "string"

  Public Enum baz_Public
    num = 1
  End Enum

  Private Enum baz_Private
    num = 2
  End Enum
  
  Sub FooRoutin()

  End Sub

  'bad
  Dim fooVar As Long '(Privateと同義)
  Const bar = "string" '(Private Constと同義)

  Sub FooRoutin()

  End Sub
  ```

<a name="declaration"></a>

## 宣言

<a name="declaration-Const"></a>

### Const
- 定数の宣言時の型指定は省略する。

  ```vbnet
  'good
  Const foo = 100
  Const bar = "string"

  'bad
  Const foo As Long = 100
  Const bar As String = "string"
  ```

<a name="declaration-Class-new"></a>

### Class - インスタンス化
- クラスの宣言時の `As New` を禁止とする。<br>
また、インスタンス化(`Set` `New`ステート)は特に理由のない限り、1行で記述するものとする。

  **[理由]** Visual Basic 6.0において、MS公式が非推奨としているため<br>

  > (参考 - MS公式)<br>
  https://docs.microsoft.com/ja-jp/previous-versions/technical-document/dd297716(v=msdn.10)?redirectedfrom=MSDN

  ```vbnet
  'good
  Dim c as clsFoo: Set c = New clsFoo 'インスタンス化

  'bad
  Dim c as New clsFoo '(cが参照されるまで、インスタンスは生成されない)
  ```

<a name="declaration-Class-nothing"></a>

### Class - インスタンスの破棄
- COM オブジェクト(`CreateObject`)の場合は、必ず `Nothing` で破棄する。

  **[理由]** メモリを確実に解放し、不具合を防止するため
  ```vbnet
  Sub FooRoutin()
      Dim cn As Object
      Set cn = CreateObject("ADODB.Connection")
      '***Process***
      Set cn = Nothing 'インスタンスの破棄
  End Sub
  ```

- ユーザ作成のクラスは破棄を省略することができる。

  **[理由]** ガベージコレクションが存在するため(ルーチンから抜けた時点で自動的に破棄)

  ```vbnet
  Sub FooRoutin()
      Dim c as clsFoo: Set c = New clsFoo
      '(ユーザ作成のクラスはインスタンスの破棄を省略可能)
  End Sub
  ```

<a name="declaration-Variant"></a>

### Variant
- 原則は、省略せずに型を宣言するものとする。

  **[理由]** 型宣言漏れなのか、区別できなくなるため

  ```vbnet
  'good
  Dim foo As Variant

  'bad
  Dim foo '(省略時はVariant型)
  ```

<a name="declaration-boolean"></a>

### 論理値
- **[VBA]** `True` 及び `False` を使うものとする。

  **[理由]** コードを見た際、Boolean値 or 数値なのか、判断が難しくなるため

  ```vbnet
  Dim foo As Boolean

  'good
  foo = True

  'bad
  foo = 1
  ```

- **[Workbook]** `True` 及び `False` のほか、`1` 及び `0` を使うことができるものとする。

  **[理由]** 短く記述することで、少ない列幅で可読することができるため

<a name="declaration-number"></a>

### 数値
- 整数 は `As Long` を使う。
- 小数点 は `As Double` を使う。

  **[理由]** 予期せぬエラー回避のため

  ```vbnet
  'good
  Dim foo As Long
  Dim bar As Double
  foo = 100
  bar = 1.5

  'bad
  Dim foo As Integer
  Dim bar As Single
  foo = 100
  bar = 1.5
  ```

<a name="declaration-array"></a>

### 配列
- indexの開始値は、特段の理由が無い限り `1` とする。

  **[理由]** 他のオブジェクト(`Collection` , `Range`等)の多くが開始値 `1` であり、統一性を持たせるため。

  ```vbnet
  'good
  Dim fooAry(1 To 3) As String

  'bad
  Dim fooAry(2) As String
  ```

<a name="global-variables"></a>

## グローバル変数

- グローバル(`Public`)変数は原則、使用禁止とする。<br>
※ただし、クラスのPublicメンバ変数を除く

<a name="how-to-routin"></a>

## ルーチンの作成方法
- データ結合を積極的に使う。定義→処理実行ルーチン と 処理ルーチン に分割すること。
  ```vbnet
  Sub MainRoutin
      Dim wsFoo As Worksheet, bar As Long
      Set wsFoo = Sheet1
      bar = 100
      ChildRoutin wsFoo , bar
  End Sub
  ```

- 制御結合は可能な限り行わないこと。<br>
例えば、複数のシートが使用する処理で、基本的にはシート全体で共通処理が多いが一部処理が違うルーチンを作成したい場合は、<u>**Worksheet単位で実行できるように**</u> 分割し作成すること。<br>
  ```vbnet
  Sub FooRoutin_Sheet1()
      Const foo = "String"
      Dim rng As Range
      Set rng = Sheet1.Cells(1,5)
      FooRoutin foo, rng
  End Sub
  ```

- また、複数のシート処理をまとめて実行させたい場合は、別のルーチンとして作成すること。
  ```vbnet
  Sub FooRoutin_ALLSheet()
      FooRoutin_Sheet1
      FooRoutin_Sheet2
      FooRoutin_Sheet3
  End Sub
  ```

- 複数のオブジェクトに処理が必要なルーチンを作成する際は、単一のオブジェクトでも動作するルーチンを作った上で、それを呼び出すルーチンを別途作成し、引数には配列を渡すこと。

  **[理由]** Initialize処理やFinally処理が1度で済み、処理速度が向上するため

  ```vbnet
  Private Sub FooRoutin(wsAry() As Worksheet)
      'Initialize
      Application.ScreenUpdating = False

      'MainProcess
      Dim i As Long
      For i = LBound(wsAry, 1) To UBound(wsAry, 1)
          ProcessRoutin wsAry(i)
      Next

      'Finally
      Application.ScreenUpdating = True
  End Sub
  ```

<a name="call-routin"></a>

## ルーチンの呼出
- `Call` による呼出を禁止とする。

  **[理由]** コードを冗長化させないため

  ```vbnet
  'good
  FooRoutin bar, baz

  'bad
  Call FooRoutin(bar, baz)
  ```

<a name="control-structure"></a>

## 制御構造

<a name="boolean-expression"></a>

### 論理値の判定

- 論理値の判定において、`True` と `False` は省略して記述する。

  **[理由]** 可読性のため

  ```vbnet
  'good
  If var Then
  If Not var Then

  'bad
  If var = True Then
  If var = False Then
  ```

<a name="notboolean-expression"></a>

### 論理値以外の判定

- 論理値以外の判定において、否定演算子を使用する場合は、原則、`Not` ではなく `<>` を使う。ただし、例外として `<>` が使用できないステートの場合は使用可能とする。

  **[理由]** 読み間違いによるヒューマンエラー防止のため

  ```vbnet
  'good
  If foo <> 100 Then bar = "string"
  If Not foo Is Nothing Then bar = "string"

  'bad
  If Not foo = 100 Then bar = "string"
  ```

<a name="error-handling"></a>

## エラーハンドリング
- VBAには Try-Catch が存在しないため、独自にコードを作成している。<br>
エラーハンドリングを行う際は、必ず以下のコードをコピーして使うこと。

  **[仕様]** Sheet上の指定セルに書き込まれたBoolean値を取得し、`True` なら `On Error GoTo` 処理を行う。

  ```vbnet
  Try:
  Dim isDebugMode_ErrorHandling As Boolean
  isDebugMode_ErrorHandling = GetStaticVariable_fromWorksheet(Ws_cfg.Range("cfg_sca_isDebugMode_ErrorHandling"))
  If isDebugMode_ErrorHandling Then On Error GoTo Catch

  Finally:
      Exit Sub '(※1) 最上位モジュールの場合    

  Catch:
      'エラー処理
      GoTo Finally
  ```

- (※1) なお、下位モジュールの場合で、エラー情報を上位モジュールへThrowしたい場合は、<br>
`Exit Sub` の部分を下記に書き換えること。

  ```vbnet
  '下位のモジュールの場合
  'エラーが発生していた際はエラーを再度起こし、上位モジュールのOn Error処理へThrowさせる
  If Err.number <> 0 Then Err.Raise Err.number, Err.source, Err.description, Err.HelpFile, Err.HelpContext
  ```

- また、上位モジュールへThrowさせずに、モジュール内で完結する意図したエラー処理を行いたい場合には、
下記のコードを用いること。

  ```vbnet
  'Try
  On Error Resume Next

      'Process
      FooRoutin

      'Catch
      If Err.number = 0 Then
          ret = True

      ElseIf Err.number = 100 Then
          '**ここにエラー処理**
          Err.Clear
          ret = False

      End If
  ```

<a name="state-join"></a>

## 行結合(行の改行)

- 行を改行する場合、 論理演算子 または 文字列連結演算子 の直後に行う。

  ```vbnet
  '制御文
  If Range("A1").Value = 100 And _
      Range("B1").Value = 200 Then
  End If

  '文字列の連結
  Cells(1, 6).Value = Cells(1, 1).Value & _
      Cells(1, 2).Value & _
      Cells(1, 3).Value & _
      Cells(1, 4).Value & _
      Cells(1, 5).Value
  ```