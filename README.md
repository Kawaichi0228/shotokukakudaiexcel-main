# 所得拡大促進税制計算エクセルソフト

<img src="logo-promotion.png" width="600px">

所得拡大促進税制計算エクセルソフトは、法人税を税額控除できる制度としてメジャーである 「所得拡大促進税制」 の計算を、簡単に・スピーディーに行うことができる本格的なエクセルマクロです。

※ソースコード非公開

***

## 主要機能

### 税務書類の出力機能(自動計算)

<img src="images/gif/beppyo.gif" width="500px">

複雑な税金の計算を、各シートに定められたフィールドへ入力することで、簡単に自動計算できるようにしました。また、計算した結果を税務書類様式へ反映させ、それを印刷できるように実装しました。

### 明細データを印刷・CSV・PDFの3種類の中から選択し、出力できる機能

<img src="images/gif/export.gif" width="500px">

ユーザが入力した各従業員の給与等データを明細として出力できる機能を実装しました。出力タイプは印刷、CSV、PDFの3種類から選ぶことができます。

### データベース機能(MySQL)

<img src="images/gif/database.gif" width="500px">

MySQLへ接続(ADODB使用)し、ExcelVBAからデータベースを操作できるライブラリを自作しました。また、データベース上に登録されたデータを表示し、選択した事業者のデータを各入力フィールドへインポートできる機能を実装しました。(※ユーザ提供版へは未実装)

***

## 補助機能

### テーブルの行追加・削除

<img src="images/gif/rowsinsertdelete.gif" width="500px">

テーブルの行を追加及び削除できる機能を実装しました。追加(削除)時に、一時的にシートの保護を解除することで実現させています。また、テーブル全体の各行ごとに色を分け(条件付き書式を使用)、見やすさを向上させました。

### 各入力列を用途のグループごとに表示切り替え

<img src="images/gif/rowscolumnstoggle.gif" width="500px">

必要な情報だけを見やすく、かつ入力しやすいように、ボタンを押すことで、グループごとの列の表示・非表示を切り替えできるようにしました。

### バリデーション機能

<img src="images/gif/validation.gif" width="500px">

任意のタイミングでユーザが入力した値をチェックし、正しい値でなければエラーメッセージを表示するバリデーション機能を実装しました。
- 出力処理の実行時
- セルの入力時(コピー&ペーストを含む)

### 入力フィールド初期化機能

<img src="images/gif/initialize.gif" width="500px">

ユーザが入力可能なフィールドを全て初期化できる機能を実装しました。(※ユーザ提供版へは未実装)

***

## その他

### コーディング規約等

markdownを使い、コーディング規約、命名規則等を作成しました。

[コーディング規約](styleguide/README.md)

### ER図・ユースケース

MySQL Workbench及びGitMindを使い、データベースのER図, ユースケースを作成しました。

<img src="images\erd.png" width="500px">
<img src="images\usecase.png" width="500px">

***

## 著作権

全て Kawaichi0228 に帰属する。

***

## 公式サイト

https://sites.google.com/view/shotokukakudai-excel
