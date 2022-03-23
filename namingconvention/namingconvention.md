# 目次

- [全体共通](#common)
- [VBA - オブジェクト(モジュール)](#object-name)
- [VBA - ルーチン(プロシージャ・関数)](#routin-name)
- [VBA - 定数](#const-name)
- [VBA - 変数](#variable-name)
- [VBA - 列挙型(Enum)](#enum-name)
- [VBA - クラスのメンバ変数](#class-member-name)
- [VBA - 任意の順番でソートしたい場合](#sort)
- [Workbook - 名前の定義セル](#book-nameRange)
- [Workbook - テーブル名 及び テーブル各見出し名](#table-name)

***

# 命名規則

<a name="common"></a>

## 全体共通
- 使用する単語が複数となる場合、単語を アンダースコア `_` で区切る **スネークケース** を用いる。

    **[理由]** インテリセンスが利用しやすい かつ 可読性向上のため

    ```vbnet
    Function FooFunction_bar_baz() As Long
        Dim fooVariable_bar As Long
        fooVariable_bar = 100
        FooFunction_bar_baz = fooVariable_bar + 200
    End Function
    ```

<a name="object-name"></a>

## オブジェクト(モジュール)
- 以下のとおり、ハンガリアン記法でprefixする。
    - #### [Worksheet モジュール]
        **Ws**_[WorksheetName](#worksheet-objectname)

        ※VBAの処理の中で1度でも使用するシートについては、必ずこれに従い命名すること。<br>
        なお、特に処理において使用しないものについては、初期名称から変更しなくても良いものとする。

    - #### [UserForm モジュール]
        **UF**_Foo

    - #### [Library モジュール (共通ライブラリ)]
        **L**_Foo

    - #### [Class モジュール]
        **C**_Foo

    - #### [Class モジュール (Projectの仕様に合わせて別途作成したもの)]
        **C**_**Pj**_Foo

<a name="worksheet-objectname"></a>

#### WorksheetName
- 対象となるワークシートのタイトルを、単語ごとの頭文字をとった形や一般的な省略形、先頭の一部をローマ字等で省略したうえで命名する。**原則、<u>命名後の変更は不可</u> とする。** そのため、慎重に決めること。

    **[理由]** 使用箇所が非常に多く(関数名、変数名、名前の定義セル等)、変更時に膨大な時間コストを費やすこととなるため

    **[example]**<br>
    Config → <u>cfg</u><br>
    基本情報入力 → <u>kihon</u>

<a name="routin-name"></a>

## ルーチン(プロシージャ・関数)
- 単語ごとの頭文字を大文字とした **パスカルケース** を採用する。

    **[理由]** 慣習による

    ```vbnet
    Sub FooProcedure()
    End Sub

    Function FooFunction() As Variant
        FooFunction = Val
    End Function
    ```

- 最上位プロシージャには **Main_** をprefixする。(ただし、必要な場合はAliasを許可する)

    **[理由]** 可読性のため

    ```vbnet
    Sub Main_FooRoutin()
    End Sub

    '必要に応じて、Aliasも許可
    Sub Bar_AliasRoutin()
        Main_FooRoutin
    End Sub
    ```

<a name="const-name"></a>

## 定数
- 1つ目の単語の頭文字を小文字とし、2つ目以降の単語の頭文字を大文字とする **キャメルケース** を採用する。

    **[理由]** 慣習ではアッパースネークケース(ex.CONST_FOO)が多いが、定数から変数(またはその逆)に変更となった際に修正が容易なキャメルケースを採用する

    ```vbnet
    Const fooConst = ""
    ```

<a name="variable-name"></a>

## 変数
- 1つ目の単語の頭文字を小文字とし、2つ目以降の単語の頭文字を大文字とする **キャメルケース** を採用する。

    **[理由]** 慣習による

    ```vbnet
    Dim fooVariable As String
    ```

<a name="enum-name"></a>

## 列挙型(Enum)
- 単語ごとの頭文字を大文字とした **パスカルケース** を採用する。

    **[理由]** 慣習による

    ```vbnet
    Enum FontColor
        Red = 3
        Blue = 6
    End Enum
    ```

<a name="class-member-name"></a>

## クラスのメンバ変数
- 単語ごとの頭文字を大文字とした **パスカルケース** を採用する。

    **[理由]** 慣習による

- また、Private変数の場合には `_` をsuffixする。

    **[理由]** 慣習はprefixであるが(ex. _ClassVar)、VBAの仕様上、先頭にアンダースコアをつけることができないため、suffixとする。

    ```vbnet
    Public Age As String
    Private Name_ As String

    Property Get Name() As String
        Name = Name_: End Property

    Property Let Name(nm As String)
        Name_ = nm: End Property
    ```

<a name="sort"></a>

## 任意の順番でソートしたい場合
- VBAでは、プロジェクトウィンドウ、インテリセンス等は強制的に昇順でソートされる仕様となっている。<br>
任意の順番でソートしたい場合は、下記のとおり アルファベット & Number で名前を定義すること。<br>
また、そのNumberは必ず `1` スタートとすること。<br>

    **[理由]** `Collcetion` や `Cells` 等の index との統一性を持たせるため

### オブジェクト
- <u>Object省略名</u> & <u>2桁のNumber</u> & <u>2桁のNumber</u> を用いてソートする。<br>
なお、上記に加えて、"_" & 1桁のNumber を用いて、モジュール構造(入れ子)等を表現することも可能とする。<br>
    **[example]**<br>
    M0101_Bar<br>
    M0102_1_Baz<br>
    M0102_2_Baz

### クラスのメンバ変数
- <u>"p"</u> & <u>1桁のNumber</u> をprefixする。(p: propertyの省略形)<br>
    **[example]**<br>
    p1_ClassProperty


<a name="book-nameRange"></a>

## 名前の定義セル
- 必ず **全て 小文字 and 半角英数字** を用いること。<br>
    なお、Foo以降はこの制限なく、自由に記述することができるものとする。<br>
    - #### [名前の定義セル] (ex.kyuyo_head_氏名)<br>
        **worksheetname**_**rangetype**_Foo

#### WorksheetName
- [(オブジェクト名 の WorksheetName を参照)](#worksheet-objectname)

#### RangeType
- 参照範囲及び用途によって、以下のとおり記述すること。
    - #### [参照範囲が1つのセル] (ex.A1)
        foo_**sca**_Bar
    - #### [参照範囲が2つ以上のセル] (ex.A1:A5)
        foo_**mul**_Bar
    - #### [参照範囲が1つ かつ テーブル見出しのセル] (ex.テーブル[[#見出し],[列1]])
        foo_**head**_Bar
    - #### [参照範囲が1列 かつ テーブルデータのセル] (ex.テーブル[[#データ],[列1]] or テーブル[列1])
        foo_**body**_Bar
    - #### [配列用のセル]
        foo_**ary**_Bar
    - #### [データの入力規則用のセル]
        foo_**list**_Bar

    **[理由]**
    コーディング時に名前の定義セルを用いる際の記述ミスを減らすため<br>
    (Range.Valueでそのまま単体の値を取得できるのか or For Each等でループさせるのかの区別が容易)

    ```vbnet
    'sca:参照範囲が1つのセル
    Dim num As Long
    num = Range("foo_sca_bar").Value
    num = num + 100

    'mul:参照範囲が2つ以上のセル
    Dim rng As Range
    For Each rng In Range("foo_mul_bar")
        num = num + rng.Value 
    Next
    ```

    ※各省略形について (sca: scalar) (mul: multi) (head: header) (body: databodyrange) (array: ary)

### (同じ種類の名前の定義セルが2つ以上必要となる場合)
- 最後尾にNumberを連番になるよう付加すること。<br>
特に、セルの削除 または 追加した際には、必ず連番へと修正するように注意すること。

    **[理由]** VBAの処理において、`For` ループ処理を使い `Range("rangeName" & i)` で連番セル全てを取得するという関数が存在するため

    **[example]**<br>
    kihon_sca_inpRange1<br>
    kihon_sca_inpRange2<br>
    kihon_sca_inpRange3

<a name="table-name"></a>

## テーブル名 及び テーブル各見出し名
- ワークシートのセルから参照させたい場合には `"g_"` 、隠蔽したい場合は `"p_"` をprefixする。なお、テーブル各見出し名については省略を可能とする。

    **[理由]** セルから参照できるテーブルの区別が容易 及び 慣習による (g: grobal , p: private)
    - #### [1行のみのテーブル名(名前定義セルのように、直接参照で扱う)]
        **g**_foo_sca_TableFoo
    - #### [2行以上のテーブル名(配列のように、indexmatch関数で一度参照してから扱う)]
        **p**_foo_ary_TableFoo
