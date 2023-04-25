# ZoppaLegacyFiles
データファイルを操作するためのコレクションです。

## 説明
テキストファイルは保存するときに様々な書式ルールによってデータを表現することができます。  
メジャーである `JSON`、`XML` など仕様がはっきりしており、公式のライブラリが存在していますが、古いデータ形式のファイルは仕様が曖昧で、その都度微調整して利用しています。  
このライブラリは私の方言として定義し、業務で使用することを目標としています。  
現在は、以下のファイル形式に対応しています。  
* CSVファイル  
* INIファイル  

## 比較、または特徴
説明でもありますが、仕様が曖昧なデータ形式の一つの実装例として使用していただきたいと思います（多分、全てのニーズには対応できません）    
業務で使用しているとエスケープについても個々のプロジェクトでも個性があるため、エスケープ解除後の文字列、エスケープ解除前の文字列を `UnEscape`プロパティ、`Text`プロパティで取得できるようにしています。  

## 依存関係  
ライブラリは .NET Standard 2.0 で記述しています。そのため、.net framework 4.6.1以降、.net core 2.0以降で使用できます。  
その他のライブラリへの依存関係はありません。  

## 使用方法
### **CSVファイル**
#### 実装仕様
CSVファイルの読み込みは `"` の扱いについて、シンプルに `"` を検出したらエスケープを開始する仕様（`CsvStreamReader`）と EXCELに合わせた仕様（`ExcelCsvStreamReader`）の二種類を用意しています。  
EXCELのCSVファイルの読み込みルールについて明確にはわからなかったため、私の理解した範囲で適用しています。  
#### 使い方
多量のデータの読み込み、書き込みを行うためにCSVファイルはストリームで行います。  
``` vb
' CSVの各列の値をCsvItemへ変換し配列で取得する
Dim values1 As New List(Of CsvItem())()
Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
    For Each v In sr
        values1.Add(v.Items)
    Next
End Using

' CSVの列数とパラメータ数が一致するコンストラクタを使用してインスタンスを生成します
Dim values2 As New List(Of Sample1)()
Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
    For Each v In sr.Select(Of Sample1)(topSkip:=1)
        values2.Add(v)
    Next
End Using

' 上記のコードは LINQ と組み合わせると以下のようになります
Dim values2_1 As List(Of Sample1)()
Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
    values2_1 = sr.Select(Of Sample1)(topSkip:=1).ToList()
End Using

' パラメータの型を引数で指定して、一致するコンストラクタを使用してインスタンスを生成します
Dim values3 As New List(Of Sample1)()
Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
    For Each v In sr.Select(Of Sample1)({DataTypeConverter.StringData, DataTypeConverter.StringData, DataTypeConverter.StringData}, topSkip:=1)
        values3.Add(v)
    Next
End Using
```

以下は`Sample1.csv`の内容です。  
```
項目1,項目2,項目3
"1","2","3"
a,b,c
1,2,3
```
以下は読み込み結果をバインドするクラスの定義です。  
``` vb
Class Sample1
    Public ReadOnly Property Item1 As String
    Public ReadOnly Property Item2 As String
    Public ReadOnly Property Item3 As String

    Public Sub New(s1 As String, s2 As String, s3 As String)
        Me.Item1 = s1
        Me.Item2 = s2
        Me.Item3 = s3
    End Sub
End Class
```
  
コンストラクタのパラメータに使用できる型は以下のようになります。
|パラメータの型|処理内で対応する型|パラメータの型を指定する機能で<br/>指定する値|備考|
|:--|:--|:--|:--|
|Boolean|BooleanDataType|DataTypeConverter.BooleanData||
|Byte()|ByteArrayDataType|DataTypeConverter.ByteArrayData|文字列を16進数と見立てて2文字ずつ変換します。|
|Byte|ByteDataType|DataTypeConverter.ByteData||
|Date|DateTimeDataType|DataTypeConverter.DateTimeData|日付変換はDateTime.TryParseメソッドで行います。<br>DateTimeDataType.DefaultFormatプロパティで書式を設定するとDateTime.TryParseExactで変換を行います。|
|Decimal|DecimalDataType|DataTypeConverter.DecimalData||
|Double|DoubleDataType|DataTypeConverter.DoubleData||
|列挙型|EnumDataType(Of 列挙型)|New EnumDataType(Of 列挙型)で<br/>生成して指定してください||
|Integer|IntegerDataType|DataTypeConverter.IntegerData||
|Long|LongDataType|DataTypeConverter.LongData||
|オブジェクト|ObjectDataType(Of 型)|New ObjectDataType(Of 型)で<br/>生成して指定してください|オブジェクト生成はObjectProvider.SetCreater(Of 型)メソッドで登録したメソッドで行います。登録がない場合は`Nothing`になります。|
|Short|ShortDataType|DataTypeConverter.ShortData||
|Single|SingleDataType|DataTypeConverter.SingleData||
|String|StringDataType|DataTypeConverter.StringData||
|TimeSpan|TimeSpanData|DataTypeConverter.TimeSpan|時刻変換はTimeSpan.TryParseメソッドで行います。<br>TimeSpanDataType.DefaultFormatプロパティで書式を設定するとTimeSpan.TryParseExactで変換を行います。|

ObjectSampleプロジェクトも参照いただければ幸いです。

### **INIファイル**
#### 実装仕様
INIファイルの基本的な仕様には従っていますが、一部の振る舞いが他の実装と異なると思っています。  
* コメントは `;` 以降の文字列となります。セクションの定義に `;` は使えません、値内で `;` を使用する場合は `\` でエスケープするか、`"`、`'` で囲むようにしてください
* 値の前端、後端を `"`、`'` で囲むことで、その中で特殊文字（`=`など）をエスケープせずに使用できます。ただし、 `"`で囲まれた範囲内で`"`を使用する場合は`\`でエスケープするか二つ続けて記述します。または`'`で前後を囲むようにしてください（`'`も同じです）  
#### 使い方
`InitializationFile`クラスを生成して `GetValue`（無名セクションは`GetNoSecssionValue`）で値を読み込み、`SetValue`（無名セクションは`SetNoSecssionValue`）で値を設定します。  
以下は空の`InitializationFile`クラスを生成して、値を設定していく例です。  
``` vb
Dim iniFile As New InitializationFile()
iniFile.SetNoSecssionValue("Number", "100") ' 無名セクション
iniFile.SetValue("LOCAL", "Number", "200")
iniFile.SetValue("OTHER", "Number", "300")
iniFile.SetNoSecssionValue("Name", "A") ' 無名セクション
iniFile.SetValue("LOCAL", "Name", "B")
iniFile.SetValue("OTHER", "Name", "C")
```
作成したINIファイルは `Save` メソッドでストリームに保存します。  
``` vb
Dim buffer As New StringBuilder()
Using sw As New IO.StringWriter(buffer)
    iniFile.Save(sw)
End Using
```
取得した文字列は以下のようになります。  
```
Number=100
Name=A
[LOCAL]
Number=200
Name=B
[OTHER]
Number=300
Name=C
```
  
次に値を取得する例です。以下のの INIファイルを読み込みます。  
```
KEY1 = "     "

[SECTION1]
KEY1=\\keydata1 ; コメント
KEY2=key\=data2

[SECTION2]
KEYA=keydataA

[SPECIAL]
KEYZ=\;\#\=\:\x70CF\r\n改行テスト\0
```
読み込み処理は以下のとおりです。  
``` vb
Dim iniFile = InitializationFile.Load("IniFiles\Sample1.ini", Encoding.GetEncoding("shift_jis"))
Dim a1 = iniFile.GetNoSecssionValue("KEY1")
Assert.True(a1.IsSome)
Assert.Equal("     ", a1.UnEscape)

Dim a2 = iniFile.GetNoSecssionValue("KEY2", "XXX")
Assert.False(a2.IsSome)
Assert.Equal("XXX", a2.UnEscape)

Dim a3 = iniFile.GetValue("SECTION1", "KEY1")
Assert.True(a3.IsSome)
Assert.Equal("\keydata1", a3.UnEscape)

Dim a4 = iniFile.GetValue("SECTION1", "KEY2")
Assert.True(a4.IsSome)
Assert.Equal("key=data2", a4.UnEscape)

Dim a5 = iniFile.GetValue("SECTION2", "KEYA")
Assert.True(a5.IsSome)
Assert.Equal("keydataA", a5.UnEscape)

Dim a6 = iniFile.GetValue("SPECIAL", "KEYZ")
Assert.True(a6.IsSome)
Assert.Equal(";#=:烏" & vbCrLf & "改行テスト" & vbNullChar, a6.UnEscape)
```

## インストール
ソースをビルドして `ZoppaLegacyFiles.dll` ファイルを生成して参照してください。  

## 作成情報
* 造田　崇（zoppa software）
* ミウラ第3システムカンパニー 
* takashi.zouta@kkmiuta.jp

## ライセンス
[apache 2.0](https://www.apache.org/licenses/LICENSE-2.0.html)
