## 説明
テキストファイルは保存するときに様々な書式ルールによってデータを表現することができます。  
メジャーである `JSON`、`XML` など仕様がはっきりしており、公式のライブラリが存在していますが、古いデータ形式のファイルは仕様が曖昧で、その都度微調整して利用しています。  
このライブラリは私の方言として定義し、業務で使用することを目標としています。  
現在は、以下のファイル形式に対応しています。  
* CSVファイル  
* INIファイル  

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

以上、簡単な説明となります。**ライブラリの詳細は[Githubのページ](https://github.com/zoppa-software/ZoppaLegacyFiles)を参照してください。**

## 更新について
* 0.9.0 ドキュメント作成途中、機能リファクタリング
* 0.0.9 ドキュメント作成途中
