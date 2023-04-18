Imports System
Imports System.Threading

Module Program
    Sub Main(args As String())
        ' 変換ルールが不明なオブジェクトについては ObjectProvider.SetCreater を使用して変換ルールを登録します
        ' 以下の例では「団体コード」列を「PreCode」に変換するための変換ルールを登録しています
        ZoppaLegacyFiles.ObjectProvider.SetCreater(Of PreCode)(
            Function(s) New PreCode(s)
        )

        ' CSVの列数とパラメータ数が一致するコンストラクタを検索し、そのコンストラクタを使用してインスタンスを生成します
        Dim datas As New List(Of Prefecture)()
        Using sr As New ZoppaLegacyFiles.Csv.CsvStreamReader("todouhuken.csv", System.Text.Encoding.UTF8)
            datas.AddRange(sr.Select(Of Prefecture)(1)) ' ヘッダを読み捨てるため topSkipに 1を指定しています
        End Using

        ' 読み込み結果を出力
        For Each dat In datas
            Console.Out.WriteLine($"{dat.no},{dat.name},{dat.area},{dat.code.Value}")
        Next
    End Sub
End Module
