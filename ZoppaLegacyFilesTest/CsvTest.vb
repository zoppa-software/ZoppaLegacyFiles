Imports System.Text
Imports Xunit
Imports ZoppaLegacyFiles
Imports ZoppaLegacyFiles.Csv

Public Class CsvTest

    Public Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

    <Fact>
    Sub SplitTest()
        Dim csv1 = CsvSpliter.CreateSpliter("あ,い,う,え,""お,を""").Split()
        Assert.Equal(csv1(0).UnEscape(), "あ")
        Assert.Equal(csv1(1).UnEscape(), "い")
        Assert.Equal(csv1(2).UnEscape(), "う")
        Assert.Equal(csv1(3).UnEscape(), "え")
        Assert.Equal(csv1(4).UnEscape(), "お,を")
        Assert.Equal(csv1(4).Text, """お,を""")

        Dim csv2 = CsvSpliter.CreateSpliter("あ,い,う,え, """"お,を""""").Split()
        Assert.Equal(csv2(0).UnEscape(), "あ")
        Assert.Equal(csv2(1).UnEscape(), "い")
        Assert.Equal(csv2(2).UnEscape(), "う")
        Assert.Equal(csv2(3).UnEscape(), "え")
        Assert.Equal(csv2(4).UnEscape(), "お")
        Assert.Equal(csv2(5).UnEscape(), "を")

        Dim csv3 = CsvSpliter.CreateSpliter("あ,い,う,え,""""お,を""""").Split()
        Assert.Equal(csv3(0).UnEscape(), "あ")
        Assert.Equal(csv3(1).UnEscape(), "い")
        Assert.Equal(csv3(2).UnEscape(), "う")
        Assert.Equal(csv3(3).UnEscape(), "え")
        Assert.Equal(csv3(4).UnEscape(), "お")
        Assert.Equal(csv3(5).UnEscape(), "を")

        Dim csv4 = CsvSpliter.CreateSpliter("あ,い,う,え, ""お,を""").Split()
        Assert.Equal(csv4(0).UnEscape(), "あ")
        Assert.Equal(csv4(1).UnEscape(), "い")
        Assert.Equal(csv4(2).UnEscape(), "う")
        Assert.Equal(csv4(3).UnEscape(), "え")
        Assert.Equal(csv4(4).UnEscape(), "お,を")

        Dim csv5 = CsvSpliter.CreateSpliter("あ,い,う,え, ""お"" ""お,を""").Split()
        Assert.Equal(csv5(0).UnEscape(), "あ")
        Assert.Equal(csv5(1).UnEscape(), "い")
        Assert.Equal(csv5(2).UnEscape(), "う")
        Assert.Equal(csv5(3).UnEscape(), "え")
        Assert.Equal(csv5(4).UnEscape(), "お お,を")

        Dim csv6 = CsvSpliter.CreateSpliter("あ,い,う,え,""お"" ""お,を""").Split()
        Assert.Equal(csv6(0).UnEscape(), "あ")
        Assert.Equal(csv6(1).UnEscape(), "い")
        Assert.Equal(csv6(2).UnEscape(), "う")
        Assert.Equal(csv6(3).UnEscape(), "え")
        Assert.Equal(csv6(4).UnEscape(), "お お,を")
    End Sub

    <Fact>
    Sub FileReadTest()
        Dim values1 As New List(Of CsvItem())()
        Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
            For Each v In sr
                values1.Add(v.Items)
            Next
        End Using
        Assert.Equal(5, values1.Count)

        Dim values2 As New List(Of Sample1)()
        Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
            For Each v In sr.Select(Of Sample1)(topSkip:=1)
                values2.Add(v)
            Next
        End Using
        Assert.Equal(4, values2.Count)

        Dim values3 As New List(Of Sample1)()
        Using sr As New CsvStreamReader("CsvFiles\Sample1.csv", Encoding.GetEncoding("shift_jis"))
            For Each v In sr.Select(Of Sample1)({DataTypeConverter.StringData, DataTypeConverter.StringData, DataTypeConverter.StringData}, topSkip:=1)
                values3.Add(v)
            Next
        End Using
        Assert.Equal(4, values3.Count)
    End Sub

End Class

Class Sample1

    Public ReadOnly Property Item1 As String

    Public ReadOnly Property Item2 As String

    Public ReadOnly Property Item3 As String

    Public Sub New(s1 As String, s2 As String, s3 As String)
        Me.Item1 = s1
        Me.Item2 = s2
        Me.Item3 = s3
    End Sub

    Public Overrides Function Equals(obj As Object) As Boolean
        Dim other = TryCast(obj, Sample1)
        Return Me.Item1 = other.Item1 AndAlso
                   Me.Item2 = other.Item2 AndAlso
                   Me.Item3 = other.Item3
    End Function

End Class