Imports System
Imports System.IO
Imports System.Text
Imports Xunit
Imports ZoppaLegacyFiles.Ini

Public Class IniTest

    Public Sub New()
        Encoding.RegisterProvider(CodePagesEncodingProvider.Instance)
    End Sub

    <Fact>
    Sub ReadTest()
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
    End Sub

    <Fact>
    Sub EscapeTest()
        Dim iniFile = InitializationFile.Load("IniFiles\Sample2.ini", Encoding.GetEncoding("shift_jis"))
        Dim a1 = iniFile.GetNoSecssionValue("KEY1")
        Assert.Equal("c:\user\desktop", a1.UnEscape)

        Dim a2 = iniFile.GetNoSecssionValue("KEY2")
        Assert.Equal("c:\user\desktop", a2.UnEscape)

        Dim a3 = iniFile.GetNoSecssionValue("KEY3")
        Assert.Equal("123""456", a3.UnEscape)

        Dim a4 = iniFile.GetNoSecssionValue("KEY4")
        Assert.Equal("123""456", a4.UnEscape)

        Dim a5 = iniFile.GetNoSecssionValue("KEY5")
        Assert.Equal(";はコメント内で無視されます", a5.UnEscape)
    End Sub

    <Fact>
    Sub SaveTest()
        Dim iniFile = InitializationFile.Load("IniFiles\Sample3.ini", Encoding.GetEncoding("shift_jis"))
        Dim buffer As New StringBuilder()
        Using sw As New IO.StringWriter(buffer)
            iniFile.Save(sw)
        End Using
        Using sr As New StreamReader("IniFiles\Sample3.ini", Encoding.GetEncoding("shift_jis"))
            Assert.Equal(sr.ReadToEnd().Trim(), buffer.ToString().Trim())
        End Using

        buffer.Clear()
        iniFile.SetValue("General", "Editor", " c:\qx\qxw32_changed.exe ")
        Using sw As New IO.StringWriter(buffer)
            iniFile.Save(sw)
        End Using
        Using sr As New StreamReader("IniFiles\Sample3_ptn1.ini", Encoding.GetEncoding("shift_jis"))
            Assert.Equal(sr.ReadToEnd().Trim(), buffer.ToString().Trim())
        End Using

        buffer.Clear()
        iniFile.SetValue("HotKey2", "Key", "Ctrl+Alt+Z")
        iniFile.SetValue("HotKey2", "Editor", "notepad.exe")
        Using sw As New IO.StringWriter(buffer)
            iniFile.Save(sw)
        End Using
        Using sr As New StreamReader("IniFiles\Sample3_ptn2.ini", Encoding.GetEncoding("shift_jis"))
            Assert.Equal(sr.ReadToEnd().Trim(), buffer.ToString().Trim())
        End Using

        buffer.Clear()
        iniFile.RemoveValue("HotKey1", "Key")
        iniFile.RemoveValue("HotKey1", "QxMacro")
        Using sw As New IO.StringWriter(buffer)
            iniFile.Save(sw)
        End Using
        Using sr As New StreamReader("IniFiles\Sample3_ptn3.ini", Encoding.GetEncoding("shift_jis"))
            Assert.Equal(sr.ReadToEnd().Trim(), buffer.ToString().Trim())
        End Using
    End Sub

    <Fact>
    Sub BlockTest()
        Dim iniFile = InitializationFile.Load("IniFiles\Sample4.ini", Encoding.GetEncoding("shift_jis"))
        Dim a1 = iniFile.GetValue("Sample", "Keyに特殊文字=")
        Assert.Equal("値も'特殊文字'", a1.UnEscape)

        Dim a2 = iniFile.GetValue("Sample", "Keyに特殊文字;")
        Assert.Equal("値も""特殊文字""", a2.UnEscape)

        Dim a3 = iniFile.GetValue("Sample", "integer")
        Assert.Equal(123, a3.Convert(Of Integer)())
    End Sub


    <Fact>
    Sub CreateTest()
        Dim iniFile As New InitializationFile()
        iniFile.SetNoSecssionValue("Number", "100") ' 無名セクション
        iniFile.SetValue("LOCAL", "Number", "200")
        iniFile.SetValue("OTHER", "Number", "300")
        iniFile.SetNoSecssionValue("Name", "A") ' 無名セクション
        iniFile.SetValue("LOCAL", "Name", "B")
        iniFile.SetValue("OTHER", "Name", "C")

        Dim buffer As New StringBuilder()
        Using sw As New IO.StringWriter(buffer)
            iniFile.Save(sw)
        End Using
        Assert.Equal("Number=100
Name=A
[LOCAL]
Number=200
Name=B
[OTHER]
Number=300
Name=C", buffer.ToString().Trim())
    End Sub

End Class
