Option Strict On
Option Explicit On

Imports System.IO
Imports System.Text

Namespace Ini

    ''' <summary>INIファイル操作機能です。</summary>
    Public NotInheritable Class InitializationFile

        ''' <summary>セクション別にキーと値を保持するディクショナリ。</summary>
        Private ReadOnly mKeyAndValue As New Dictionary(Of Section, Dictionary(Of String, KeyAndValue))()

        ''' <summary>INIファイルの各行の情報です。</summary>
        Private ReadOnly mIniLines As New List(Of IIniLine)()

        ''' <summary>コンストラクタ。</summary>
        Public Sub New()

        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="lines">INIファイルの各行。</param>
        Private Sub New(lines As List(Of String))
            Dim currentSection As New Section()
            For row As Integer = 0 To lines.Count - 1
                ' 不要な空白を取り除いた後、セクションかキー／値の何れかか判定する
                '
                ' 1. コメント行ならば読み飛ばし
                ' 2. セクションの場合、カレントセクションを更新
                ' 3. キー／値の場合、カレントセクションのリストに追加
                Dim s = If(lines(row), "").Trim()
                Dim kv = New String() {"", "", "", ""}
                Dim sp As Integer
                If s <> "" Then
                    If s(0) = ";"c Then                                                                 ' 1
                        Me.mIniLines.Add(New EmptyLine(row, lines(row)))
                    ElseIf IsSection(s) Then
                        currentSection = New Section(s.Substring(1, s.Length - 2), row, lines(row))     ' 2
                        Me.mIniLines.Add(currentSection)
                    ElseIf IsKeyPair(s, kv, sp) Then
                        If Not Me.mKeyAndValue.ContainsKey(currentSection) Then
                            Me.mKeyAndValue.Add(currentSection, New Dictionary(Of String, KeyAndValue)())
                        End If
                        If Not Me.mKeyAndValue(currentSection).ContainsKey(kv(0)) Then
                            Dim keyAndVal As New KeyAndValue(kv, sp, row, lines(row))
                            Me.mKeyAndValue(currentSection).Add(kv(0), keyAndVal)                       ' 3
                            Me.mIniLines.Add(keyAndVal)
                        Else
                            Throw New ArgumentException($"既に同じキーが登録されています。[{currentSection.Name}]{kv(0)}")
                        End If
                    Else
                        Throw New ArgumentException("セクション、キー／値以外の値が与えられました")
                    End If
                Else
                    Me.mIniLines.Add(New EmptyLine(row, lines(row)))
                End If
            Next
        End Sub

        ''' <summary>INIファイルを読み込みます。</summary>
        ''' <param name="iniString">INIファイルのパス。</param>
        ''' <returns>INIファイル情報。</returns>
        Public Shared Function LoadIni(iniString As String) As InitializationFile
            Dim lines As New List(Of String)()
            Using sr As New IO.StringReader(iniString)
                Do While sr.Peek() <> -1
                    Dim ln = sr.ReadLine()
                    If ln.Trim() <> "" Then
                        lines.Add(ln)
                    End If
                Loop
            End Using

            Return New InitializationFile(lines)
        End Function

        ''' <summary>指定したエンコードで INIファイルを読み込みます。</summary>
        ''' <param name="iniFilePath">INIファイルのパス。</param>
        ''' <param name="encode">エンコード（指定がない場合はシステムのデフォルトエンコード）</param>
        ''' <returns>INIファイルクラス。</returns>
        Public Shared Function Load(iniFilePath As String, Optional encode As Text.Encoding = Nothing) As InitializationFile
            ' エンコードの指定が無ければデフォルトエンコードを設定
            Dim enc = encode
            If enc Is Nothing Then
                enc = Text.Encoding.Default
            End If

            ' 全ての行を読み込み、インスタンスの生成へ
            Dim fi As New IO.FileInfo(iniFilePath)
            If fi.Exists Then
                Dim lines As New List(Of String)()
                Using sr As New IO.StreamReader(iniFilePath, enc)
                    Do While sr.Peek() <> -1
                        Dim ln = sr.ReadLine()
                        lines.Add(ln)
                    Loop
                End Using

                Return New InitializationFile(lines)
            Else
                Throw New IO.FileNotFoundException("指定したファイルが存在しません")
            End If
        End Function

        ''' <summary>文字列がセクションかどうかを判定します。</summary>
        ''' <param name="ln">読み込んだ行文字列。</param>
        ''' <returns>セクションならば真。</returns>
        Private Shared Function IsSection(ln As String) As Boolean
            If ln.Length = 0 OrElse ln(0) <> "["c Then
                Return False
            End If

            Dim esc = False
            Dim sln = ln
            For i As Integer = 0 To ln.Length - 1
                If ln(i) = "\"c AndAlso i < ln.Length - 1 AndAlso ln(i + 1) = "\"c Then
                    i += 1
                ElseIf ln(i) = "\"c Then
                    esc = True
                ElseIf Not esc AndAlso ln(i) = ";"c Then
                    sln = ln.Substring(0, i)
                    Exit For
                ElseIf esc Then
                    esc = False
                End If
            Next

            sln = sln.Trim()
            Return (sln.Length > 2 AndAlso sln(0) = "["c AndAlso sln(sln.Length - 1) = "]"c)
        End Function

        ''' <summary>値の文字列がユニコード文字指定ならば真を返します。</summary>
        ''' <param name="str">値の文字列。</param>
        ''' <returns>ユニコード文字指定ならば真。</returns>
        Private Shared Function IsUnicode(str As String) As Boolean
            If str(0) = "x"c OrElse str(0) = "X"c Then
                For i As Integer = 1 To 4
                    If (str(i) >= "0"c AndAlso str(i) <= "9"c) OrElse
                   (str(i) >= "a"c AndAlso str(i) <= "f"c) OrElse
                   (str(i) >= "A"c AndAlso str(i) <= "F"c) Then
                    Else
                        Return False
                    End If
                Next
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>文字列がキー／値かどうかを判定します。</summary>
        ''' <param name="ln">読み込んだ行文字列。</param>
        ''' <param name="outKeyAndValue">キーと値の配列。</param>
        ''' <param name="splitPos">キーと値の分割位置。</param>
        ''' <returns>キー／値ならば真。</returns>
        Private Shared Function IsKeyPair(ln As String, outKeyAndValue As String(), ByRef splitPos As Integer) As Boolean
            Dim strs = New Text.StringBuilder() {
            New Text.StringBuilder(),
            New Text.StringBuilder()
        }
            Dim index As Integer = 0
            Dim esc As Boolean = False
            Dim endc As Char
            For i As Integer = 0 To ln.Length - 1
                If Not esc Then
                    If ln(i) = "\"c AndAlso i < ln.Length - 5 AndAlso IsUnicode(ln.Substring(i + 1, 5)) Then
                        strs(index).Append(ln.Substring(i, 6))
                        i += 5
                    ElseIf ln(i) = "\"c AndAlso i < ln.Length - 1 Then
                        strs(index).Append(ln.Substring(i, 2))
                        i += 1
                    ElseIf ln(i) = "="c Then
                        index += 1
                        If index > 1 Then
                            Return False
                        End If
                        splitPos = i + 1
                    ElseIf ln(i) = """"c OrElse ln(i) = "'"c Then
                        If strs(index).ToString().Trim() = "" Then
                            esc = True
                            endc = ln(i)
                        End If
                        strs(index).Append(ln(i))
                    ElseIf ln(i) = ";"c Then
                        Exit For
                    Else
                        strs(index).Append(ln(i))
                    End If
                Else
                    If (ln(i) = "\"c OrElse ln(i) = endc) AndAlso (i < ln.Length - 1 AndAlso ln(i + 1) = endc) Then
                        strs(index).Append(ln(i + 1))
                        i += 1
                    ElseIf ln(i) = endc Then
                        esc = False
                    End If
                    strs(index).Append(ln(i))
                End If
            Next

            If index = 1 Then
                For p As Integer = 0 To 1
                    Dim str = strs(p).ToString().Trim()
                    Dim unstr As New StringBuilder()

                    esc = (str(0) = """"c OrElse str(0) = "'"c)
                    For i As Integer = If(esc, 1, 0) To str.Length - 1
                        If Not esc Then
                            If str(i) = "\"c AndAlso i < str.Length - 1 Then
                                Dim c As Char? = Nothing
                                Select Case str(i + 1)
                                    Case "0"c
                                        c = CChar(vbNullChar)
                                    Case "t"c, "T"c
                                        c = CChar(vbTab)
                                    Case "r"c, "R"c
                                        c = CChar(vbCr)
                                    Case "n"c, "N"c
                                        c = CChar(vbLf)
                                    Case ";"c, "#"c, "="c, ":"c, "\"c
                                        c = str(i + 1)
                                    Case "x"c, "X"c
                                        If i < str.Length - 5 Then
                                            Try
                                                Dim num As Integer = 0
                                                For j As Integer = i + 2 To i + 5
                                                    num = (num << 4) + System.Convert.ToInt32($"{str(j)}", 16)
                                                Next
                                                c = ChrW(num)
                                                i += 4
                                            Catch ex As Exception
                                                ' 空実装
                                            End Try
                                        End If
                                End Select

                                If c.HasValue Then
                                    unstr.Append(c)
                                    i += 1
                                End If
                            Else
                                unstr.Append(str(i))
                            End If
                        Else
                            If (str(i) = "\"c OrElse str(i) = str(0)) AndAlso (i < str.Length - 1 AndAlso str(i + 1) = str(0)) Then
                                unstr.Append(str(i + 1))
                                i += 1
                            ElseIf str(i) = str(0) Then
                                esc = False
                            Else
                                unstr.Append(str(i))
                            End If
                        End If
                    Next

                    outKeyAndValue(p * 2) = unstr.ToString()
                    outKeyAndValue(p * 2 + 1) = str
                Next
                Return True
            Else
                Return False
            End If
        End Function

        ''' <summary>キーを指定して値を取得します（無名セクション）</summary>
        ''' <param name="key">キー。</param>
        ''' <param name="defaultValue">デフォルト値。</param>
        ''' <returns>取得した値。</returns>
        Public Function GetNoSecssionValue(key As String, Optional defaultValue As String = "") As ValueResult
            Return GetValue(Nothing, key, defaultValue)
        End Function

        ''' <summary>セクションとキーを指定して、値を取得します。</summary>
        ''' <param name="section">セクション名。</param>
        ''' <param name="key">キー。</param>
        ''' <param name="defaultValue">デフォルト値。</param>
        ''' <returns>取得した値。</returns>
        Public Function GetValue(secssion As String, key As String, Optional defaultValue As String = "") As ValueResult
            Dim sec = If(secssion Is Nothing, New Section(), New Section(secssion))
            Dim val As KeyAndValue = Nothing
            If Me.mKeyAndValue.ContainsKey(sec) AndAlso Me.mKeyAndValue(sec).TryGetValue(key, val) Then
                Return New ValueResult(True, val.ValueUnEscape, val.Value)
            Else
                Return New ValueResult(False, defaultValue, defaultValue)
            End If
        End Function

        ''' <summary>キーを指定して値を設定します（無名セクション）</summary>
        ''' <param name="key">キー。</param>
        ''' <param name="newValue">値。</param>
        Public Sub SetNoSecssionValue(key As String, newValue As String)
            Me.SetValue(Nothing, key, newValue)
        End Sub

        ''' <summary>セクションとキーを指定して、値を設定します。</summary>
        ''' <param name="section">セクション名。</param>
        ''' <param name="key">キー。</param>
        ''' <param name="newValue">値。</param>
        Public Sub SetValue(secssion As String, key As String, newValue As String)
            Dim sec = If(secssion Is Nothing, New Section(), New Section(secssion))
            Dim val As KeyAndValue = Nothing
            If Me.mKeyAndValue.ContainsKey(sec) AndAlso Me.mKeyAndValue(sec).TryGetValue(key, val) Then
                ' 既存の項目をアップデート
                val.UpdateValue(newValue)
            Else
                ' 新規追加
                '
                ' 1. 新規セクション追加
                ' 2. キー／値を追加
                Dim inspos As Double = -1
                If Not Me.mKeyAndValue.ContainsKey(sec) Then
                    Me.mKeyAndValue.Add(sec, New Dictionary(Of String, KeyAndValue)())
                    sec.LineNo = Me.mIniLines.Count
                    Me.mIniLines.Add(sec)
                    inspos = Me.mIniLines.Count
                Else
                    If Me.mKeyAndValue(sec).Count > 0 Then
                        inspos = Me.mKeyAndValue(sec).Last().Value.LineNo + 0.5
                    Else
                        inspos = sec.LineNo + 0.5
                    End If
                End If

                Dim addValue = New KeyAndValue(key, newValue, inspos)
                Me.mKeyAndValue(sec).Add(key, addValue)
                Me.mIniLines.Add(addValue)

                Me.AjustLines()
            End If
        End Sub

        ''' <summary>キーを指定して値を削除します（無名セクション）</summary>
        ''' <param name="key">キー。</param>
        Public Sub RemoveValue(key As String)
            Me.RemoveValue(Nothing, key)
        End Sub

        ''' <summary>セクションとキーを指定して、値を削除します。</
        ''' <param name="section">セクション名。</param>
        ''' <param name="key">キー。</param>
        Public Sub RemoveValue(secssion As String, key As String)
            Dim sec = If(secssion Is Nothing, New Section(), New Section(secssion))
            Dim val As KeyAndValue = Nothing
            If Me.mKeyAndValue.ContainsKey(sec) AndAlso Me.mKeyAndValue(sec).TryGetValue(key, val) Then
                Me.mKeyAndValue(sec).Remove(key)
                Me.mIniLines.Remove(val)

                If Me.mKeyAndValue(sec).Count = 0 AndAlso Not sec.DefaultSection Then
                    Me.mKeyAndValue.Remove(sec)
                    Me.mIniLines.Remove(sec)
                End If

                Me.AjustLines()
            End If
        End Sub

        ''' <summary>INIファイルを保存します。</summary>
        ''' <param name="path">ファイルパス。</param>
        ''' <param name="encode">エンコード。</param>
        Public Sub Save(path As String, Optional encode As Text.Encoding = Nothing)
            ' エンコードの指定が無ければデフォルトエンコードを設定
            Dim enc = encode
            If enc Is Nothing Then
                enc = Text.Encoding.Default
            End If

            Using sw As New IO.StreamWriter(path, False, enc)
                Me.Save(sw)
            End Using
        End Sub

        ''' <summary>INIファイルを保存します。</summary>
        ''' <param name="sw">書き込むストリーム。</param>
        Public Sub Save(sw As TextWriter)
            Me.AjustLines()

            For Each ln In Me.mIniLines
                sw.WriteLine(ln.WriteStr)
            Next
        End Sub

        ''' <summary>行番号を再割り当てします。</summary>
        Private Sub AjustLines()
            Me.mIniLines.Sort(Function(l, r) l.LineNo.CompareTo(r.LineNo))
            For i As Integer = 0 To Me.mIniLines.Count - 1
                Me.mIniLines(i).LineNo = i
            Next
        End Sub

    End Class

End Namespace
