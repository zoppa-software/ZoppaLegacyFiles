Option Strict On
Option Explicit On

Imports System.IO

Namespace Csv

    ''' <summary>カンマ区切りで文字列を分割する機能です（EXCEL）</summary>
    Public NotInheritable Class ExcelCsvSpliter
        Inherits Spliter(Of ExcelCsvItem)

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        Private Sub New(inputStream As StreamReader)
            MyBase.New(inputStream)
        End Sub

        ''' <summary>コンストラクタ。</summary>
        ''' <param name="inputText">入力文字列。</param>
        Private Sub New(inputText As String)
            MyBase.New(inputText)
        End Sub

        ''' <summary>カンマ区切り分割機能を生成します。</summary>
        ''' <param name="inputStream">入力ストリーム。</param>
        ''' <returns>カンマ区切り分割機能。</returns>
        Public Shared Function CreateSpliter(inputStream As StreamReader) As ExcelCsvSpliter
            Return New ExcelCsvSpliter(inputStream)
        End Function

        ''' <summary>カンマ区切り分割機能を生成します。</summary>
        ''' <param name="inputText">分割する文字列。</param>
        ''' <returns>カンマ区切り分割機能。</returns>
        Public Shared Function CreateSpliter(inputText As String) As ExcelCsvSpliter
            Return New ExcelCsvSpliter(inputText)
        End Function

        ''' <summary>一行読み込み、分割を行います。</summary>
        ''' <param name="readStream">読み込みストリーム。</param>
        ''' <returns>読み込んだ文字列と分割位置リスト。</returns>
        Protected Overrides Function ReadLine(readStream As TextReader) As (readChars As Char(), splitPoint As List(Of Integer))
            Dim rchars As New List(Of Char)(4096)
            Dim spoint As New List(Of Integer)(256)
            Dim esc As Boolean = False
            Dim index As Integer = 0

            spoint.Add(0)
            Do While readStream.Peek() <> -1
                Dim c = System.Convert.ToChar(readStream.Read())

                Select Case c
                    Case ChrW(13)
                        rchars.Add(c)
                        index += 1
                        If Not esc AndAlso readStream.Peek() = 10 Then
                            readStream.Read()
                            rchars.Add(ChrW(10))
                            index -= 1
                            Exit Do
                        End If

                    Case ChrW(10)
                        rchars.Add(c)
                        index += 1
                        If Not esc Then
                            index -= 1
                            Exit Do
                        End If

                    Case """"c
                        rchars.Add(c)
                        index += 1

                        If esc Then
                            If readStream.Peek() = AscW(""""c) Then
                                rchars.Add(""""c)
                                index += 1
                                readStream.Read()
                            Else
                                esc = False
                            End If
                        ElseIf rchars.Count < 2 OrElse
                               rchars(rchars.Count - 2) = ","c Then
                            esc = True
                        End If

                    Case ","c
                        rchars.Add(c)
                        index += 1
                        If Not esc Then
                            spoint.Add(index)
                        End If

                    Case Else
                        rchars.Add(c)
                        index += 1
                End Select
            Loop
            spoint.Add(index + 1)

            Return (rchars.ToArray(), spoint)
        End Function

    End Class

End Namespace
