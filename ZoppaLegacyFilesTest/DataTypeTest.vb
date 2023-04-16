Imports Xunit
Imports ZoppaLegacyFiles

Namespace ZoppaLegacyFilesTest
    Public Class DataTypeTest

        <Fact>
        Sub ConvertDataTypeTest()
            Dim tp1 = DataTypeConverter.Convert(GetType(NoEnum))
            With tp1
                Dim v1 = .ToValue("ONE")
                Assert.Equal(NoEnum.ONE, v1)
                Dim v2 = .ToValue("Two")
                Assert.Equal(NoEnum.TWO, v2)
                Dim v3 = .ToValue("THREE")
                Assert.Equal(NoEnum.THREE, v3)
                Dim v4 = .ToValue("4")
                Assert.Equal(NoEnum.TWO + NoEnum.TWO, v4)

                Assert.Throws(Of FormatException)(Sub() .ToValue("PPP"))
            End With

            Dim tp2 = DataTypeConverter.Convert(GetType(TimeSpan))
            With tp2
                Dim v1 = .ToValue("12:10:33.123")
                Assert.Equal(New TimeSpan(0, 12, 10, 33, 123), v1)
            End With

            Dim tp3 = DataTypeConverter.Convert(GetType(String))
            With tp3
                Dim v1 = .ToValue("convert test")
                Assert.Equal("convert test", v1)
            End With

            Dim tp4 = DataTypeConverter.Convert(GetType(Single))
            With tp4
                Dim v1 = .ToValue("1.1234")
                Assert.Equal(1.1234F, v1)
            End With

            Dim tp99 = DataTypeConverter.Convert(GetType(Byte()))
            With tp99
                Dim v1 = .ToValue("1F2E3D4C")
                Assert.Equal(New Byte() {&H1F, &H2E, &H3D, &H4C}, v1)
            End With
        End Sub

    End Class

    Enum NoEnum

        ONE = 1

        TWO = 2

        THREE = 3

    End Enum

End Namespace

