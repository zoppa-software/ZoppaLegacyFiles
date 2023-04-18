Imports System
Imports System.Threading

Module Program
    Sub Main(args As String())
        ' �ϊ����[�����s���ȃI�u�W�F�N�g�ɂ��Ă� ObjectProvider.SetCreater ���g�p���ĕϊ����[����o�^���܂�
        ' �ȉ��̗�ł́u�c�̃R�[�h�v����uPreCode�v�ɕϊ����邽�߂̕ϊ����[����o�^���Ă��܂�
        ZoppaLegacyFiles.ObjectProvider.SetCreater(Of PreCode)(
            Function(s) New PreCode(s)
        )

        ' CSV�̗񐔂ƃp�����[�^������v����R���X�g���N�^���������A���̃R���X�g���N�^���g�p���ăC���X�^���X�𐶐����܂�
        Dim datas As New List(Of Prefecture)()
        Using sr As New ZoppaLegacyFiles.Csv.CsvStreamReader("todouhuken.csv", System.Text.Encoding.UTF8)
            datas.AddRange(sr.Select(Of Prefecture)(1)) ' �w�b�_��ǂݎ̂Ă邽�� topSkip�� 1���w�肵�Ă��܂�
        End Using

        ' �ǂݍ��݌��ʂ��o��
        For Each dat In datas
            Console.Out.WriteLine($"{dat.no},{dat.name},{dat.area},{dat.code.Value}")
        Next
    End Sub
End Module
