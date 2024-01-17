    Dim oShp As Shape
    Dim oEffect As Effect
    Dim sldNum As Long
    Dim wavNum As Long
    Dim wavePath As String

    cd = ActivePresentation.Path

    ' すべてのスライドから音声ファイルを削除
    With ActivePresentation
        For i = 1 To .Slides.Count
            Set oSlide = .Slides(i)

            ' スライドからすべての音声ファイルを削除
            For Each oShp In oSlide.Shapes
                If oShp.Type = msoMedia Then
                    oShp.Delete
                End If
            Next oShp
        Next i
    End With

    ' 新しい音声ファイルを埋め込む
    With ActivePresentation
        For i = 1 To .Slides.Count
            sldNum = i ' スライド番号を取得
            wavNum = 1 ' スライドごとの音声番号を初期化
            wavePath = cd & "¥" & sldNum & "-" & wavNum & ".wav" ' ファイルのパス

            ' audioオブジェクトの埋め込み
            Set oSlide = .Slides(i)
            Set oShp = oSlide.Shapes.AddMediaObject2(wavePath, False, True, 10, 10)

            ' Set audio to play automatically
            Set oEffect = oSlide.TimeLine.MainSequence.AddEffect(oShp, msoAnimEffectMediaPlay, , msoAnimTriggerWithPrevious)

            With oShp.AnimationSettings.PlaySettings
                .PlayOnEntry = msoTrue ' 指定したビデオやサウンドは、アニメーションを実行するときに自動再生されます。
                .HideWhileNotPlaying = msoTrue ' 指定したメディア クリップは、スライド ショーの実行時は再生時にのみ表示されます。
                .PauseAnimation = msoTrue ' 指定したメディア クリップの再生が終了するまでスライド ショーは一時停止します。
            End With
        Next i
    End With

    MsgBox "音声ファイルを削除し、新しい音声ファイルを埋め込みました。"
End Sub
