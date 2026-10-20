Public Class PageSetupSystem


    Private Sub PageSetupSystem_Loaded(sender As Object, e As RoutedEventArgs) Handles Me.Loaded

        '重复加载部分
        PanBack.ScrollToHome()

        '非重复加载部分
        Static IsLoaded As Boolean = False
        If IsLoaded Then Return
        IsLoaded = True

        AniControlEnabled += 1
        Reload()
        SliderLoad()
        AniControlEnabled -= 1

    End Sub
    Public Sub Reload()
        SettingService.RefreshSettings(Me)
    End Sub
    Public Sub Reset()
        Try
            SettingService.ResetSettings(Me)
            Logger.Info("已初始化启动器页设置")
            Hint("已初始化启动器页设置！", HintType.Green, False)
        Catch ex As Exception
            Logger.Error(ex, "初始化启动器页设置失败", LogBehavior.Alert)
        End Try
        Reload()
    End Sub

    '滑动条
    Private Sub SliderLoad()
        SliderDownloadThread.GetHintText = Function(v) v + 1
        SliderDownloadSpeed.GetHintText =
        Function(v)
            Select Case v
                Case Is <= 14
                    Return (v + 1) * 0.1 & " M/s"
                Case Is <= 31
                    Return (v - 11) * 0.5 & " M/s"
                Case Is <= 41
                    Return (v - 21) & " M/s"
                Case Else
                    Return "无限制"
            End Select
        End Function
        SliderDebugAnim.GetHintText = Function(v) If(v > 29, "关闭", (v / 10 + 0.1) & "x")
    End Sub
    Private Sub SliderDownloadThread_PreviewChange(sender As Object, e As RouteEventArgs) Handles SliderDownloadThread.PreviewChange
        If SliderDownloadThread.Value < 100 Then Return
        If Not Settings.Get(Of Boolean)("HintDownloadThread") Then
            Settings.Set("HintDownloadThread", True)
            MyMsgBox("如果设置过多的下载线程，可能会导致下载时出现非常严重的卡顿。" & vbCrLf &
                     "一般设置 64 线程即可满足大多数下载需求，除非你知道你在干什么，否则不建议设置更多的线程数！", "警告", "我知道了", IsWarn:=True)
        End If
    End Sub

    '调试模式
    Private Sub CheckDebugMode_Change() Handles CheckDebugMode.Change
        If AniControlEnabled = 0 Then Hint("部分调试信息将在刷新或启动器重启后切换显示！",, False)
    End Sub

    '启动器公告
    Private Sub ComboSystemActivity_SelectionChanged(sender As Object, e As SelectionChangedEventArgs) Handles ComboSystemActivity.SelectionChanged
        If AniControlEnabled <> 0 Then Return
        If ComboSystemActivity.SelectedIndex <> 2 Then Return
        If MyMsgBox("若选择此项，即使在将来出现严重问题时，你也无法获取相关通知。" & vbCrLf &
                    "例如，如果发现某个版本游戏存在严重 Bug，你可能就会因为无法得到通知而导致无法预知的后果。" & vbCrLf & vbCrLf &
                    "一般选择 仅在有重要通知时显示公告 就可以让你尽量不受打扰了。" & vbCrLf &
                    "除非你在制作服务器整合包，或时常手动更新启动器，否则极度不推荐选择此项！", "警告", "我知道我在做什么", "取消", IsWarn:=True) = 2 Then
            If e.RemovedItems.Count > 0 Then ComboSystemActivity.SelectedItem = e.RemovedItems(0)
        End If
    End Sub
    Private Sub BtnSystemUpdate_Click(sender As Object, e As EventArgs) Handles BtnSystemUpdate.Click
        UpdateCheckByButton()
    End Sub
    ''' <summary>
    ''' 启动器是否已经是最新版？
    ''' 若返回 Nothing，则代表无更新缓存文件或出错。
    ''' </summary>
    Public Shared Function IsLauncherNewest() As Boolean?
        Try
            '确认服务器公告是否正常
            Dim ServerContent As String = If(FileUtils.TryReadAsString(PathTemp & "Cache\Notice.cfg"), "")
            If ServerContent.Split("|").Count < 3 Then Return Nothing
            '确认是否为最新
            Return ServerContent.Split("|")(If(BuildType = BuildTypes.Release, 2, 1)) <= VersionCode
        Catch ex As Exception
            Logger.Error(ex, "确认启动器更新失败")
            Return Nothing
        End Try
    End Function

#Region "导出 / 导入设置"

    '只备份可移植的普通偏好。账号、令牌、游戏路径、启动命令和付费主题状态均不应进入明文文件。
    Private Shared ReadOnly PortableSettingKeys As String() = {
        "SystemSystemActivity", "SystemSystemTelemetry",
        "ToolHelpChinese", "ToolDownloadThread", "ToolDownloadSpeed", "ToolDownloadSource",
        "ToolDownloadVersion", "ToolDownloadTranslateV2", "ToolDownloadMod", "ToolModLocalNameStyle",
        "ToolUpdateRelease", "ToolUpdateSnapshot",
        "LaunchArgumentIndieV2", "LaunchArgumentVisible", "LaunchArgumentPriority",
        "LaunchArgumentWindowWidth", "LaunchArgumentWindowHeight", "LaunchArgumentWindowType",
        "LaunchArgumentRam", "LaunchAdvanceGraphicCard", "LaunchAdvanceGC", "LaunchRamType", "LaunchRamCustom",
        "UiLauncherTransparent", "UiLauncherLogo", "UiLauncherEmail", "UiBackgroundColorful",
        "UiBackgroundOpacity", "UiBackgroundBlur", "UiBackgroundSuit", "UiMusicVolume",
        "UiMusicStop", "UiMusicStart", "UiMusicRandom", "UiMusicAuto"
    }

    Private Shared Function IsValidPortableValue(Key As String, Value As JToken) As Boolean
        If Settings.Entries(Key).Type Is GetType(Boolean) Then Return Value.Type = JTokenType.Boolean
        If Value.Type <> JTokenType.Integer Then Return False
        Dim Number = Value.Value(Of Integer)()
        Select Case Key
            Case "SystemSystemActivity", "ToolDownloadSource", "ToolDownloadVersion", "ToolDownloadMod", "LaunchArgumentPriority", "LaunchRamType"
                Return Number >= 0 AndAlso Number <= 2
            Case "ToolDownloadTranslateV2", "LaunchArgumentIndieV2", "LaunchArgumentWindowType", "LaunchAdvanceGC"
                Return Number >= 0 AndAlso Number <= 4
            Case "ToolModLocalNameStyle"
                Return Number >= 0 AndAlso Number <= 1
            Case "ToolDownloadThread"
                Return Number >= 0 AndAlso Number <= 255
            Case "ToolDownloadSpeed"
                Return Number >= 0 AndAlso Number <= 42
            Case "LaunchArgumentVisible"
                Return Number = 0 OrElse (Number >= 2 AndAlso Number <= 5)
            Case "LaunchArgumentWindowWidth", "LaunchArgumentWindowHeight"
                Return Number >= 1 AndAlso Number <= 99999
            Case "LaunchRamCustom"
                Return Number >= 0 AndAlso Number <= 49
            Case "UiLauncherTransparent"
                Return Number >= 0 AndAlso Number <= 600
            Case "UiBackgroundOpacity", "UiMusicVolume"
                Return Number >= 0 AndAlso Number <= 1000
            Case "UiBackgroundBlur"
                Return Number >= 0 AndAlso Number <= 40
            Case "UiBackgroundSuit"
                Return Number >= 0 AndAlso Number <= 12
            Case Else
                Return False
        End Select
    End Function

    Private Sub BtnSystemSettingExp_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemSettingExp.Click
        Dim SavePath = Dialogs.SaveFile("导出 PCL 设置", "PCL-settings.json", filter:={ ("json", "JSON 设置备份") })
        If String.IsNullOrEmpty(SavePath) Then Return
        Try
            Dim Values As New JObject
            For Each Key In PortableSettingKeys
                Values.Add(Key, JToken.FromObject(Settings.GetSafe(Key)))
            Next
            Dim Backup As New JObject From {
                {"format", "PCL-orange-settings"},
                {"version", 1},
                {"settings", Values}
            }
            FileUtils.Write(SavePath, Backup.ToString(Formatting.Indented))
            Hint("设置已导出（不包含账号和密码）！", HintType.Green)
        Catch ex As Exception
            Logger.Error(ex, "导出设置失败", LogBehavior.Alert)
        End Try
    End Sub
    Private Sub BtnSystemSettingImp_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemSettingImp.Click
        Dim ImportPath = Dialogs.SelectFile("导入 PCL 设置", False, filter:="JSON 设置备份|*.json").FirstOrDefault()
        If String.IsNullOrEmpty(ImportPath) Then Return
        Try
            If New FileInfo(ImportPath).Length > 1024 * 1024 Then Throw New InvalidDataException("设置备份文件过大")
            Dim Backup = JObject.Parse(FileUtils.ReadAsString(ImportPath))
            If CStr(Backup("format")) <> "PCL-orange-settings" OrElse CInt(Backup("version")) <> 1 Then
                Throw New InvalidDataException("不是受支持的 PCL 设置备份")
            End If
            Dim Values = TryCast(Backup("settings"), JObject)
            If Values Is Nothing Then Throw New InvalidDataException("设置备份缺少 settings 对象")

            '先检查整份备份，再修改任何设置，以免损坏或不兼容的文件只导入一部分。
            Dim Pending As New Dictionary(Of String, Object)
            For Each PropertyItem As JProperty In Values.Properties()
                If Not PortableSettingKeys.Contains(PropertyItem.Name) Then Continue For
                If Not IsValidPortableValue(PropertyItem.Name, PropertyItem.Value) Then
                    Throw New InvalidDataException("设置值无效：" & PropertyItem.Name)
                End If
                Pending.Add(PropertyItem.Name, PropertyItem.Value.ToObject(Settings.Entries(PropertyItem.Name).Type))
            Next
            If Pending.Count = 0 Then Throw New InvalidDataException("设置备份中没有可导入的偏好")
            If MyMsgBox($"将导入 {Pending.Count} 项偏好，并覆盖对应的当前设置。账号与密码不会导入。", "导入设置", "导入", "取消") <> 1 Then Return

            AniControlEnabled += 1
            Try
                For Each Entry In Pending
                    Settings.SetSafe(Entry.Key, Entry.Value)
                Next
                Reload()
                FrmSetupLaunch?.Refresh()
                FrmSetupUI?.Refresh()
            Finally
                AniControlEnabled -= 1
            End Try
            Hint("设置已导入，部分更改可能需要重启启动器生效！", HintType.Green)
        Catch ex As Exception
            Logger.Error(ex, "导入设置失败", LogBehavior.Alert)
        End Try
    End Sub

    Private Sub BtnSystemOpenLog_Click(sender As Object, e As MouseButtonEventArgs) Handles BtnSystemOpenLog.Click
        OpenExplorer(Paths.Base & "PCL\")
    End Sub

#End Region

End Class
