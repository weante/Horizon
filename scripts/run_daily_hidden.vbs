' Horizon 每日简报 —— 无窗口启动器
'
' 计划任务以 InteractiveToken 运行 powershell.exe 时会闪出控制台窗口。用户若在
' 任务运行期间把窗口关掉，Windows 会把 CTRL_CLOSE_EVENT 发给整棵进程树，
' Horizon 随之被杀，任务退出码变成 0xC000013A（STATUS_CONTROL_C_EXIT），
' 当天日报就没了。这里用 wscript 以完全隐藏的方式启动 PowerShell，
' 并把 PowerShell 的退出码原样回传给计划任务。
'
' 用法（与 run_daily.ps1 参数一致）：
'   wscript.exe //B run_daily_hidden.vbs
'   wscript.exe //B run_daily_hidden.vbs -Hours 12
'   wscript.exe //B run_daily_hidden.vbs -MailOnly
Option Explicit

Dim shell, cmd, extra, rc, i
Set shell = CreateObject("WScript.Shell")

extra = ""
For i = 0 To WScript.Arguments.Count - 1
    extra = extra & " " & WScript.Arguments(i)
Next

cmd = "powershell.exe -NoProfile -ExecutionPolicy Bypass -File " & _
      """E:\DeepSeek Harness\Horizon\scripts\run_daily.ps1""" & extra

' 0 = 隐藏窗口，True = 等待结束
rc = shell.Run(cmd, 0, True)
WScript.Quit rc
