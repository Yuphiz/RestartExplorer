On Error Resume Next

Dim ArrayPathFonders(), oAppShell, WindowOfoAppShell, Shell
Set Dictionary = CreateObject("Scripting.Dictionary")
Set oAppShell = CreateObject("Shell.Application")
Set WindowOfoAppShell=oAppShell.Windows()
Set Shell = CreateObject("WScript.Shell")




'直接关闭快速访问、此电脑、桌面、公共桌面 
Dictionary.Add "::{679F85CB-0220-4080-B29B-5540CC05AAB6}",null
' win11速访问换了id
Dictionary.Add "::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",null
Dictionary.Add "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}",null
Dictionary.Add Shell.SpecialFolders("Desktop"),null
Dictionary.Add Shell.SpecialFolders("AllUsersDesktop"),null
Dictionary.Add "G:\Downloads",null




'声明
'脚本：资源管理器辅助工具_整合版
'版本：v0.4
'说明：本脚本可以关闭重复文件夹、重启资源管理器并打开上次的目录
'链接：https://github.com/Yuphiz/RestartExplorer

'更新：  v0.4  a [增加] windows 11 dev版快速访问(主文件夹)换了新的guid，本次更新增加了新的guid
        '      b [修复] 增加了taskkill，用来修复当用户不存在tskill时重启资源管理器失败的bug
            '  c [增加] 增加了后台检测资源管理器崩溃重启事件，只要触发重启事件，就自动打开上次的文件夹
            '  d [修复] 其他小优化修复
            '  e 改名为 restartexplorer

'更新：  v0.3  a [修复] 修复当ie窗口打开时使用脚本报错

        'v0.2 a 修复带有空格的文件不能重启的错误
        '      b 整合成单文件

'已知问题
'未知

'作者：YUPHIZ
'版权：此脚本版权归YUPHIZ所有
        '凡用此脚本从事法律不允许的事情的，均与本作者无关
        '此脚本遵循 gpl3.0 and later协议



select case WScript.Arguments.count
    case 0
        call RestartExplorer()
    case 1
        call WithArguments()
end select


sub WithArguments()
    select case WScript.Arguments(0)
        case "--UpdateCrashState"
            call Update_CrashState(true)
        case "--SetSchtasksAutoRun"
            processid=Get_SchTask("AutoRun")
            if processid = "" then processid=0
            TASK_CREATION = true
            call Set_AutoRunSchTask(processid,TASK_CREATION)
            Update_CrashState(false)
            call AutoRestartExplorerWhenCrash()
        case "--AutoRestartExplorerWhenCrash"
            Update_CrashState(false)
            call AutoRestartExplorerWhenCrash()
        case "--AutoRestartExplorerWhenCrashInbackgound"
            call AutoRestartExplorerWhenCrashInbackgound() 'Z'
        case "--RemoveTasksch"
            call Remove_Tasksch()

        case "--CloseAll"
            call CloseAllFolders()
        case "--CloseDuplicate" 'Y'
            call CloseDuplicateFolder()
        case else
            call RestartExplorer()
    end select
end sub


sub RestartExplorer()    '重启资源管理器并重新打开上次的目录
On Error Resume Next
n=-1
For Each Oneof in WindowOfoAppShell
    if Instr(1, Oneof.FullName, "\explorer.exe", 1) > 0 Then
        If Not Dictionary.Exists(Oneof.Document.Folder.Self.Path) Then
            n = n + 1
            ReDim Preserve ArrayPathFonders(n)
            ArrayPathFonders(n) = Oneof.Document.Folder.Self.Path
            Dictionary.Add Oneof.Document.Folder.Self.Path ,null
        end if
    end if
Next

    Shell.Run "Tskill explorer",0,True
    if err.number = -2147024894 then 
        err.number = 0
        Shell.Run "taskkill /im explorer.exe /f",0,True
        ' Shell.Run "explorer",0,True
        Shell.Run "cmd /c start explorer",0,True
        if err.number = -2147024894 then
            msgbox "重启资源管理器失败，请手动重启"
            exit sub
        end if
    end if

For Each Oneof in ArrayPathFonders
    Shell.Run """"&Oneof&""""
Next
end sub



sub CloseDuplicateFolder()   '关闭重复文件夹
For i=WindowOfoAppShell.count-1 to 0 step -1
    if Instr(1, WindowOfoAppShell(i).FullName, "\explorer.exe", 1) > 0 Then
            ' msgbox WindowOfoAppShell(i).Document.Folder.Self.Path
        If Dictionary.Exists(WindowOfoAppShell(i).Document.Folder.Self.Path) Then
            WindowOfoAppShell(i).quit 'H'
        else
            Dictionary.Add WindowOfoAppShell(i).Document.Folder.Self.Path ,null
        end if
    end if
next
end sub



sub CloseAllFolders()    '关闭所有文件夹
On Error Resume Next
Set FolderWindows = CreateObject("Shell.Application").Windows
For i=FolderWindows.count-1 to 0 step -1
    if Instr(1, Oneof.FullName, "\explorer.exe", 1) > 0 Then  'I'
        FolderWindows(i).Quit
    end if
Next
end sub


function Update_CrashState(CrashState)
    UserName=CreateObject("WScript.Network").UserName
    set ShellTask=CreateObject("Schedule.Service")
    ShellTask.connect()
    set rootFolder=ShellTask.getfolder("\")
    
    set taskDefinition=ShellTask.NewTask(0)
    
    set Settings = taskDefinition.Settings
    Settings.StartWhenAvailable = true
    Settings.DisallowStartIfOnBatteries = false
    Settings.ExecutionTimeLimit= "PT5M"

    TASK_TRIGGER_EVENT = 0
    set triggers = taskDefinition.Triggers
    set trigger = triggers.Create(TASK_TRIGGER_EVENT)
    trigger.Subscription = "<QueryList><Query Id='0' Path='Application'><Select Path='Application'>*[System[Provider[@Name='Microsoft-Windows-Winlogon'] and (Level=4 or Level=0) and (EventID=1002)]]</Select></Query></QueryList>"
    trigger.Enabled = true
    
    taskDefinition.Data = CrashState
            
    set Action = taskDefinition.Actions.Create(0)
    Action.Id = "RestartExplorerWhenCrash"
    Action.Path = "wscript"
    Action.Arguments= """"&Wscript.ScriptFullName&""" --UpdateCrashState"
    
    CreateOrUpdateTask = 6
    rootFolder.RegisterTaskDefinition "\YuphizScript\"&UserName&"\RestartExplorer\RestartExplorerWhenCrash",taskDefinition,CreateOrUpdateTask,null,null,3
end function

function Set_AutoRunSchTask(Data,isEnable)
    UserName=CreateObject("WScript.Network").UserName
    set ShellTask=CreateObject("Schedule.Service")
    ShellTask.connect()
    set rootFolder=ShellTask.getfolder("\")
    
    set taskDefinition=ShellTask.NewTask(0)
    
    set Settings = taskDefinition.Settings
    Settings.StartWhenAvailable = true
    Settings.DisallowStartIfOnBatteries = false
    Settings.ExecutionTimeLimit= "PT5M"

    TASK_TRIGGER_LOGON = 9
    set triggers = taskDefinition.Triggers
    set trigger =  triggers.Create(TASK_TRIGGER_LOGON)
    trigger.UserId = UserName
    trigger.delay = "PT60S"
    trigger.Enabled = true
    
    taskDefinition.Data = Data
            
    set Action = taskDefinition.Actions.Create(0)
    Action.Id = "RestartExplorerWhenCrash"
    Action.Path = "wscript"
    Action.Arguments= """"&Wscript.ScriptFullName&""" --AutoRestartExplorerWhenCrash"
    
    TASK_CREATE_OR_UPDATE = 6
    TASK_DISABLE = 8
    TASK_CREATION = TASK_CREATE_OR_UPDATE
    
    rootFolder.RegisterTaskDefinition "\YuphizScript\"&UserName&"\RestartExplorer\AutoRun",taskDefinition,TASK_CREATION,null,null,3
end function

sub AutoRestartExplorerWhenCrash()
    set WR=getobject("winmgmts:\\.\root\cimv2")
    processid=Get_SchTask("AutoRun")
    if processid <>""  then
        if processid <> 0 then set ps=WR.execquery("select * from win32_process where processid = "&processid)
        for each Oneof in ps
            if Oneof.name = "wscript.exe" then 
                msgbox "RestartExplorer已存在后台，不需要重复启动"
                wscript.quit
            end if
        next
    end if

    Set Shell = CreateObject("WScript.Shell")
    Set oExec = Shell.exec("wscript "&Wscript.ScriptFullName &" --AutoRestartExplorerWhenCrashInbackgound")
    NewProcessId = oExec.processid
    TASK_CREATION = false
    Set_AutoRunSchTask NewProcessId,TASK_CREATION
end sub


function Get_SchTask(TaskName)
    UserName=CreateObject("WScript.Network").UserName
    set ShellTask=CreateObject("Schedule.Service")
    ShellTask.connect()
    set rootFolder=ShellTask.getfolder("\")
    On Error Resume Next
    set SchTasksObj = rootFolder.gettask("\YuphizScript\"&UserName&"\RestartExplorer\"&TaskName)
    Data = SchTasksObj.Definition.Data
    On Error Goto 0
    if isnull(Data) then Data = null
    Get_SchTask = Data
end function

sub Stop_Background(processid)
set WR=getobject("winmgmts:\\.\root\cimv2") 
set ps=WR.execquery("select * from win32_process where processid = "&processid)
for each Oneof in ps
    if Oneof.name = "wscript.exe" then 
        CreateObject("WScript.Shell").run "taskkill /im " &processid& " /f",0  'P'
        exit for
    end if
next
end sub


sub AutoRestartExplorerWhenCrashInbackgound()
    Set Shell = CreateObject("WScript.Shell")
    count=-1
    Dim ArrayPathFonders(), oAppShell, WindowOfoAppShell, Shell

    do until 1>2
        Set Dictionary = CreateObject("Scripting.Dictionary") 'U'
        Set oAppShell = CreateObject("Shell.Application")
        Set WindowOfoAppShell=oAppShell.Windows()
        

        '直接关闭快速访问、此电脑、桌面、公共桌面
        Dictionary.Add "::{679F85CB-0220-4080-B29B-5540CC05AAB6}",null
        ' win11速访问换了id
        Dictionary.Add "::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",null
        Dictionary.Add "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}",null
        Dictionary.Add Shell.SpecialFolders("Desktop"),null
        Dictionary.Add Shell.SpecialFolders("AllUsersDesktop"),null
        Dictionary.Add "G:\Downloads",null

        n=-1
        On Error Resume Next
        For Each Oneof in WindowOfoAppShell
            if Instr(1, Oneof.FullName, "\explorer.exe", 1) > 0 Then
                If Not Dictionary.Exists(Oneof.Document.Folder.Self.Path) Then
                    n = n + 1
                    ReDim Preserve ArrayPathFonders(n)
                    ArrayPathFonders(n) = Oneof.Document.Folder.Self.Path
                    Dictionary.Add Oneof.Document.Folder.Self.Path,null
                end if
            end if
        Next
        On Error goto 0

        CrashState =  Get_SchTask("RestartExplorerWhenCrash")
        if CrashState then
            ' wscript.sleep 1000
            On Error Resume Next
            count = Ubound(ArrayPathFonders)
            On Error goto 0
            if count>=0 then
                ' Shell.popup "资源管理器崩溃了，正在打开崩溃前的目录",1
                if count >= 9 then count = 9
                For i=0 to count
                    Shell.Run """"&ArrayPathFonders(i)&""""
                    ' msgbox ArrayPathFonders(i)
                Next
            end if
            
            Erase ArrayPathFonders  '清空数组
            Dictionary=RemoveAll  '清空字典
            Update_CrashState(false)
            count = -1
        end if
        wscript.sleep 3000
    loop
end sub


' 移除任务计划
sub Remove_Tasksch()
    processid=Get_SchTask("AutoRun")
    if processid <>"" then
        if processid <> 0 then Stop_Background(processid)
    end if
    title = "RestartExplorer"
    UserName=CreateObject("WScript.Network").UserName
    set Shell=CreateObject("Wscript.Shell") 

    askdelete=Shell.popup( _
      "防误操作，真的要删除【"&title&"】吗？", _
      0, _
      "防误操作，请再确认",_
      1+48+256+4096 _
    )
    if askdelete=2 then wscript.quit

    set ShellTask=createobject("Schedule.Service")
    call ShellTask.connect()
    On Error Resume Next
    set rootFolder=ShellTask.getfolder("\YuphizScript\"&UserName)
        call rootFolder.DeleteTask(title&"\RestartExplorerWhenCrash",0)
        call rootFolder.DeleteTask(title&"\autorun",0)
    On Error goto 0
    rootFolder.deleteFolder title,0
    wscript.sleep 700
    Shell.popup _
      "成功卸载【"&title&"】"&vbcrlf&_
      "已删除全部的任务计划",_
    3
end sub
