On Error Resume Next

Dim ArrayPathFonders(), oAppShell, WindowOfoAppShell, Shell
Set Dictionary = CreateObject("Scripting.Dictionary")
Set oAppShell = CreateObject("Shell.Application")
Set WindowOfoAppShell=oAppShell.Windows()
Set Shell = CreateObject("WScript.Shell")




'ֱ�ӹرտ��ٷ��ʡ��˵��ԡ����桢�������� 
Dictionary.Add "::{679F85CB-0220-4080-B29B-5540CC05AAB6}",null
' win11�ٷ��ʻ���id
Dictionary.Add "::{F874310E-B6B7-47DC-BC84-B9E6B38F5903}",null
Dictionary.Add "::{20D04FE0-3AEA-1069-A2D8-08002B30309D}",null
Dictionary.Add Shell.SpecialFolders("Desktop"),null
Dictionary.Add Shell.SpecialFolders("AllUsersDesktop"),null
Dictionary.Add "G:\Downloads",null




'����
'�ű�����Դ��������������_���ϰ�
'�汾��v0.4
'˵�������ű����Թر��ظ��ļ��С�������Դ�����������ϴε�Ŀ¼
'���ӣ�https://github.com/Yuphiz/RestartExplorer

'���£�  v0.4  a [����] windows 11 dev����ٷ���(���ļ���)�����µ�guid�����θ����������µ�guid
        '      b [�޸�] ������taskkill�������޸����û�������tskillʱ������Դ������ʧ�ܵ�bug
            '  c [����] �����˺�̨�����Դ���������������¼���ֻҪ���������¼������Զ����ϴε��ļ���
            '  d [�޸�] ����С�Ż��޸�
            '  e ����Ϊ restartexplorer

'���£�  v0.3  a [�޸�] �޸���ie���ڴ�ʱʹ�ýű�����

        'v0.2 a �޸����пո���ļ����������Ĵ���
        '      b ���ϳɵ��ļ�

'��֪����
'δ֪

'���ߣ�YUPHIZ
'��Ȩ���˽ű���Ȩ��YUPHIZ����
        '���ô˽ű����·��ɲ����������ģ����뱾�����޹�
        '�˽ű���ѭ gpl3.0 and laterЭ��



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


sub RestartExplorer()    '������Դ�����������´��ϴε�Ŀ¼
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
            msgbox "������Դ������ʧ�ܣ����ֶ�����"
            exit sub
        end if
    end if

For Each Oneof in ArrayPathFonders
    Shell.Run """"&Oneof&""""
Next
end sub



sub CloseDuplicateFolder()   '�ر��ظ��ļ���
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



sub CloseAllFolders()    '�ر������ļ���
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
                msgbox "RestartExplorer�Ѵ��ں�̨������Ҫ�ظ�����"
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
        

        'ֱ�ӹرտ��ٷ��ʡ��˵��ԡ����桢��������
        Dictionary.Add "::{679F85CB-0220-4080-B29B-5540CC05AAB6}",null
        ' win11�ٷ��ʻ���id
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
                ' Shell.popup "��Դ�����������ˣ����ڴ򿪱���ǰ��Ŀ¼",1
                if count >= 9 then count = 9
                For i=0 to count
                    Shell.Run """"&ArrayPathFonders(i)&""""
                    ' msgbox ArrayPathFonders(i)
                Next
            end if
            
            Erase ArrayPathFonders  '�������
            Dictionary=RemoveAll  '����ֵ�
            Update_CrashState(false)
            count = -1
        end if
        wscript.sleep 3000
    loop
end sub


' �Ƴ�����ƻ�
sub Remove_Tasksch()
    processid=Get_SchTask("AutoRun")
    if processid <>"" then
        if processid <> 0 then Stop_Background(processid)
    end if
    title = "RestartExplorer"
    UserName=CreateObject("WScript.Network").UserName
    set Shell=CreateObject("Wscript.Shell") 

    askdelete=Shell.popup( _
      "������������Ҫɾ����"&title&"����", _
      0, _
      "�������������ȷ��",_
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
      "�ɹ�ж�ء�"&title&"��"&vbcrlf&_
      "��ɾ��ȫ��������ƻ�",_
    3
end sub
