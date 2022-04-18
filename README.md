# RestartExplorer
此脚本
  1.可以重启资源管理器并重新打开重启前的文件夹
  2.关闭重复的文件夹
  
​<img src="https://raw.githubusercontent.com/Yuphiz/Public/main/RestartExplorer/%E9%87%8D%E5%90%AF%E8%B5%84%E6%BA%90%E7%AE%A1%E7%90%86%E5%99%A8%E8%87%AA%E5%8A%A8%E6%89%93%E5%BC%80%E4%B8%8A%E6%AC%A1%E7%9B%AE%E5%BD%95.gif"  height = "450" alt="GUI demo" align=center />
  
此处演示的是后台版

### 使用方法：
  默认不带参数，双击运行：重启资源管理器并重新打开重启前的文件夹
  
  带参数：
  
  --CloseAll 关闭打开的所有文件夹窗口  
  --CloseDuplicate 关闭重复文件夹  
  --SetSchtasksAutoRun 设置开机启动、后台运行，可以在资源管理器崩溃重启后自动打开上次目录  
  --RemoveTasksch 移除开机启动、关闭后台  
  
 release压缩包自带了4个相对路径的快捷方式，方便使用  
	

### 更新日志：
### 更新 0.4  
	01.[增加] windows 11 dev版快速访问(主文件夹)换了新的guid，本次更新增加了新的guid
	02.[增加] 增加了后台检测资源管理器崩溃重启事件，只要触发重启事件，就自动打开上次的文件夹
	03.[修复] 增加了taskkill，用来修复当用户不存在tskill时重启资源管理器失败的bug
	04.[修复] 其他小优化修复
	05.改名为 restartexplorer
  
### 更新 0.3  
	01.[修复] 修复当ie窗口打开时使用脚本报错
  
### 更新 0.2 
	01.[修复] 修复带有空格的文件不能重启的错误
	02.整合成单文件

   
个人开发不易，如果觉得解决了你的问题，请捐赠支持开发者。你的支持将会让工具越来越好
![image](https://github.com/Yuphiz/Public/blob/main/Yuphiz_Pay.jpg)
