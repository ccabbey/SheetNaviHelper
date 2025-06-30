#Requires AutoHotkey v2.0


/** 
 * @class - SheetInfo  
 * Sheet实体类, 封装Sheet对象引用、表名、可见性、激活状态等信息。  
 */
class SheetInfo{
        static statusMap:=Map(-1,'',0,'[隐藏]',2,'[深度隐藏]')
        SheetObj:=Object
        Name:=''

        /** 附加了工作表可见性的表名。格式: Name[Visibility] */
        DisplayName:=''
        Visibility:=''
        Active:=false

        /** @constructor  
         * @param {WorkSheet} sh - 工作表对象引用
         */
        __New(sh){
            this.SheetObj:=sh
            this.Name:=sh.Name
            this.Visibility:=sh.visible
            this.Active:=sh.Parent.activesheet.name==sh.name?true:false
            this.DisplayName:=this.Name . SheetInfo.statusMap[this.Visibility]
        }
}


/** @class EventHandler 
 * 捕获预定义的Excel Application或Workbook事件和上下文。 
*/
class EventHandler{

    MonitoredEvents:= 'SheetActivate,WorkbookNewSheet,WorkbookActivate'

    /**
     * @constructor
     * @param {ExcelHook} hook - ExcelHook实例
    */
    __New(hook){
        this.Hook:=hook
    }

    /**
     * 当Excel COM对象广播任意事件时，ExcelHook会捕获事件名称和上下文并由EventHandler处理。  
     * 如果事件名称在监视清单中，EventHandler会调用ExcelHook的广播方法将事件转发给观察者。
     * @param EventName 捕获的事件名称
     * @param args 事件的上下文，通常为Sheet/Workbook和Application对象
     */
    __Call(EventName, args) {
        if InStr(this.MonitoredEvents,EventName){
            Debug A_ThisFunc, '捕获到事件' EventName
            this.Hook.RelayEvent(eventName,args[1])
        }
            
    }
}

/**
 * @description ExcelHook
 * Excel钩子类, 接管活动Excel进程的事件。  
 * 所有对Excel的操作都应该通过ExcelHook执行。  
*/
class ExcelHook {

    /** @constructor */
    __New() {
        /** Excel COM对象引用 */
        this.Excel:= Object

        /** 事件处理类对象引用 */
        this.EventHandler:= EventHandler(this)

        /** Observer - 应用层观察者
         * ExcelHook监听到的事件会转发给观察者进行后续处理 */
        this.Observer:=Object

    }

    /**
     * 自动注册观察者并立即开始监听Excel事件。成功后会手动广播一次 WorkbookActivate事件。
     * @param observer 拥有event处理方法的客户端观察者
     */
    Auto(observer){
        this.SetObserver(observer)
        this.HookExcelApp()
        this.ListenExcelEvents()
        ;判断excel进程是否是一个空壳，如果无法访问Sheets.Count说明没有活动工作簿
        try{
            test:=this.excel.sheets.count
            this.RelayEvent('WorkbookActivate',this.Excel.Activeworkbook)
        }
        catch{
            ; do nothing
        }
            
        
    }
    /**
     * 注册观察者并建立双向引用
     * @param observer 拥有event处理方法的客户端观察者
     */
    SetObserver(observer){
        this.Observer:=observer
        observer.Hook:=this
    }
    /** @method - HookExcelApp
     * @description  
     * 尝试绑定当前活动的Excel进程。  
     * 如果活动进程不存在, 程序会一直等待。
    */
    HookExcelApp() {
        ; 由于未知的原因，使用WinWait方法来检测Excel进程是否存在，会导致
        ;   后续的ComObjActive方法报错 => Error: (0x800401E3) 操作无法使用.
        ;   猜测是由于Excel此时还没有完成初始化COM对象，解决方案是加一个延时
        if !WinExist('ahk_exe excel.exe'){
            Debug(A_ThisFunc,'Excel进程不存在, 等待程序运行...')
            WinWait('ahk_exe excel.exe')
            Sleep 250 
        }
        ; 当存在多个Excel进程时, ComObjActive会优先绑定更早创建的那个Excel。
        ; 通过Excel.Workbooks.Count==0可以判断该进程是否是一个空壳。  
        this.Excel := ComObjActive('Excel.Application')
        Debug A_ThisFunc, '已绑定Excel COM对象' 
        ; 设置定时任务监视Excel是否被用户关闭
        SetTimer(ObjBindMethod(this,'ProcessDaemon'),10000) 
    }

    /** @method - ProcessDaemon  
     * 监视Excel是否正在运行, 如果被用户手动关闭, 则释放Excel COM对象引用。
     */
    ProcessDaemon(){
        if !WinExist('ahk_exe excel.exe'){
            this.Excel:=''
            Debug A_ThisFunc, 'Excel已关闭, COM对象引用已释放'
            SetTimer ,0 ; 取消调用者设置的定时任务
            ; 尝试绑定新的Excel进程
            this.HookExcelApp()
        }
        else{
            static prompted:=false
            if !prompted{
                Debug A_ThisFunc, 'Excel进程存在...'
                prompted:=true
            }
            ; Do nothing, just log
        }
    }

    /** @method - ListenExcelEvents
     * 通过COM接口监听EXCEL事件
     * @param {EventHandler} handler - 事件处理类, 其中包含对应的事件处理方法。 
     * @returns {ExcelHook}  
     * 注意: 脚本必须保留对 ComObj 的引用, 否则它将被自动释放并断开与其 COM 对象的连接, 从而阻止检测到任何进一步的事件。
     */
    ListenExcelEvents() {
        try{
            ComObjConnect(this.Excel, this.EventHandler)
            Debug A_ThisFunc, '开始监听Excel事件...'
        }
        catch{
        }
    }

    /** @method - RelayEvent  
     * 将事件和上下文转发到ExcelHook的观察者，如果观察者具有与该事件相对应的回调方法。  
     * 例如: 如果事件名称为SheetActivate， 而观察者有对应的OnSheetActive方法，
     * 那么此事件和对应的上下文(此例中为WorkSheet对象)会转发给观察者。
     * @param {string} eventName - 事件名称
     * @param {object} args - 事件上下文
     */
    RelayEvent(eventName,obj?){
        ;Debug A_ThisFunc, '检查观察者是否存在 ' eventName ' 事件的回调方法'
        ;msgbox !this.Observer ', ' this.Observer.HasMethod('On' eventName)
        if this.Observer && this.Observer.HasMethod('On' eventName){
            Debug A_ThisFunc, '转发 ' eventName ' 事件到客户端...'
            if IsSet(obj)
                this.Observer.On%eventName%(obj)
            else
                this.Observer.On%eventName%()
        }
        else{
            Debug A_ThisFunc, '客户端没有实现 On' eventName '方法...'
        }
    }

    /**
     * 通过ExcelHook获取活动工作簿的工作表信息对象集合。参见 SheetInfo 类。  
     * 如果仅需要获取工作表名称列表，使用GetSheetNameList()
     */
    GetSheetInfoList(){
        list:=[]
        for sh in this.Excel.ActiveWorkbook.WorkSheets
            list.Push SheetInfo(sh)
        return list
    }
}

/** @description
 * 输出到控制台
 * @param {String} caller 调用函数名
 * @param {String} message Debug消息
 */
Debug(caller, message){
    OutputDebug A_ScriptName ' => ' StrReplace(caller,'.Prototype','') ' => ' message
}

