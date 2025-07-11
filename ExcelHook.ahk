#Requires AutoHotkey v2.0

;events enum

XL_SheetActivate := 'SheetActivate'
XL_WorkbookNewSheet := 'WorkbookNewSheet'
XL_WorkbookActivate := 'WorkbookActivate'
XL_HookReady := 'HookReady'   ; when excel hook is established

/**
 * @description ExcelHook
 * Excel钩子类, 接管活动Excel进程的事件。  
 * 所有对Excel的操作都应该通过ExcelHook执行。  
 */
class ExcelHook {

    __New() {
        this.Excel := Object
        this.EventHandler := EventHandler()
        this.Observer := Object
    }

    /**
     * 自动注册观察者并立即开始监听Excel事件。成功后会主动广播一次 HookReady 事件。
     * @param observer 拥有event处理方法的客户端观察者
     */
    Auto(observer) {
        this.SetObserver(observer)
        this.HookExcelApp()
        ; 设置定时任务监视Excel是否被用户关闭
        SetTimer(ObjBindMethod(this, 'ProcessDaemon'), 5000)
        this.Listen('On')

    }

    AddEventsToListen(events*) {
        for event in events
            if !InStr(this.EventHandler.ListenedEvents, event, 0)
                this.EventHandler.ListenedEvents .= event . ','
    }

    RemoveEventsToListen(events*) {
        for event in events
            if !InStr(this.EventHandler.ListenedEvents, event, 0)
                this.EventHandler.ListenedEvents := StrReplace(this.EventHandler.ListenedEvents, event ',')
    }

    /**
     * 注册观察者并建立双向引用
     * @param observer 拥有event处理方法的客户端观察者
     */
    SetObserver(observer) {
        this.EventHandler.Observer := observer
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
        if !WinExist('ahk_exe excel.exe') {
            Debug(A_ThisFunc, 'Excel process does not exist, waiting...')
            WinWait('ahk_exe excel.exe')
            Sleep 250   ; will case 'Error: (0x800401E3)' if no waiting
        }
        ; 当存在多个Excel进程时, ComObjActive会优先绑定更早创建的那个Excel。
        this.Excel := ComObjActive('Excel.Application')
        Debug A_ThisFunc, 'Excel COM Object is HOOKED.'
        ;判断excel进程是否是一个空壳，如果无法访问Sheets.Count说明没有活动工作簿
        try {
            test := this.excel.sheets.count
            this.Listen 'On'
            this.EventHandler.RelayEvent(XL_HookReady, this.Excel.Activeworkbook)
        }

    }

    /** @method - ProcessDaemon
     * 监视Excel是否正在运行, 如果被用户手动关闭, 则释放Excel COM对象引用。
     */
    ProcessDaemon() {
        static prompted := false
        if !WinExist('ahk_exe excel.exe') {
            this.Excel := ''
            Debug A_ThisFunc, 'Excel process closed, COM obj released.'
            prompted := true
            ; 尝试绑定新的Excel进程
            this.HookExcelApp()
        }
        else {
            if !prompted {
                Debug A_ThisFunc, 'Excel process exist...'
                prompted := true
            }
            ; Do nothing, just log
        }
    }

    /** Listen
     * 通过COM接口监听EXCEL事件
     * @param {String} OnOff - On : 开始监听; Off : 停止监听  
     * 注意: 脚本必须保留对 ComObj 的引用, 否则它将被自动释放并断开与其 COM 对象的连接, 从而阻止检测到任何进一步的事件。
     */
    Listen(OnOff := 'On') {
        if OnOff == 'On' {
            ComObjConnect(this.Excel, this.EventHandler)
            Debug A_ThisFunc, 'Listen Started...'
        }
        else {
            ComObjConnect(this.Excel)
            Debug A_ThisFunc, 'Listen Stoped...'
        }
    }

    ActivateSheet(SheetName) {
        this.excel.sheets(SheetName).activate
    }
}

/** @class EventHandler
 * 捕获预定义的Excel Application或Workbook事件和上下文。 
 */
class EventHandler {

    __New() {
        this.ListenedEvents := ''
        this.Observer := unset
    }
    /**
     * 当Excel COM对象广播任意事件时，ExcelHook会捕获事件名称和上下文并由EventHandler处理。  
     * 如果事件名称在监视清单中，EventHandler会调用ExcelHook的广播方法将事件转发给观察者。
     * @param EventName 捕获的事件名称
     * @param args 事件的上下文，通常为Sheet/Workbook和Application对象
     */
    __Call(EventName, args) {
        if InStr(this.ListenedEvents, EventName) {
            Debug A_ThisFunc, 'Event captured: ' EventName
            this.RelayEvent(eventName, args[1])
        }

    }

    /** @method - RelayEvent
     * 将事件和上下文转发到ExcelHook的观察者，如果观察者具有与该事件相对应的回调方法。  
     * 例如: 如果事件名称为SheetActivate， 而观察者有对应的OnSheetActive方法，
     * 那么此事件和对应的上下文(此例中为WorkSheet对象)会转发给观察者。
     * @param {string} eventName - 事件名称
     * @param {object} args - 事件上下文
     */
    RelayEvent(eventName, obj?) {
        if this.Observer.HasMethod('On' eventName) {
            Debug A_ThisFunc, 'Foward ' eventName ' event to observer...'
            if IsSet(obj)
                this.Observer.On%eventName%(obj)
            else
                this.Observer.On%eventName%()
        }
        else {
            Debug A_ThisFunc, 'Client observer has no On' eventName ' method...'
        }
    }

}

/** @description
 * 输出到控制台
 * @param {String} caller 调用函数名
 * @param {String} message Debug消息
 */
Debug(caller := '', message := '') {
    OutputDebug A_ScriptName ' => ' StrReplace(caller, '.Prototype', '') ' => ' message
}
