#Requires AutoHotkey v2.0.18+

#Include GuiReSizer.ahk
#Include ExcelHook.ahk

#SingleInstance Force

;global vars


; program entrance
app:=Program()
app.ShowGUI()

/**@var {ExcelHook} hook */
hook:=ExcelHook()
hook.HookExcelApp()
hook.Observer:=app
app.SetExcelHook(hook)
hook.ListenExcelEvents()



class Program{

    __New(){
        this.ui:=this.MakeGUI()
        this.hook:=object
    }
    MakeGUI(){
        ui:=Gui('+Resize +AlwaysOnTop '),ui.Opt('+MinSize200x200')
        
        ui.MarginX:=ui.MarginY:=5
        ui.Title:='NaviHelper'
        ;ui.SetFont('S11','Arial')
        ui.OnEvent('Size', GuiReSizer)

        ui.controls:={}
        listbox:=ui.controls.sheetListbox:= ui.AddListBox('section w200 h200')
        listbox.SetFont('S11','Arial')
        ;listbox.OnEvent('DoubleClick',OnListboxDbClick)
        listbox.Width:= -70
        listbox.Height:= -5

        vis :=ui.controls.visibleRadio:=ui.AddRadio('section ys', '可见')
        hid :=ui.controls.hiddenRadio:=ui.AddRadio('', '隐藏')
        vhid:=ui.controls.veryHiddenRadio:=ui.AddRadio('', '深度')
        vis.X:=hid.X:=vhid.X:=-60

        btn1:=ui.controls.assignButton:=ui.AddButton('xs','指定')
        btn2:=ui.controls.refreshButton:=ui.AddButton('xs y+1','刷新')
        btn2.OnEvent('Click',(*)=>Pause(-1))
        btn1.X:=btn2.X:=-60

        ui.OnEvent('Close',(*)=>ExitApp())
        return ui
    }

    ShowGUI(){
        this.ui.show('AutoSize')
    }

    SetExcelHook(hook){
        this.hook:=hook
    }

    OnWorkbookActivate(wb){
        Debug A_ThisFunc, '收到转发事件 WorkbookActivate'
        shlist:=this.hook.GetSheetList()
        list:=[]
        this.ui.controls.sheetListbox.delete
        /**@var {SheetInfo} sh */
        for _,sh in shlist
            list.push sh.name
        this.ui.controls.sheetListbox.add list
    }

    OnSheetActivate(sh){
        Debug A_ThisFunc, '收到转发事件 SheetActivate'
    }

}