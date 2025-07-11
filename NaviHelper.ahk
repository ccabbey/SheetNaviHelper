#Requires AutoHotkey v2.0.18+

#Include GuiReSizer.ahk
#Include ExcelHook.ahk

#SingleInstance Force

;global vars

; program entrance
app := Program()
app.ShowGUI()

hook := ExcelHook()
hook.AddEventsToListen(XL_SheetActivate, XL_WorkbookActivate)
hook.Auto(app)

class Program {

    __New() {
        this.ui := this.MakeGUI()
        this.hook := object
        this.SheetListCache := []
    }
    MakeGUI() {
        ui := Gui('+Resize +AlwaysOnTop '), ui.Opt('+MinSize200x200')

        ui.MarginX := ui.MarginY := 5
        ui.Title := 'NaviHelper'
        ;ui.SetFont('S11','Arial')
        ui.OnEvent('Size', GuiReSizer)

        ui.controls := {}
        sb := ui.controls.SearchEdit := ui.AddEdit('w200 r1', '')
        sb.SetFont('S11', 'Arial')
        sb.Width := -70
        sb.OnEvent('Change', (*) => this.OnSearchEditChange(sb.Text))
        lb := ui.controls.sheetListbox := ui.AddListBox('section w200 h200')
        lb.SetFont('S11', 'Arial')
        lb.OnEvent('DoubleClick', (*) => this.OnListboxDbClick())
        lb.Width := -70, lb.Height := -5

        rb1 := ui.controls.visibleRadio := ui.AddRadio('section ys', '可见')
        rb2 := ui.controls.hiddenRadio := ui.AddRadio('', '隐藏')
        rb3 := ui.controls.veryHiddenRadio := ui.AddRadio('', '深度')
        rb1.X := rb2.X := rb3.X := -60

        btn1 := ui.controls.assignButton := ui.AddButton('xs', '指定')
        btn2 := ui.controls.refreshButton := ui.AddButton('xs y+1', '刷新')
        btn2.OnEvent('Click', (*) => Pause(-1))
        btn1.X := btn2.X := -60

        ui.OnEvent('Close', (*) => ExitApp())
        return ui
    }

    ShowGUI() {
        this.ui.show('AutoSize')
        HotIfWinActive('ahk_class AutoHotkeyGUI')
        Hotkey('Tab', (*) => this.ui.controls.SearchEdit.Focus())
    }

    OnWorkbookActivate(wb) {
        Debug A_ThisFunc, 'Event Received: WorkbookActivate'
        this.UpdateListbox()
        this.ui.Title := wb.name
    }

    OnSheetActivate(sh) {
        Debug A_ThisFunc, 'Event Received: SheetActivate'
        this.UpdateListbox()
        this.SheetListCache := this.hook.GetSheetList('DisplayName')
    }

    OnWorkbookBeforeClose(wb) {
        Debug A_ThisFunc, 'Event Received: WorkbookBeforeClose'
        this.ui.controls.sheetListbox.delete()
    }

    OnHookReady(xl) {
        Debug A_ThisFunc, 'Event Received: HookReady'
    }
    UpdateListbox() {
        try {
            shlist := this.hook.GetSheetInfoList()
            list := []
            activeid := ''
            for sh in shlist {
                list.Push(sh.DisplayName)
                if sh.Active
                    activeid := A_Index
            }
            lb := this.ui.controls.sheetListbox
            lb.Delete()
            lb.Add(list)
            lb.Value := activeid
        }
    }

    OnListboxDbClick() {
        target := this.ui.controls.sheetListbox.Text
        this.hook.ActivateSheet(target)
    }

    OnSearchEditChange(keyword) {
        lb := this.ui.controls.sheetListbox
        if keyword = '' {
            lb.delete()
            lb.add this.SheetListCache
            return
        }

        ret := []
        for e in this.SheetListCache
            if InStr(e, keyword)
                ret.push e

        lb.delete()
        lb.add ret
    }
}
