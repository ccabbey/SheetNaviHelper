#Requires AutoHotkey v2.0


Persistent

xl:=ComObjActive('excel.application')
hdl:=Handler()
ComObjConnect(xl, hdl)


class Handler{
    static EventWhitelist:= 'SheetActivate,' .
        'WorkbookNewSheet,' .
        'WorkbookActivate,' 

    __Call(Name, Params) {

        if !InStr(Handler.EventWhitelist,Name)
            return

        plist:=''
        for p in Params
            plist .= Type(p) ','
        msgbox name ', ' plist
    }
}