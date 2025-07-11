#Requires AutoHotkey v2.0

/**
 * @class - SheetInfo  
 * Sheet实体类, 封装Sheet对象引用、表名、可见性、激活状态等信息。  
 */
class SheetInfo {
    static statusMap := Map(-1, '', 0, '[隐藏]', 2, '[深度隐藏]')
    SheetObj := Object
    Name := ''

    /** 附加了工作表可见性的表名。格式: Name[Visibility] */
    DisplayName := ''
    Visibility := ''
    Active := false

    /** @constructor
     * @param {WorkSheet} sh - 工作表对象引用
     */
    __New(sh) {
        this.SheetObj := sh
        this.Name := sh.Name
        this.Visibility := sh.visible
        this.Active := sh.Parent.activesheet.name == sh.name ? true : false
        this.DisplayName := this.Name . SheetInfo.statusMap[this.Visibility]
    }
}
