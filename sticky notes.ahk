#Requires AutoHotkey v2.0
#SingleInstance Force

/*
Project:    Sticky Notes
Author:     kunkel321
Tool used:  Claude AI
Version:    1-17-2025
Forum:      https://www.autohotkey.com/boards/viewtopic.php?f=83&t=135340
Repository: https://github.com/kunkel321/Stickies     

Hotkeys:
* Ctrl+Shift+N - Create new note
* Ctrl+Shift+S - Toggle main window visibility

Notes are created in the top/left of the screen, cascading down, so they don't hide each other.  The top portion of the note is the Drag Bar.  Drag from there to move the note around.   To edit a sticky note, double-click Drag Bar or right-click, then choose edit from context menu.  

Note will open in editor where user can access most of the note options.  After moving and editing notes, click "Save Status" on main window. Main window is hidden at start up, so toggle to view it.  Notes are saved to an ini file in same location as script file. The ini file "sticky_notes.ini" will be created in same location as script.

If a note's text has multiple lines, and any of those lines start with 
[] or
[x] then those lines of text will be made into checkbox controls.  

If you change the length of text after the note have already been made, or if you check/uncheck boxes, I recommend doing "Save Status" from the note context menu or the main form.  This will save all notes and properties to the ini file.  Then do "Load Notes" (or "ReLoad Notes".) IF the note size hasn't adjusted correctly. 

This script is unique because nearly every bit of code was created with AI prompts, then pasted into the ahk editor.  A great deal of human input was needed, but very little of the actual code was human-generated.  Edit: In later versions, more human code was added.
*/

; Hotkey Configuration.  Change hotkeys if desired.
class OptionsConfig {
    static TOGGLE_MAIN_WINDOW := "^+s"  ; Ctrl+Shift+S
    static NEW_NOTE := "^+n"            ; Ctrl+Shift+N
    static APP_ICON := "sticky.ico"
    static INI_FILE := "sticky_notes.ini"
}

; Global Constants
class StickyNotesConfig {
    static DEFAULT_WIDTH := 200
    ; Removed DEFAULT_HEIGHT since we want natural height
    static DEFAULT_BG_COLOR := "FFFFCC"  ; Light yellow
    static DEFAULT_FONT := "Arial"
    static DEFAULT_FONT_SIZE := 12
    static DEFAULT_FONT_COLOR := "000000"  ; Black
    
    ; Color options for notes (expanded pastel shades)
    static COLORS := Map(
        "Light Yellow", "FFFFCC",
        "Light Pink", "FFE4E1",
        "Light Blue", "E6F3FF",
        "Light Green", "E8FFE8",
        "Light Purple", "E6E6FA",
        "Light Orange", "FFE4C4",
        "Light Cyan", "E0FFFF",
        "Light Coral", "FFE4E1",
        "Pale Gold", "FFE4B5",
        "Mint", "F5FFFA",
        "Lavender", "E6E6FA",
        "Peach", "FFDAB9",
        "Soft Red", "FFB6B6",
        "Soft Teal", "B6E6E6",
        "Soft Lime", "E6FFB6",
        "Soft Rose", "FFB6D9",
        "Soft Sky", "B6D9FF",
        "Soft Violet", "E6B6FF",
        "Soft Salmon", "FFD1B6",
        "Soft Sage", "D1E6B6"
    )

    ; Font color options (expanded darker shades)
    static FONT_COLORS := Map(
        "Black", "000000",
        "Navy Blue", "000080",
        "Dark Green", "006400",
        "Dark Red", "8B0000",
        "Purple", "800080",
        "Brown", "8B4513",
        "Dark Gray", "696969",
        "Maroon", "800000",
        "Dark Olive", "556B2F",
        "Indigo", "4B0082",
        "Dark Slate", "2F4F4F",
        "Dark Brown", "654321",
        "Deep Red", "B22222",
        "Forest Green", "228B22",
        "Dark Purple", "663399",
        "Deep Blue", "00008B",
        "Dark Teal", "008080",
        "Dark Magenta", "8B008B",
        "Burgundy", "800020",
        "Dark Orange", "D2691E"
    )
    
    ; Font size options
    static FONT_SIZES := [32, 28, 24, 20, 18, 16, 14, 12, 11, 10, 9, 8, 6]
}

; Main Application Class
class StickyNotes {
    ; Static properties for tracking application state
    static noteCount := 0
    static notes := Map()
    
    __New() {
        ; Initialize components
        this.InitializeComponents()
        
        ; Set up system tray
        this.SetupSystemTray()

        ; Set up hotkeys
        this.SetupHotkeys()
        
        ; Load saved notes but keep main window hidden
        this.LoadNotesOnStartup()
    }
    
    InitializeComponents() {
        ; Create main window
        try {
            this.mainWindow := MainWindow()
        } catch as err {
            MsgBox("Error initializing main window: " err.Message)
            ExitApp()
        }
        
        ; Initialize note manager
        try {
            this.noteManager := NoteManager()
        } catch as err {
            MsgBox("Error initializing note manager: " err.Message)
            ExitApp()
        }
    }

    SetupHotkeys() {
        ; Bind hotkeys to methods
        HotKey(OptionsConfig.TOGGLE_MAIN_WINDOW, (*) => this.ToggleMainWindow())
        HotKey(OptionsConfig.NEW_NOTE, (*) => this.CreateNewNote())
    }
    
    ToggleMainWindow(*) {
        if WinExist("ahk_id " this.mainWindow.gui.Hwnd) {
            this.mainWindow.Hide()
        } else {
            this.mainWindow.Show()
        }
    }
    
    CreateNewNote(*) {
        this.noteManager.CreateNote()
    }
    
    LoadNotesOnStartup() {
        this.noteManager.LoadSavedNotes()
    }

    SetupSystemTray() {
        ; Basic tray menu setup - will be expanded in SystemTrayManager component
        appName := StrReplace(A_ScriptName, ".ahk") 
        A_TrayMenu.Delete()  ; Clear default menu
        A_TrayMenu.Add(appName, (*) => False) ; Shows name of app at top of menu.
        A_TrayMenu.Add() 
        A_TrayMenu.Add("Show Main Window", (*) => this.mainWindow.Show())
        A_TrayMenu.Add("New Sticky Note", (*) => this.CreateNewNote())
        A_TrayMenu.Add()  ; Separator
        A_TrayMenu.AddStandard   ; Put the standard menu items back.
        A_TrayMenu.Add()  ; Separator
        A_TrayMenu.Add("Start with Windows", (*) => this.StartUpStickies())
        if FileExist(A_Startup "\sticky notes.lnk")
            A_TrayMenu.Check("Start with Windows")
        A_TrayMenu.Add("Open Note ini File", (*) => Run(OptionsConfig.INI_FILE))

        A_TrayMenu.Default := appName ; Set default tray icon
        ;TraySetIcon("imageres.dll",279) ; Yellow sticky with green 'up' arrow
        TraySetIcon(OptionsConfig.APP_ICON)
    }

    StartUpStickies() {
        if FileExist(A_Startup "\sticky notes.lnk") {
            FileDelete(A_Startup "\sticky notes.lnk")
            MsgBox("Sticky Notes will NO LONGER auto start with Windows.",, 4096)
        }
        Else {
            FileCreateShortcut(A_WorkingDir "\sticky notes.exe", A_Startup "\sticky notes.lnk", A_WorkingDir,,,,,)
            MsgBox("Sticky Notes will auto start with Windows.",, 4096)
        }
        Reload()
    }

    ExitApp() {
        ; Cleanup and save before exit
        this.noteManager.SaveAllNotes()
        ExitApp()
    }
}

; Create and start the application
global app := StickyNotes()

class Note {
    ; Note properties
    gui := ""            ; Keep this as empty string
    dragArea := ""       ; Keep this as empty string
    id := 0
    content := ""
    bgcolor := ""
    font := ""
    fontSize := 0
    isBold := false
    fontColor := ""
    isOnTop := false
    width := StickyNotesConfig.DEFAULT_WIDTH
    editor := ""
    controls := []
    currentX := ""
    currentY := ""

__New(idNum, options) {
    ; Store note ID and properties
    this.id := idNum
    this.content := options.content
    this.bgcolor := options.bgcolor
    this.font := options.font
    this.fontSize := options.fontSize
    this.isBold := options.HasOwnProp("isBold") ? options.isBold : false
    this.fontColor := options.fontColor
    this.isOnTop := options.isOnTop
    this.width := options.HasOwnProp("width") ? options.width : StickyNotesConfig.DEFAULT_WIDTH
    
    ; Create the basic GUI
    if (this.isOnTop) {
        this.gui := Gui("-Caption +AlwaysOnTop +Owner")
    } else {
        this.gui := Gui("-Caption +Owner")
    }
    
    this.gui.BackColor := this.bgcolor
    
    ; Create drag area
    this.dragArea := this.gui.Add("Text", 
        "x0 y0 w" this.width " h20 Background" this.bgcolor)
        
    ; Set default font
    this.gui.SetFont("s" this.fontSize (this.isBold ? " bold" : ""), this.font)
    
    ; Create note content
    if (InStr(this.content, "[]") || InStr(this.content, "[x]")) {
        this.CreateComplexNote()
    } else {
        this.CreateSimpleNote()
    }
    
    ; Show GUI with position
    if (options.HasOwnProp("x") && options.HasOwnProp("y") 
        && options.x != "" && options.y != ""
        && options.x is Integer && options.y is Integer) {
        this.gui.Show("x" options.x " y" options.y)
    } else {
        this.gui.Show()
    }
    
    ; Set up events
    this.SetupEvents()
    this.UpdatePosition()
}

    ; Helper method to get object keys
    GetObjectKeys(obj) {
        keys := ""
        for key in obj.OwnProps() {
            keys .= key ", "
        }
        return keys
    }

    CreateSimpleNote() {
        ; Add single text control for simple notes
        txt := this.gui.Add("Text", 
            "y25 x5 w" this.width,
            this.content)
        txt.SetFont("c" this.fontColor (this.isBold ? " bold" : ""))
        
        ; Store control reference
        this.controls := [{
            control: txt,
            type: "text",
            text: this.content
        }]
    }

    CreateComplexNote() {
        ; Initialize control array
        this.controls := []
        currentY := 25  ; Start below drag area
        
        ; Parse content for checkboxes
        parsed := this.ParseCheckboxContent()
        
        ; Create controls
        for item in parsed {
            if (item.type == "checkbox") {
                ; Create checkbox with proper font color
                cb := this.gui.Add("Checkbox",
                    "y" currentY " x5 w" this.width " c" this.fontColor,  ; Add color directly in options
                    item.text)
                cb.Value := item.checked
                
                ; Apply bold if needed
                if (this.isBold) {
                    cb.SetFont("bold")
                }
                
                ; Store checkbox reference and text
                this.controls.Push({
                    control: cb,
                    type: "checkbox",
                    text: item.text,
                    fullText: item.fullText
                })
                
                ; Make checkbox interactive
                cb.OnEvent("Click", this.SaveCheckboxState.Bind(this))
                
                currentY += 25
            } else {
                ; Create text with proper font color
                txt := this.gui.Add("Text", 
                    "y" currentY " x5 w" this.width " c" this.fontColor,  ; Add color directly in options
                    item.text)
                
                ; Apply bold if needed
                if (this.isBold) {
                    txt.SetFont("bold")
                }
                
                ; Store text control reference
                this.controls.Push({
                    control: txt,
                    type: "text",
                    text: item.text
                })
                
                currentY += 20
            }
        }
    }
                    
    ParseCheckboxContent() {
        parsed := []
        
        ; Split content into lines
        loop parse, this.content, "`n", "`r" {
            currentLine := A_LoopField
            
            ; Skip empty lines
            if (currentLine = "") {
                parsed.Push({
                    type: "text",
                    text: ""
                })
                continue
            }
            
            ; Check for checkbox formatting
            if (SubStr(currentLine, 1, 2) = "[]" || SubStr(currentLine, 1, 3) = "[x]") {
                isChecked := SubStr(currentLine, 2, 1) = "x"
                ; Always get text after `] `, regardless of checked state
                afterBracket := InStr(currentLine, "] ")
                if (afterBracket) {
                    lineText := SubStr(currentLine, afterBracket + 2)
                    ; Only create checkbox if there's actual text
                    if (lineText != "") {
                        parsed.Push({
                            type: "checkbox",
                            text: lineText,  ; Store just the text part
                            checked: isChecked,
                            fullText: currentLine  ; Store the complete line for reference
                        })
                        continue
                    }
                }
            }
            
            ; If we get here, treat as regular text
            parsed.Push({
                type: "text",
                text: currentLine
            })
        }
        
        return parsed
    }
    
    SetupEvents() {
        ; Set up right-click menu
        this.gui.OnEvent("ContextMenu", this.ShowContextMenu.Bind(this))
        
        ; Set up drag handling
        this.dragArea.OnEvent("Click", this.StartDrag.Bind(this))
        
        ; Add double-click handlers both using "DblClick"
        ;this.gui.OnEvent("DoubleClick", this.DoubleClickHandler.Bind(this))
        this.dragArea.OnEvent("DoubleClick", this.DoubleClickHandler.Bind(this))
    }

    DoubleClickHandler(*) {
        ;FileAppend("Double-click detected`n", "error_log.txt")
        this.Edit()
    }

    
    StartDrag(*) {
        
        SetWinDelay(-1)
        CoordMode("Mouse", "Screen")
        
        startWinX := 0
        startWinY := 0
        startMouseX := 0
        startMouseY := 0
        currentMouseX := 0
        currentMouseY := 0
        
        WinGetPos(&startWinX, &startWinY, , , this.gui)
        MouseGetPos(&startMouseX, &startMouseY)
        
        while GetKeyState("LButton", "P") {
            MouseGetPos(&currentMouseX, &currentMouseY)
            WinMove(
                startWinX + (currentMouseX - startMouseX),
                startWinY + (currentMouseY - startMouseY),
                , , this.gui
            )
        }
        
        ; Update position after drag
        this.UpdatePosition()
        
        SetWinDelay(100)
        CoordMode("Mouse", "Window")
    }
    

    ShowContextMenu(*) {
        noteMenu := Menu()
        noteMenu.Add("Edit this Note", this.Edit.Bind(this))
        noteMenu.Add("Hide this Note", this.Hide.Bind(this))
        noteMenu.Add("Delete this Note", this.Delete.Bind(this))
        noteMenu.Add()
        noteMenu.Add("Show Main Window", (*) => app.mainWindow.Show())
        noteMenu.Add("New Sticky Note", (*) => app.noteManager.CreateNote())
        noteMenu.Add()
        noteMenu.Add("Save Status", (*) => app.noteManager.SaveAllNotes())
        noteMenu.Add("Show Hidden", (*) => app.noteManager.ShowHiddenNotes())
        noteMenu.Add("ReLoad Notes", (*) => app.noteManager.LoadSavedNotes())

        noteMenu.Show()
    }
    
    SaveCheckboxState(ctrl, *) {
        ; Get checkbox text and value
        checkboxText := ctrl.Text
        isChecked := ctrl.Value
        
        ; Update the content line by line
        newLines := []
        foundMatch := false
        
        loop parse, this.content, "`n", "`r" {
            currentLine := A_LoopField
            
            ; Check if this line is a checkbox line
            if (SubStr(currentLine, 1, 2) = "[]" || SubStr(currentLine, 1, 3) = "[x]") {
                ; Get the text part after "] "
                afterBracket := InStr(currentLine, "] ")
                if (afterBracket) {
                    lineText := SubStr(currentLine, afterBracket + 2)
                    if (lineText = checkboxText) {
                        ; Found matching checkbox - update state
                        newLines.Push((isChecked ? "[x]" : "[]") " " checkboxText)
                        foundMatch := true
                        continue
                    }
                }
            }
            
            ; If no match or not a checkbox line, keep as is
            newLines.Push(currentLine)
        }
        
        ; Update content if we found and updated the checkbox
        if (foundMatch) {
            this.content := ""
            for line in newLines {
                this.content .= (A_Index > 1 ? "`n" : "") line
            }
            
            ; Save to INI immediately
            storage := NoteStorage()
            storage.SaveNote(this)
        }
    }
    
    Edit(*) {
        ; Create editor if it doesn't exist
        if (!this.editor) {
            this.editor := NoteEditor(this)
        }
        
        ; Show editor
        this.editor.Show()
    }
            
    UpdateContent(newContent, options := "") {
        ; Update properties
        this.content := newContent
        if (options) {
            if (options.HasOwnProp("bgcolor"))
                this.bgcolor := options.bgcolor
            if (options.HasOwnProp("font"))
                this.font := options.font
            if (options.HasOwnProp("fontSize"))
                this.fontSize := options.fontSize
            if (options.HasOwnProp("fontColor"))
                this.fontColor := options.fontColor
            if (options.HasOwnProp("isOnTop"))
                this.isOnTop := !!options.isOnTop  ; Force boolean
            if (options.HasOwnProp("width"))
                this.width := options.width
        }
        
        ; Store current position
        x := 0
        y := 0
        this.gui.GetPos(&x, &y)
        
        ; Destroy old GUI and create new one
        this.gui.Destroy()
        
        ; Create new GUI with current AlwaysOnTop setting
        if (this.isOnTop) {
            this.gui := Gui("-Caption +AlwaysOnTop")
        } else {
            this.gui := Gui("-Caption")
        }
        
        this.gui.BackColor := this.bgcolor
        
        ; Create drag area
        this.dragArea := this.gui.Add("Text", 
            "x0 y0 w" this.width " h20 Background" this.bgcolor)
            
        ; Set default font
        this.gui.SetFont("s" this.fontSize, this.font)
        
        ; Create note content
        if (InStr(this.content, "[]") || InStr(this.content, "[x]")) {
            this.CreateComplexNote()
        } else {
            this.CreateSimpleNote()
        }
        
        ; Show GUI at previous position
        this.gui.Show("x" x " y" y)
        
        ; Set up events again
        this.SetupEvents()
        
        ; Update stored position
        this.UpdatePosition()
    }

    ValidatePosition(x, y) {
        try {
            ; Get screen dimensions
            MonitorGet(MonitorGetPrimary(), &left, &top, &right, &bottom)
            
            ; Check if position is within screen bounds
            if (x < left || x > right || y < top || y > bottom) {
                ; FileAppend(
                ;     "Position out of bounds - x: " x ", y: " y 
                ;     . " (screen: " left "," top "," right "," bottom ")`n",
                ;     "error_log.txt"
                ; )
                return false
            }
            return true
        } catch {
            return false
        }
    }



    Hide(*) {
        this.gui.Hide()
        storage := NoteStorage()
        storage.MarkNoteHidden(this.id)
    }
    
    Show(*) {
        this.gui.Show()
    }
    
    Delete(*) {
        if (MsgBox("Are you sure you want to delete this note?",, "YesNo") = "Yes") {
            ; Delete from storage first
            storage := NoteStorage()
            storage.DeleteNote(this.id)
            
            ; Then destroy the GUI
            this.Destroy()
        }
    }
    
    Destroy(*) {
        if (this.editor) {
            this.editor.Destroy()
        }
        this.gui.Destroy()
    }
        
    UpdatePosition(*) {
        try {
            if (!this.gui || !WinExist("ahk_id " this.gui.Hwnd)) {
                return false
            }
            
            x := 0
            y := 0
            this.gui.GetPos(&x, &y)
            this.currentX := x
            this.currentY := y
            return {x: x, y: y}
        } catch as err {
            return false
        }
    }
}

class NoteManager {
    notes := Map()
    ;noteCount := 0
    storage := NoteStorage()
    
        ; Static properties for cascade positioning
    static CASCADE_OFFSET := 20      ; Pixels to offset each new note
    static MAX_CASCADE := 10         ; Maximum number of cascaded notes before reset
    static currentCascadeCount := 0  ; Track number of cascaded notes
    static baseX := ""              ; Base X position (will be set on first use)
    static baseY := ""              ; Base Y position (will be set on first use)
    
    GetCascadePosition() {
        ; Initialize base position if not set
        if (NoteManager.baseX = "" || NoteManager.baseY = "") {
            ; Get primary monitor's work area (excludes taskbar)
            MonitorGetWorkArea(MonitorGetPrimary(), &left, &top, &right, &bottom)
            
            ; Set initial position near top-left, but not at edge
            NoteManager.baseX := left + 50
            NoteManager.baseY := top + 50
        }
        
        ; Reset cascade if maximum reached
        if (NoteManager.currentCascadeCount >= NoteManager.MAX_CASCADE) {
            NoteManager.currentCascadeCount := 0
        }
        
        ; Calculate new position
        x := NoteManager.baseX + (NoteManager.CASCADE_OFFSET * NoteManager.currentCascadeCount)
        y := NoteManager.baseY + (NoteManager.CASCADE_OFFSET * NoteManager.currentCascadeCount)
        
        ; Increment cascade counter
        NoteManager.currentCascadeCount++
        
        return {x: x, y: y}
    }
    
    CreateNote(options := "") {
        ; Generate timestamp ID for new note
        newId := FormatTime(A_Now, "yyyyMMddHHmmss")
        
        ; Get cascade position
        pos := this.GetCascadePosition()
        
        ; Set default options if none provided
        if !options {
            options := {
                content: "Right-click here, or double-click header, to edit. Drag header to move.",
                bgcolor: StickyNotesConfig.DEFAULT_BG_COLOR,
                font: StickyNotesConfig.DEFAULT_FONT,
                fontSize: StickyNotesConfig.DEFAULT_FONT_SIZE,
                fontColor: StickyNotesConfig.DEFAULT_FONT_COLOR,
                x: pos.x,           ; Use cascade position
                y: pos.y,           ; Use cascade position
                isOnTop: false,
                isAutoSize: false,
            }
        }
        
        ; Create new note
        try {
            newNote := Note(newId, options)
            this.notes[newId] := newNote
            
            ; Save initial state
            this.storage.SaveNote(newNote)
            
            return newId
        } catch as err {
            MsgBox("Error creating note: " err.Message "`n" err.Stack)
            return 0
        }
    }

    SaveAllNotes() {
        try {
            ; FileAppend("`nStarting SaveAllNotes...`n", "error_log.txt")
            ; FileAppend("Current number of notes: " this.notes.Count "`n", "error_log.txt")
            
            for id, note in this.notes {
                try {
                    ; Get and log current position
                    x := 0
                    y := 0
                    note.gui.GetPos(&x, &y)
                    ; FileAppend("Saving note " id " at position x=" x ", y=" y "`n", "error_log.txt")
                    
                    ; Update position properties
                    note.currentX := x
                    note.currentY := y
                    
                    ; Save to storage
                    this.storage.SaveNote(note)
                    
                } ;catch as err {
                ;     FileAppend("Error saving note " id ": " err.Message "`n", "error_log.txt")
                ; }
            }
            
            ; FileAppend("SaveAllNotes complete`n", "error_log.txt")
            return true
        } catch as err {
            ; FileAppend("Error in SaveAllNotes: " err.Message "`n", "error_log.txt")
            return false
        }
    }

    LoadSavedNotes() {
        try {
            ; FileAppend("`nStarting LoadSavedNotes...`n", "error_log.txt")
            
            ; Get all saved notes from storage
            savedNotes := this.storage.LoadAllNotes()
            ; FileAppend("Found " savedNotes.Length " notes in storage`n", "error_log.txt")
            
            ; First destroy all existing note GUIs and editors
            ; FileAppend("Current notes in memory: " this.notes.Count "`n", "error_log.txt")
            for id, existingNote in this.notes {
                try {
                    ; FileAppend("Destroying note " id "`n", "error_log.txt")
                    if (existingNote.editor) {
                        existingNote.editor.Destroy()
                    }
                    if (existingNote.gui) {
                        existingNote.gui.Destroy()
                    }
                } ;catch as err {
                ;     FileAppend("Error destroying note " id ": " err.Message "`n", "error_log.txt")
                ; }
            }
            
            ; Reset note tracking completely
            this.notes := Map()  ; Create new Map
            this.noteCount := 0
            
            ; Process each saved note
            for noteData in savedNotes {
                try {
                    ; FileAppend("Loading note " noteData.id " with position x=" noteData.x ", y=" noteData.y "`n", "error_log.txt")
                    
                    if (noteData.id > this.noteCount) {
                        this.noteCount := noteData.id
                    }
                    
                    ; Only create visible GUI if note is not hidden
                    if (!noteData.isHidden) {
                        newNote := Note(noteData.id, noteData)
                        this.notes[noteData.id] := newNote
                    }
                    
                } catch as err {
                    ; FileAppend("Error creating note " noteData.id ": " err.Message "`n", "error_log.txt")
                    continue
                }
            }
            
            ; FileAppend("LoadSavedNotes complete. Notes in memory: " this.notes.Count "`n", "error_log.txt")
            return true
        } catch as err {
            ; FileAppend("Error in LoadSavedNotes: " err.Message "`n", "error_log.txt")
            return false
        }
    }


    DeleteNote(id) {
        if !this.notes.Has(id) {
            return false
        }
        
        try {
            ; Get note reference
            note := this.notes[id]
            
            ; Remove from storage first
            this.storage.DeleteNote(id)
            
            ; Destroy note GUI
            note.Destroy()
            
            ; Remove from collection - THIS IS THE KEY ADDITION
            this.notes.Delete(id)
            
            return true
        } catch as err {
            MsgBox("Error deleting note " id ": " err.Message)
            return false
        }
    }
    
    HideNote(id) {
        if !this.notes.Has(id) {
            return false
        }
        
        try {
            ; Get note reference
            note := this.notes[id]
            
            ; Save current state
            this.storage.SaveNote(note)
            
            ; Hide the note
            note.Hide()
            
            return true
        } catch as err {
            MsgBox("Error hiding note " id ": " err.Message)
            return false
        }
    }
    
    ShowHiddenNotes() {
        try {
            ; Get list of hidden notes
            hiddenNotes := this.storage.GetHiddenNotes()
            
            if (hiddenNotes.Length < 1) {
                MsgBox("No hidden notes found.")
                return false
            }
            
            ; Create menu of hidden notes
            hiddenMenu := Menu()
            
            for noteData in hiddenNotes {
                ; Create a preview of the note content
                preview := StrLen(noteData.content) > 40 
                    ? SubStr(noteData.content, 1, 37) "..."
                    : noteData.content
                    
                ; Add menu item
                id := noteData.id  ; Local copy for closure
                ; menuText := "Note " id ": " preview
                menuText := "---> " preview
                hiddenMenu.Add(menuText, this.RestoreNote.Bind(this, id))
            }
            
            ; Show menu
            hiddenMenu.Show()
            return true
            
        } catch as err {
            MsgBox("Error showing hidden notes: " err.Message)
            return false
        }
    }
    
    RestoreNote(id, *) {
        try {
            ; Get note data from storage
            noteData := this.storage.LoadNote(id)
            if !noteData {
                throw Error("Could not load note data")
            }
            
            ; Create new note with saved data
            newNote := Note(id, noteData)
            this.notes[id] := newNote
            
            ; Update note count if necessary
            if (id > this.noteCount) {
                this.noteCount := id
            }
            
            ; Mark as not hidden
            this.storage.MarkNoteVisible(id)
            
            return true
        } catch as err {
            MsgBox("Error restoring note " id ": " err.Message)
            return false
        }
    }
    
    UpdateNotePositions() {
        ; Update positions of all visible notes
        for id, note in this.notes {
            note.UpdatePosition()
        }
    }
}

class NoteStorage {

    ; Helper method to generate note section name
    GetNoteSectionName(id) {
        return "Note-" id
    }
 
    SaveNote(note) {
        try {
            ; If the note doesn't have a timestamp ID, create one
            if !RegExMatch(note.id, "^\d{14}$") {  ; Check if id is not already a timestamp
                note.id := FormatTime(A_Now, "yyyyMMddHHmmss")
            }
            
            sectionName := this.GetNoteSectionName(note.id)
            
            ; First, normalize line endings to LF
            contentFormatted := note.content
            contentFormatted := StrReplace(contentFormatted, "`r`n", "`n")  ; Windows CRLF -> LF
            contentFormatted := StrReplace(contentFormatted, "`r", "`n")    ; Old Mac CR -> LF
            
            ; Then convert LF to literal '\n' for INI storage
            contentFormatted := StrReplace(contentFormatted, "`n", "\n")
            
            ; Save all at once to prevent partial writes
            IniWrite(contentFormatted, OptionsConfig.INI_FILE, sectionName, "Content")
            IniWrite(note.bgcolor, OptionsConfig.INI_FILE, sectionName, "Color")
            IniWrite(note.font, OptionsConfig.INI_FILE, sectionName, "Font")
            IniWrite(note.fontSize, OptionsConfig.INI_FILE, sectionName, "FontSize")
            IniWrite(note.isBold, OptionsConfig.INI_FILE, sectionName, "IsBold")
            IniWrite(note.fontColor, OptionsConfig.INI_FILE, sectionName, "FontColor")
            
            ; Save position
            WinGetPos(&x, &y, , , note.gui)
            IniWrite(x, OptionsConfig.INI_FILE, sectionName, "PosX")
            IniWrite(y, OptionsConfig.INI_FILE, sectionName, "PosY")
            
            ; Save settings
            IniWrite(note.isOnTop, OptionsConfig.INI_FILE, sectionName, "IsOnTop")
            IniWrite(note.width, OptionsConfig.INI_FILE, sectionName, "Width")
            
            return true
        } catch as err {
            MsgBox("Error saving note to INI: " err.Message)
            return false
        }
    }
    
    LoadNote(id) {
        try {
            sectionName := this.GetNoteSectionName(id)
            
            ; Check if note exists
            if !this.NoteExists(id) {
                return false
            }
            
            ; Load raw content from INI
            rawContent := IniRead(OptionsConfig.INI_FILE, sectionName, "Content", "")
            
            ; Convert literal '\n' back to actual newlines
            unescapedContent := StrReplace(rawContent, "\n", "`n")
            
            ; Load all note data
            noteData := {
                id: id,  ; Keep the timestamp ID
                content: unescapedContent,
                bgcolor: IniRead(OptionsConfig.INI_FILE, sectionName, "Color", StickyNotesConfig.DEFAULT_BG_COLOR),
                font: IniRead(OptionsConfig.INI_FILE, sectionName, "Font", StickyNotesConfig.DEFAULT_FONT),
                fontSize: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "FontSize", StickyNotesConfig.DEFAULT_FONT_SIZE)),
                isBold: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "IsBold", "0")),
                fontColor: IniRead(OptionsConfig.INI_FILE, sectionName, "FontColor", StickyNotesConfig.DEFAULT_FONT_COLOR),
                x: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "PosX", "")),
                y: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "PosY", "")),
                isOnTop: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "IsOnTop", "0")),
                width: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "Width", StickyNotesConfig.DEFAULT_WIDTH)),
                isHidden: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "Hidden", "0"))
            }
            
            return noteData
        } catch as err {
            return false
        }
    }

    static isLoading := false
    
    LoadAllNotes() {
        try {
            ; Prevent recursive loading
            if (NoteStorage.isLoading) {
                return []
            }
            
            NoteStorage.isLoading := true
            
            notes := []
            sections := IniRead(OptionsConfig.INI_FILE)
            if !sections {
                NoteStorage.isLoading := false
                return notes
            }
            
            loop parse, sections, "`n", "`r" {
                ; Update regex to match new timestamp format
                if RegExMatch(A_LoopField, "Note-(\d{14})", &match) {
                    noteId := match[1]
                    if noteData := this.LoadNote(noteId) {
                        notes.Push(noteData)
                    }
                }
            }
            
            NoteStorage.isLoading := false
            return notes
            
        } catch as err {
            NoteStorage.isLoading := false
            return []
        }
    }
    
    DeleteNote(id) {
        try {
            IniDelete(OptionsConfig.INI_FILE, this.GetNoteSectionName(id))
            return true
        } catch as err {
            MsgBox("Error deleting note from INI: " err.Message)
            return false
        }
    }
    
    GetHiddenNotes() {
        try {
            hiddenNotes := []
            
            sections := IniRead(OptionsConfig.INI_FILE)
            if !sections {
                return hiddenNotes
            }
            
            loop parse, sections, "`n", "`r" {
                ; Update regex to match new timestamp format
                if RegExMatch(A_LoopField, "Note-(\d{14})", &match) {
                    noteId := match[1]
                    
                    ; Check if note is marked as hidden
                    isHidden := IniRead(OptionsConfig.INI_FILE, A_LoopField, "Hidden", "0")
                    if (isHidden = "1") {
                        if noteData := this.LoadNote(noteId) {
                            hiddenNotes.Push(noteData)
                        }
                    }
                }
            }
            
            return hiddenNotes
        } catch as err {
            MsgBox("Error getting hidden notes: " err.Message)
            return []
        }
    }
    

    MarkNoteHidden(id) {
        try {
            sectionName := this.GetNoteSectionName(id)
            IniWrite("1", OptionsConfig.INI_FILE, sectionName, "Hidden")
            return true
        } catch as err {
            MsgBox("Error marking note as hidden: " err.Message)
            return false
        }
    }
    
    MarkNoteVisible(id) {
        try {
            sectionName := this.GetNoteSectionName(id)
            IniWrite("0", OptionsConfig.INI_FILE, sectionName, "Hidden")
            return true
        } catch as err {
            MsgBox("Error marking note as visible: " err.Message)
            return false
        }
    }
    
    NoteExists(id) {
        try {
            return IniRead(OptionsConfig.INI_FILE, this.GetNoteSectionName(id), "Content", "") != ""
        } catch {
            return false
        }
    }
}

class NoteEditor {
    ; Editor properties
    note := ""
    gui := ""
    editControl := ""
    
    __New(note) {
        this.note := note
        this.CreateGui()
    }
    
    CreateGui() {
        ; Create editor window - don't set window background color
        this.gui := Gui("+AlwaysOnTop", "Edit Note " this.note.id)
        
        ; Create edit control with matching background color and correct bold state
        this.editControl := this.gui.Add("Edit",
            "x5 y5 w" StickyNotesConfig.DEFAULT_WIDTH " h200 +Multi +WantReturn Background" this.note.bgcolor,
            this.note.content)
            
        ; Set font properties including bold state
        this.editControl.SetFont(
            "s" this.note.fontSize 
            (this.note.isBold ? " bold" : " norm"), 
            this.note.font
        )
        this.editControl.SetFont("c" this.note.fontColor)
        
        ; Add formatting controls
        this.AddFormattingControls()
        
        ; Add save/cancel buttons
        this.AddActionButtons()
        
        ; Event handlers
        this.gui.OnEvent("Close", (*) => this.Hide())
    }

    AddFormattingControls() {
        ; Create format group below the edit control
        this.gui.Add("GroupBox", 
            "x5 y215 w" StickyNotesConfig.DEFAULT_WIDTH " h125",
            "Formatting")
        
        ; Background color dropdown
        this.gui.Add("Text", "xp+10 yp+20", "Background:")
        colorDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w100", this.GetColorList())
        colorDropdown.Text := this.GetColorName(this.note.bgcolor)
        colorDropdown.OnEvent("Change", (*) => this.UpdateBackgroundColor(colorDropdown.Text))
        
        ; Font dropdown - full width
        this.gui.Add("Text", "x15 y+10", "Font:")
        fontDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w" (StickyNotesConfig.DEFAULT_WIDTH - 50),
            ["Arial", "Times New Roman", "Verdana", "Courier New", "Comic Sans MS", "Calibri", "Segoe UI", "Georgia", "Tahoma"])
        fontDropdown.Text := this.note.font
        fontDropdown.OnEvent("Change", (*) => this.UpdateFont(fontDropdown.Text))
        
        ; Font size and bold controls on same line
        this.gui.Add("Text", "x15 y+10", "Size:")
        sizeDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w60", StickyNotesConfig.FONT_SIZES)
        sizeDropdown.Text := this.note.fontSize
        sizeDropdown.OnEvent("Change", (*) => this.UpdateFontSize(sizeDropdown.Text))
        
        ; Add Bold checkbox next to size
        boldCB := this.gui.Add("Checkbox", "x+15 yp+3", "Bold")
        boldCB.Value := this.note.isBold
        boldCB.OnEvent("Click", (*) => this.UpdateBold(boldCB.Value))
        
        ; Font color dropdown
        this.gui.Add("Text", "x15 y+10", "Font Color:")
        fontColorDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w100", this.GetFontColorList())
        fontColorDropdown.Text := this.GetFontColorName(this.note.fontColor)
        fontColorDropdown.OnEvent("Change", (*) => this.UpdateFontColor(fontColorDropdown.Text))
    }

    UpdateBold(isBold) {
        this.note.isBold := isBold
        this.editControl.SetFont(isBold ? "bold" : "norm")
    }
    
    GetColorList() {
        colorList := []
        for name, code in StickyNotesConfig.COLORS {
            colorList.Push(name)
        }
        return colorList
    }
    
    GetFontColorList() {
        colorList := []
        for name, code in StickyNotesConfig.FONT_COLORS {
            colorList.Push(name)
        }
        return colorList
    }
    
    GetColorName(colorCode) {
        for name, code in StickyNotesConfig.COLORS {
            if (code = colorCode)
                return name
        }
        return "Light Yellow"  ; Default
    }
    
    GetFontColorName(colorCode) {
        for name, code in StickyNotesConfig.FONT_COLORS {
            if (code = colorCode)
                return name
        }
        return "Black"  ; Default
    }
    
    UpdateBackgroundColor(colorName) {
        if (colorCode := StickyNotesConfig.COLORS[colorName]) {
            ; Update note's stored color
            this.note.bgcolor := colorCode
            
            ; Update edit control's background color
            this.editControl.Opt("Background" colorCode)
            
            ; Force redraw of the edit control
            this.editControl.Redraw()
        }
    }  
    
    UpdateFont(fontName) {
        this.note.font := fontName
        this.editControl.SetFont(, fontName)
    }
    
    UpdateFontSize(size) {
        this.note.fontSize := size
        this.editControl.SetFont("s" size)
    }
    
    UpdateFontColor(colorName) {
        if (colorCode := StickyNotesConfig.FONT_COLORS[colorName]) {
            this.note.fontColor := colorCode
            this.editControl.SetFont("c" colorCode)
        }
    }
    
   
    AddActionButtons() {
        ; Add Note Options group
        this.gui.Add("GroupBox", 
            "x5 y+10 w" StickyNotesConfig.DEFAULT_WIDTH " h70",  ; Made taller for both controls
            "Note Options")
        
        ; Add Width control with UpDown
        this.gui.Add("Text", "xp+10 yp+20", "Width:")
        widthEdit := this.gui.Add("Edit", "x+5 yp-3 w60 Number", this.note.width)
        widthUpDown := this.gui.Add("UpDown", "Range50-500", this.note.width)  ; Added range
        widthEdit.OnEvent("Change", (*) => this.UpdateWidth(widthEdit.Value))
        
        ; Add Always on Top checkbox below
        alwaysOnTopCB := this.gui.Add("Checkbox", "x15 y+10", "Always on Top")
        alwaysOnTopCB.Value := this.note.isOnTop
        alwaysOnTopCB.OnEvent("Click", (*) => this.ToggleAlwaysOnTop(alwaysOnTopCB.Value))
        
        ; Add Save/Cancel buttons at the bottom
        this.gui.Add("Button", "x5 y+15 w60", "Save")
            .OnEvent("Click", (*) => this.Save())
        
        this.gui.Add("Button", "x+10 w60", "Cancel")
            .OnEvent("Click", (*) => this.Hide())
    }

    ; And update the ToggleAlwaysOnTop method to actually set the property
    ToggleAlwaysOnTop(value) {
        this.note.isOnTop := value ? true : false
    }

    UpdateWidth(newWidth) {
        ; Ensure width is a positive number
        width := Max(50, Integer(newWidth))  ; Minimum width of 50 to prevent too-narrow notes
        
        ; Update stored width
        this.note.width := width
    }
    
    Save(*) {
        ; Get current content
        newContent := this.editControl.Text
        
        ; Update note with all properties including width
        this.note.UpdateContent(newContent, {
            bgcolor: this.note.bgcolor,
            font: this.note.font,
            fontSize: this.note.fontSize,
            isBold: this.note.isBold,
            fontColor: this.note.fontColor,
            isOnTop: this.note.isOnTop,  ; Added missing comma here
            width: this.note.width
        })
        
        ; Save to storage immediately
        (NoteStorage()).SaveNote(this.note)
        
        ; Hide editor
        this.Hide()
    }
    
    Show(*) {
        this.gui.Show("AutoSize")
    }
    
    Hide(*) {
        this.gui.Hide()
    }
    
    Destroy(*) {
        this.gui.Destroy()
    }
}

class MainWindow {
    ; Window properties
    gui := ""
    
    __New() {
        this.CreateGui()
    }
    
    CreateGui() {
        ; Create main window
        this.gui := Gui("+AlwaysOnTop +Resize", "Sticky Notes Manager")
        
        ; Add buttons vertically
        this.gui.Add("Button", "w120 h30", "New Note")
            .OnEvent("Click", (*) => app.noteManager.CreateNote())
            
        this.gui.Add("Button", "w120 h30 y+5", "Load Notes")
            .OnEvent("Click", (*) => app.noteManager.LoadSavedNotes())
            
        this.gui.Add("Button", "w120 h30 y+5", "Show Hidden")
            .OnEvent("Click", (*) => app.noteManager.ShowHiddenNotes())
            
        this.gui.Add("Button", "w120 h30 y+5", "Save Status")
            .OnEvent("Click", (*) => app.noteManager.SaveAllNotes())
            
        this.gui.Add("Button", "w120 h30 y+5", "Exit App")
            .OnEvent("Click", (*) => this.ExitApp())
        
        ; Add help text
        helpText := "Right-click notes to edit them`nDrag notes to move them"
        this.gui.Add("Text", "y+10 w120", helpText)
        
        ; Set up events
        this.gui.OnEvent("Close", (*) => this.ExitApp())
        this.gui.OnEvent("Escape", (*) => this.Hide())

        ; this.gui.Opt("+Owner") ; enable this to prevent main gui taskbar icon.
        
        ; Show the window
        ; this.Show() ; disable this to keep window from showing at startup. 
    }
    
    Show(*) {
        ; Position window near top of screen
        this.gui.Show("AutoSize y50")
    }
    
    Hide(*) {
        this.gui.Hide()
    }
    
    Destroy(*) {
        this.gui.Destroy()
    }
    
    ExitApp(*) {
        ; Save all notes state before exiting
        app.noteManager.SaveAllNotes()
        ExitApp()
    }
}
