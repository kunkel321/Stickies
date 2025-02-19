#Requires AutoHotkey v2.0
#SingleInstance Force

^Esc::ExitApp
/*
Project:    Sticky Notes
Author:     kunkel321
Tool used:  Claude AI
Version:    2-18-2025
Forum:      https://www.autohotkey.com/boards/viewtopic.php?f=83&t=135340
Repository: https://github.com/kunkel321/Stickies     

Features, Functionality, Usage, and Tips:
-----------------------------------------
- Create sticky notes that persist between script restarts
- Notes cascade from top-left of screen for better visibility
- Cascading Positions: New note doesn't fully cover previous note
- Notes cycle through different pastel background colors automatically
- Note font color can optionally cycle too, or just use black.
- Several formatting options including fonts, colors, sizing, and borders
- Visual customization: Border thickness changes with bold text
- Border color always matches font color. 
- Stick notes to windows: Notes can be attached to specific application windows
- Window persistence: Notes "stuck to" specific windows reappear when window reopens
- Notes can be unstuck by clicking the window button again
- Set alarms for individual notes with optional weekly recurrence
- Many editor dialog options support accelerator keys (Alt+A for alarm, etc.)
- Visual alert: Notes can shake when alarms trigger (configurable)
- Multiple alarm repeats: Choose between once, 3x, or 10x alarm repetitions
- Alarm sounds: Custom alarm sounds can be added to the Sounds folder
- Smart alarm management: System detects and reports missed alarms on startup
- Main window with note management, preview, and search functionality
- Color-coded notes: Notes maintain their color scheme in preview list
- Rich preview: Hover over notes in manager for formatted preview with original fonts/colors
- Search functionality: Filter notes by content in main window
- Selective display: Filter note list to show only hidden or visible notes
- Access main window with Win+Shift+S (hidden by default)
- Notes' show alarm times and window attachments in note manager listview
- Bulk operations: Select multiple notes to hide/unhide/delete simultaneously
- Notes are created using Win+Shift+N or from clipboard with Win+Shift+C
- Hotkeys can be changed near top of code
- Tips button in note manager shows hotkeys and other tips
- Drag notes by their top bar to reposition
- Configurable drag area: Option to maximize note space by minimizing drag area
- Notes auto-save position when moved
- Turn off note deletion warning
- Double-click top bar or right-click for editing options
- System tray icon provides quick access to common functions
- Start with Windows: Option available in tray menu
- Checkbox Creation: Any text line starting with [] or [x] becomes an interactive checkbox
- Checkbox Safety: Alt+Click required by default to prevent accidental toggles
- Alarm System: Set one-time or recurring alarms with custom sounds
- All note data saved to sticky_notes.ini in script directory
- Use manual "Save Status" after significant changes
- "Load/Reload Notes" refreshes all notes from storage
- Hidden notes can be restored through main window
- Check error_log.txt for troubleshooting (if enabled; warning: system hog)
- Right-click menu provides quick access to several functions
- Multiple selection: Use Ctrl+Click to select multiple notes in manager

Known Issues:
-------------
- In note listview, when previewing, must let preview time-out before previewing another
- In note listview, if notes are sorted, colors will not sort with them

Development Note:
-----------------
This script was developed primarily through AI-assisted coding, with Claude AI 
generating most of the base code structure and functionality. Later versions 
include additional human-written code for enhanced features and bug fixes.
Thanks go to the humans, Hellbent for his borderGui class, and 
Justme for his ToolTipOptions and LV_Colors classes. 
*/

; #+N:: Create new note
; #+C:: Create new note from clipboard text
; #+S:: Toggle main window visibility

; Hotkeys and other Configuration.  Change if desired.
class OptionsConfig {
    static ERROR_LOG             := false  ; Mostly for AI feedback.  Causes lag.  Recommend setting 'false.'      
    static TOGGLE_MAIN_WINDOW    := "+#s"  ; Shift+Win+S. Shows/Hide Note Manager window.
    static NEW_NOTE              := "+#n"  ; Shift+Win+N. New note.
    static NEW_CLIPBOARD_NOTE    := "+#c"  ; Shift+Win+C. New note from text on clipboard.
    static APP_ICON              := "sticky.ico" ; Homemade icon that kunkel321 made.
    static INI_FILE              := "sticky_notes.ini" ; The note storage file. 
    static AUTO_OPEN_EDITOR      := true   ; Should Note Editor auto open upon note creation?    
    static MAX_NOTE_WORDS        := 200    ; Maximum words in a clipboard note before truncating.
    static MAX_NOTE_LINES        := 35     ; Maximum lines in a clipboard note before truncating.
    static CHECKBOX_MODIFIER_KEY := "Alt"  ; To prevent accidental checkbox clicks.  "" = don't require modifer.
    static DEFAULT_BORDER        := true   ; Whether new notes have borders by default.
    static CYCLE_FONT_COLOR      := true   ; Should new notes have colored font by default? False = black.
    static WARN_ON_DELETE_NOTE   := false  ; Whether to show confirmation before deleting notes (except multiple deletion).
    ; Visual shake for alarms settings.
    static DEFAULT_ALARM_SHAKE   := true    ; Whether new notes have borders by default.
    static SHAKE_STEPS           := 40      ; Number of shakes (40 takes about 3 seconds, depending on computer).
    static MAX_SHAKE_DISTANCE    := 12      ; Maximum shake distance in pixels -- Start with this much shake.
    static MIN_SHAKE_DISTANCE    := 0       ; Minimum shake distance in pixels -- Fade down to this much.
    ; Sticky note drag area settings.
    static RESERVE_DRAG_AREA     := false   ; Whether to reserve space at top for drag area.
    static DRAG_AREA_OFFSET      := 25      ; Y-offset of note text when reserving drag area.
    static NO_RESERVE_OFFSET     := 5       ; Y-offset when not reserving drag area.
    static NO_RESERVE_X_OFFSET   := 25      ; X-offset for drag area when not reserving space (makes room to click checkboxes).
}

; Global Constants
class StickyNotesConfig {
    static DEFAULT_WIDTH        := 200
    ; Removed DEFAULT_HEIGHT since we want natural height
    static DEFAULT_FONT         := "Arial"
    static DEFAULT_FONT_SIZE    := 12
    
    ; Color options for notes (expanded pastel shades) new notes assume default color in this order, starting with a random item from the list.
    static COLORS := Map(
        "Light Yellow"  , "FFFFCC",
        "Soft Sky"      , "B6D9FF",
        "Pale Gold"     , "FFE4B5",
        "Soft Red"      , "FFB6B6",
        "Soft Lime"     , "E6FFB6",
        "Mint"          , "F5FFFA",
        "Soft Teal"     , "B6E6E6",
        "Lavender"      , "E6E6FA",
        "Light Green"   , "E8FFE8",
        "Peach"         , "FFDAB9",
        "Soft Rose"     , "FFB6D9",
        "Soft Violet"   , "E6B6FF",
        "Soft Salmon"   , "FFD1B6",
        "Light Pink"    , "FFE4E1",
        "Light Blue"    , "E6F3FF",
        "Light Purple"  , "E6E6FA",
        "Light Orange"  , "FFE4C4",
        "Light Cyan"    , "E0FFFF",
        "Light Coral"   , "FFE4E1",
        "Soft Sage"     , "D1E6B6"
    )
    static DEFAULT_BG_COLOR := "FFFFCC"  ; Light yellow

    ; Font color options (expanded darker shades) new notes assume default font color in this order, starting with a random item from the list.
    static FONT_COLORS := Map(
        "Black"     , "000000",
        "Navy Blue"     , "000080",
        "Dark Green"    , "015701",
        "Dark Red"      , "8B0000",
        "Purple"    , "800080",
        "Brown"         , "8B4513",
        "Dark Gray"     , "575757",
        "Maroon"    , "800000",
        "Dark Olive"    , "435523",
        "Indigo"        , "4B0082",
        "Dark Slate"    , "2F4F4F",
        "Dark Brown"    , "654321",
        "Deep Red"      , "B22222",
        "Forest Green"  , "1e701e",
        "Dark Purple"   , "562b82",
        "Deep Blue"     , "00008B",
        "Dark Teal"     , "036666",
        "Dark Magenta"  , "8B008B",
        "Burgundy"      , "800020",
        "Dark Orange"   , "a65012"
    )
    
    ; Font size options
    static FONT_SIZES := [32, 28, 24, 20, 18, 16, 14, 12, 11, 10, 9, 8, 6]
}

; If ColorThemeIntegrator app is present, uses color settings.
; Assumes that file is in grandparent folder of this file.
; ## NOTE ## that this won't affect the color of the notes. 
; It is for the dialogs. "Note font color" and this font color are different things. 
settingsFile := A_ScriptDir "\..\colorThemeSettings.ini" 
If FileExist(SettingsFile) {  ; Get colors from ini file. 
    fontColor := IniRead(settingsFile, "ColorSettings", "fontColor")
    listColor := IniRead(settingsFile, "ColorSettings", "listColor")
    formColor := IniRead(settingsFile, "ColorSettings", "formColor")
}
Else { ; If color app not found, uses these hex codes.
    fontColor := "0x1F1F1F", listColor := "0xFFFFFF", formColor := "0xE5E4E2"
}

class StickyNotes {
    noteCount := 0
    notes := Map()
    noteManager := ""
    mainWindow := ""
    windowCheckTimer := 0
    hasStuckNotes := false

    static cycleComplete := Map() 
    
    __New() {
        this.InitializeComponents()
        this.SetupSystemTray()
        this.SetupHotkeys()
        this.LoadNotesOnStartup()
        this.CheckMissedAlarms()
        SetTimer(this.CheckAlarms.Bind(this), 1000)
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
            this.noteManager := NoteManager(this.mainWindow)
        } catch as err {
            MsgBox("Error initializing note manager: " err.Message)
            ExitApp()
        }
    }

    SetupHotkeys() {
        ; Bind hotkeys to methods
        HotKey(OptionsConfig.TOGGLE_MAIN_WINDOW, (*) => this.ToggleMainWindow())
        HotKey(OptionsConfig.NEW_NOTE, (*) => this.CreateNewNote())
        HotKey(OptionsConfig.NEW_CLIPBOARD_NOTE, (*) => this.CreateClipboardNote())  ; Add this line
    }
    
    ToggleMainWindow(*) {
        if WinExist("ahk_id " this.mainWindow.gui.Hwnd) {
            this.mainWindow.Hide()
        } else {
            this.mainWindow.Show()
            app.mainWindow.PopulateNoteList()
        }
    }
    
    ProcessClipboardText() {
        ; Check if clipboard has text
        if (!A_Clipboard) {
            MsgBox("No text found on clipboard.", "Sticky Notes", 48)
            return ""
        }

        text := A_Clipboard
        needsTruncation := false
        truncatedText := text

        ; Count words (rough approximation)
        wordCount := 0
        loop parse, text, A_Space "`n`r`t"
            wordCount++

        ; Count lines
        lineCount := 0
        loop parse, text, "`n", "`r"
            lineCount++

        ; Check if text needs truncation
        if (wordCount > OptionsConfig.MAX_NOTE_WORDS || lineCount > OptionsConfig.MAX_NOTE_LINES) {
            needsTruncation := true
            truncatedLines := []
            currentWords := 0
            
            ; Process line by line
            loop parse, text, "`n", "`r" {
                ; Stop if we hit max lines
                if (A_Index > OptionsConfig.MAX_NOTE_LINES) {
                    break
                }

                ; Count words in this line
                lineWords := 0
                loop parse, A_LoopField, A_Space
                    lineWords++

                ; Check if adding this line would exceed word limit
                if (currentWords + lineWords > OptionsConfig.MAX_NOTE_WORDS) {
                    ; Add partial line up to word limit if possible
                    if (currentWords < OptionsConfig.MAX_NOTE_WORDS) {
                        words := []
                        wordCount := 0
                        loop parse, A_LoopField, A_Space {
                            if (currentWords + wordCount < OptionsConfig.MAX_NOTE_WORDS) {
                                words.Push(A_LoopField)
                            } else {
                                break
                            }
                            wordCount++
                        }
                        if (words.Length > 0) {
                            truncatedLines.Push(Join(words, " "))
                        }
                    }
                    break
                }

                truncatedLines.Push(A_LoopField)
                currentWords += lineWords
            }

            truncatedText := Join(truncatedLines, "`n") "`n..."
            MsgBox("Note text has been truncated due to length.`n`nOriginal text had "
                . wordCount . " words and " . lineCount . " lines.", "Sticky Notes", 64)
        }

        return truncatedText
    }

    CreateClipboardNote(*) {
        ; Process clipboard text
        text := this.ProcessClipboardText()
        if (!text) {
            return  ; No valid text to create note with
        }

        ; Get cascade position from note manager
        pos := this.noteManager.GetCascadePosition()

        ; Create the note
        noteId := this.noteManager.CreateNote({
            content: text,
            bgcolor: this.noteManager.GetNextColor(),
            font: StickyNotesConfig.DEFAULT_FONT,
            fontSize: StickyNotesConfig.DEFAULT_FONT_SIZE,
            fontColor: this.noteManager.GetNextFontColor(),
            x: pos.x,  ; Use cascade position
            y: pos.y,
            isOnTop: false,
            isAutoSize: false,
        })

        ; Auto-open editor if enabled
        if (OptionsConfig.AUTO_OPEN_EDITOR && noteId) {
            this.noteManager.notes[noteId].Edit()
        }
    }

    CreateNewNote(*) {
        LogError("Starting CreateNewNote at " A_TickCount "`n")
        noteId := this.noteManager.CreateNote()
        LogError("CreateNewNote completed at " A_TickCount "`n")
        
        ; Auto-open editor if enabled
        if (OptionsConfig.AUTO_OPEN_EDITOR && noteId) {
            this.noteManager.notes[noteId].Edit()
        }
    }
    
    LoadNotesOnStartup() {
        this.noteManager.LoadSavedNotes()
        this.UpdateWindowFollowing()  ; Add this line
    }

    ResetNoteAlarmCycle(noteId) {
        LogError("ResetNoteAlarmCycle called for " noteId "`n")
        
        if (this.cycleComplete.Has(noteId)) {
            this.cycleComplete.Delete(noteId)
            LogError("Cycle reset for " noteId "`n")
        }
    }

    CheckAlarms() {
        ; Keep these maps at class level to preserve state
        if (!this.HasOwnProp("playCount"))
            this.playCount := Map()
        if (!this.HasOwnProp("lastPlayTime"))
            this.lastPlayTime := Map()
        if (!this.HasOwnProp("cycleComplete"))
            this.cycleComplete := Map()

        ResetNoteAlarmCycle(noteId) {
            this.cycleComplete.Delete(noteId)
            this.playCount.Delete(noteId)
            this.lastPlayTime.Delete(noteId)
        }
        
        currentTime := FormatTime(, "h:mm tt")
        LogError("Checking alarms at " currentTime "`n")
        currentDay := FormatTime(, "ddd")
        
        ; Convert current day to short format
        dayMap := Map(
            "Sun", "Su",
            "Mon", "Mo",
            "Tue", "Tu",
            "Wed", "We",
            "Thu", "Th",
            "Fri", "Fr",
            "Sat", "Sa"
        )
        shortDay := dayMap[currentDay]
        
        ; First collect notes with active alarms
        activeAlarms := []
        for id, note in this.noteManager.notes {
            if (!note.hasAlarm)
                continue

            ; Only log when alarm state changes
            static lastAlarmState := Map()
            if (!lastAlarmState.Has(id) || lastAlarmState[id] != note.hasAlarm) {
                LogError("Alarm state changed for note " id ": hasAlarm=" note.hasAlarm "`n")
                lastAlarmState[id] := note.hasAlarm
            }

            ; Reset tracking if alarm time has changed
            if (this.playCount.Has(id)) {
                if (note.alarmTime != FormatTime(this.lastPlayTime[id], "h:mm tt")) {
                    this.playCount.Delete(id)
                    this.lastPlayTime.Delete(id)
                    this.cycleComplete.Delete(id)
                    LogError("Reset tracking for note " id " - time changed`n")
                }
            }
            
            activeAlarms.Push({id: id, note: note})
        }

        ; If no active alarms, we can return early
        if (activeAlarms.Length = 0)
            return

        ; Process only notes with active alarms
        for alarmNote in activeAlarms {
            id := alarmNote.id
            note := alarmNote.note
            
            LogError("Note " id ": Alarm set for " note.alarmTime "`n")
                
            if (currentTime != note.alarmTime)
                continue
                
            ; For notes with weekday settings
            if (note.alarmDays) {
                if (!InStr(note.alarmDays, shortDay))
                    continue
            }
            
            ; Check if we've already handled this alarm
            if (this.cycleComplete.Has(id) && this.cycleComplete[id]) {
                LogError("Skipping note " id " - cycle complete`n")
                continue
            }

            LogError("Checking note " id " - currentTime: " currentTime " alarmTime: " note.alarmTime "`n")
                        
            ; Initialize play count for this note if needed
            if (!this.playCount.Has(id)) {
                this.playCount[id] := 0
                this.lastPlayTime[id] := A_Now
            }
            
            ; Check if enough time has passed since last play
            if (A_Now - this.lastPlayTime[id] >= 0.1) {
                ("Current play count: " (this.playCount.Has(id) ? this.playCount[id] : 0) " of " note.alarmRepeatCount "`n")
                
                ; Trigger the alarm if we haven't reached the repeat count
                if (!this.playCount.Has(id) || this.playCount[id] < note.alarmRepeatCount) {
                    ; Trigger the alarm
                    note.HandleAlarm()
                    
                    ; Initialize or increment play count
                    if (!this.playCount.Has(id)) {
                        this.playCount[id] := 1
                    } else {
                        this.playCount[id]++
                    }
                    
                    this.lastPlayTime[id] := A_Now
                    LogError("Played alarm, count now: " this.playCount[id] "`n")
                }
                
                ; Check if we've finished all repeats
                if (this.playCount.Has(id) && this.playCount[id] >= note.alarmRepeatCount) {
                    LogError("Completed all repeats (" note.alarmRepeatCount ")`n")
                        
                    this.cycleComplete[id] := true
                    
                ; If it's a one-time alarm, remove it
                if (!note.alarmDays) {
                    note.hasAlarm := false
                    note.alarmTime := ""
                    note.alarmSound := ""
                    note.alarmRepeatCount := 1
                    
                    ; Update button text if editor is open
                    if (note.editor) {
                        note.editor.addAlarmBtn.Text := "Add &Alarm..."
                    }
                    
                    storage := NoteStorage()
                    storage.SaveNote(note)
                }
                    
                    ; Clean up tracking maps
                    this.playCount.Delete(id)
                    this.lastPlayTime.Delete(id)
                }
            }
        }
    }

    CheckMissedAlarms() {
        currentTime := FormatTime(A_Now, "h:mm tt")
        currentDay := FormatTime(A_Now, "ddd")
        dayMap := Map(
            "Sun", "Su", "Mon", "Mo", "Tue", "Tu",
            "Wed", "We", "Thu", "Th", "Fri", "Fr", "Sat", "Sa"
        )
        shortDay := dayMap[currentDay]

        LogError("Starting CheckMissedAlarms check at " FormatTime(A_Now, "HH:mm:ss") "`n")

        notesWithAlarms := []
        for id, note in this.noteManager.notes {
            if (note.hasAlarm) {
                notesWithAlarms.Push({id: id, note: note})
                LogError("Found note " id " with alarm set for " note.alarmTime "`n")
            }
        }

        if (notesWithAlarms.Length = 0) {
            LogError("No notes with alarms found`n")
            return
        }

        currentParts := StrSplit(FormatTime(A_Now, "HH:mm"), ":")
        currentHour := Integer(currentParts[1])
        currentMinute := Integer(currentParts[2])

        missedAlarms := []
        for alarmNote in notesWithAlarms {
            note := alarmNote.note
            
            alarmParts := StrSplit(note.alarmTime, " ")
            timeParts := StrSplit(alarmParts[1], ":")
            hour := Integer(timeParts[1])
            minute := Integer(timeParts[2])
            if (alarmParts[2] = "PM" && hour != 12)
                hour += 12
            else if (alarmParts[2] = "AM" && hour = 12)
                hour := 0

            if ((hour < currentHour) || (hour = currentHour && minute < currentMinute)) {
                if (!note.alarmDays || InStr(note.alarmDays, shortDay)) {
                    todayDate := FormatTime(A_Now, "yyyyMMdd")
                    if (note.lastPlayDate != todayDate) {
                        previewText := StrLen(note.content) > 50 ? SubStr(note.content, 1, 47) "..." : note.content
                        missedAlarms.Push({
                            time: note.alarmTime,
                            recurrence: note.alarmDays ? note.alarmDays : "Once",
                            sound: note.alarmSound,
                            preview: previewText
                        })
                    }
                }
            }
        }

        if (missedAlarms.Length > 0) {
            combinedMessage := "Missed Alarms:`n`n"
            for alarm in missedAlarms {
                combinedMessage .= "Time: " alarm.time "`n"
                    . "Recurrence: " alarm.recurrence "`n"
                    . "Sound: " alarm.sound "`n"
                    . "Note preview: " alarm.preview "`n`n"
            }
            MsgBox(combinedMessage, "Missed Alarms")
        }
    }

    UpdateWindowFollowing() {
        ; Check if we have any stuck notes
        hasStuckNotes := false
        for id, note in this.noteManager.notes {
            if (note.isStuckToWindow) {
                hasStuckNotes := true
                break
            }
        }
        
        ; Start or stop timer based on whether we have stuck notes
        if (hasStuckNotes && !this.windowCheckTimer) {
            LogError("Starting window check timer`n")
            this.windowCheckTimer := SetTimer(this.CheckStuckWindows.Bind(this), 500)  ; Changed from 100 to 500ms
        }
        else if (!hasStuckNotes && this.windowCheckTimer) {
            LogError("Stopping window check timer`n")
            SetTimer(this.windowCheckTimer, 0)
            this.windowCheckTimer := 0
        }
    }
        ; Add this method to the StickyNotes class:
    CheckStuckWindows() {
        static lastStates := Map()  ; Track last known state for each note
        
        for id, note in this.noteManager.notes {
            if (!note.isStuckToWindow)
                continue
                    
            LogError("Checking stuck note " id " for window " note.stuckWindowTitle "`n")
            
            ; Simple window existence check
            targetWindow := WinExist("ahk_class " note.stuckWindowClass)
            windowExists := (targetWindow && InStr(WinGetTitle("ahk_id " targetWindow), note.stuckWindowTitle))
            
            ; Get current note visibility state, checking borderGui if it exists
            try {
                isCurrentlyVisible := false
                if (note.borderGui && note.borderGui.Gui1) {
                    isCurrentlyVisible := WinExist("ahk_id " note.borderGui.Gui1.Hwnd) ? true : false
                } else if (note.gui) {
                    isCurrentlyVisible := WinExist("ahk_id " note.gui.Hwnd) ? true : false
                }
                LogError("Note " id " visibility check - isCurrentlyVisible: " isCurrentlyVisible "`n")
            } catch Error as err {
                LogError("Error checking note " id " visibility: " err.Message "`n")
                isCurrentlyVisible := false
            }
            
            ; Only act if state has changed
            if (!lastStates.Has(id) || lastStates[id] != windowExists) {
                storage := NoteStorage()
                isHidden := storage.IsStuckNoteHidden(id)
                
                if (windowExists && !isHidden && !isCurrentlyVisible) {
                    LogError("Showing stuck note " id " with border: " (note.borderGui ? "yes" : "no") "`n")
                    note.Show()
                } else if (!windowExists && isCurrentlyVisible) {
                    LogError("Hiding stuck note " id " with border: " (note.borderGui ? "yes" : "no") "`n")
                    note.Hide()
                }
                
                lastStates[id] := windowExists
            }
        }
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
        ; Try to use custom icon first, fall back to system icon if not found
        if FileExist(OptionsConfig.APP_ICON)
            TraySetIcon(OptionsConfig.APP_ICON)
        else
            TraySetIcon("imageres.dll",279)  ; Yellow sticky with green 'up' arrow
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

; Helper function to join array elements
Join(arr, delimiter) {
    result := ""
    for item in arr {
        result .= (A_Index = 1 ? "" : delimiter) item
    }
    return result
}

; Create and start the application
global app := StickyNotes()

/*
BorderGui class for note borders
Graciously made by HellBent
https://www.autohotkey.com/boards/viewtopic.php?p=597649&sid=4af1ffa41437ec2526d652565324d494#p597649
Minor changes made from original.
*/
class BorderGui {
    static init := This._Setup()
    static Handles := Map()
    static _Setup() {
        OnMessage(0x0201, This._ClickEvent.Bind(This))
    }
    static _ClickEvent(wParam, lParam, uMsg, hWnd) {
        if (This.Handles.HasProp(hWNd)) {
            PostMessage(0xA1, 2)
            KeyWait("LButton")
            return 0
        }
    }

    __New(GuiObject, borderColor := "0x000000", top := 20, bottom := 5, left := 5, right := 5, activate := 1, isOnTop := false) {
        if (!IsObject(GuiObject)) {
            MsgBox("The BorderGui class requires you to pass a gui object.", "Error", 8192)
            return 0
        }
        if (GuiObject.HasProp("HasParent") && GuiObject.HasParent) {
            MsgBox("The target gui already has a parent window", "Error", 8192)
            return 0
        }

        This.Child := GuiObject
        ; Create border GUI with conditional AlwaysOnTop
        This.Gui1 := Gui("-Caption +ToolWindow" (isOnTop ? " +AlwaysOnTop" : ""), This.Child.Title)
        This.Gui1.BackColor := borderColor
        
        ; Get child window dimensions
        w := 0, h := 0, x := 0, y := 0
        This.Child.GetPos(&x, &y, &w, &h)
        
        ; Hide child before reparenting
        This.Child.Hide()
        
        ; Set up parent/child relationship
        This.Child.Opt("+Parent" This.Gui1.Hwnd)
        
        ; Show child with offset for border
        This.Child.Show("x" left " y" top " w" w " h" h ((activate) ? ("") : (" NA")))
        This.Child.HasParent := 1
        
        ; Calculate border window position and size
        posW := w + left + right, posH := h + top + bottom
        posX := x - left, posY := y - top
        
        ; Show border window
        This.Gui1.Show("x" posX " y" posY " w" posW " h" posH)
        
        ; Store handle in static map
        handle := This.Gui1.Hwnd
        BorderGui.Handles.%handle% := This
    }

    SetAlwaysOnTop(state) {
        if (this.Gui1) {
            this.Gui1.Opt(state ? "+AlwaysOnTop" : "-AlwaysOnTop")
        }
    }

    Release() {
        if (!this.Child || !this.Gui1)
            return
            
        try {
            ; Get final positions
            this.Child.GetPos(&xOffset, &yOffset)
            this.Gui1.GetPos(&parentX, &parentY)
            
            ; Calculate absolute position for child
            absoluteX := parentX + xOffset
            absoluteY := parentY + yOffset
            
            ; Remove parent relationship
            this.Child.Opt("-Parent")
            this.Child.HasParent := 0
            
            ; Remove from static handles and destroy
            BorderGui.Handles.DeleteProp(this.Gui1.Hwnd)
            
            ; Important: Hide and destroy BEFORE showing child
            WinSetExStyle(-0x80, this.Gui1)  ; Remove WS_EX_TOOLWINDOW
            this.Gui1.Destroy()
            this.Gui1 := ""
            
            ; Now show child at correct position
            this.Child.Show(Format("x{} y{}", absoluteX, absoluteY))
            this.Child := ""
        } catch as err {
            LogError("Error in BorderGui.Release: " err.Message "`n")
        }
    }

    __Delete() {
        this.Release()
    }
}

class Note {
    ; Note properties
    gui := ""           
    dragArea := ""    
    borderGui := ""  
    id := 0
    hasBorder := false
    content := ""
    bgcolor := ""
    font := ""
    fontSize := 0
    fontColor := ""
    isBold := false
    isOnTop := false
    width := StickyNotesConfig.DEFAULT_WIDTH
    editor := ""
    controls := []
    currentX := ""
    currentY := ""

    ; alarm-related properties
    hasAlarm := false
    alarmTime := ""
    alarmSound := ""
    alarmDays := ""
    alarmRepeatCount := 1
    visualShake := true  
    lastPlayDate := ""

    ; window-stick properties
    isStuckToWindow := false
    stuckWindowTitle := ""
    stuckWindowClass := ""

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

        ; Initialize alarm properties
        this.hasAlarm := options.HasOwnProp("hasAlarm") ? options.hasAlarm : false
        this.alarmTime := options.HasOwnProp("alarmTime") ? options.alarmTime : ""
        this.alarmSound := options.HasOwnProp("alarmSound") ? options.alarmSound : ""
        this.alarmDays := options.HasOwnProp("alarmDays") ? options.alarmDays : ""
        this.alarmRepeatCount := options.HasOwnProp("alarmRepeatCount") ? options.alarmRepeatCount : 1
        this.lastPlayDate := options.HasOwnProp("lastPlayDate") ? options.lastPlayDate : ""

        ; Initialize window-sticking properties
        this.isStuckToWindow := options.HasOwnProp("isStuckToWindow") ? !!options.isStuckToWindow : false
        this.stuckWindowTitle := options.HasOwnProp("stuckWindowTitle") ? options.stuckWindowTitle : ""
        this.stuckWindowClass := options.HasOwnProp("stuckWindowClass") ? options.stuckWindowClass : ""
        
        LogError("Initializing note " idNum " with window properties:`n")
        LogError("  isStuckToWindow: " this.isStuckToWindow "`n")
        LogError("  stuckWindowTitle: " this.stuckWindowTitle "`n")
        LogError("  stuckWindowClass: " this.stuckWindowClass "`n")
        ; Create the basic GUI
        if (this.isOnTop) {
            this.gui := Gui("-Caption +AlwaysOnTop +Owner")
        } else {
            this.gui := Gui("-Caption +Owner")
        }
        
        this.gui.BackColor := this.bgcolor
        
        ; Create drag area
        xPos := OptionsConfig.RESERVE_DRAG_AREA ? 0 : OptionsConfig.NO_RESERVE_X_OFFSET
        this.dragArea := this.gui.Add("Text", 
            "x" xPos " y0 w" (this.width - xPos) " h20 Background" this.bgcolor)

            
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

        ; Initialize border if needed
        this.hasBorder := options.HasOwnProp("hasBorder") ? options.hasBorder : OptionsConfig.DEFAULT_BORDER
        if (this.hasBorder) {
            thickness := this.isBold ? 4 : 2
            this.borderGui := BorderGui(this.gui, this.fontColor, thickness, thickness, thickness, thickness, 1)
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
        yPos := OptionsConfig.RESERVE_DRAG_AREA ? OptionsConfig.DRAG_AREA_OFFSET : OptionsConfig.NO_RESERVE_OFFSET
        txt := this.gui.Add("Text", 
            "y" yPos " x5 w" this.width,
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
        currentY := OptionsConfig.RESERVE_DRAG_AREA ? OptionsConfig.DRAG_AREA_OFFSET : OptionsConfig.NO_RESERVE_OFFSET
        
        ; Parse content for checkboxes
        parsed := this.ParseCheckboxContent()
        
        ; Create controls
        for item in parsed {
            if (item.type == "checkbox") {
                ; Create checkbox with proper font color
                cb := this.gui.Add("Checkbox",
                    "y" currentY " x5 w" this.width " -wrap c" this.fontColor,  ; Add color directly in options
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
        
        ; double-click handlers both using "DblClick"
        this.dragArea.OnEvent("DoubleClick", this.DoubleClickHandler.Bind(this))

        ; WM_ACTIVATE handler for saving position
        OnMessage(0x0006, this.WM_ACTIVATE.Bind(this))
    }

    DoubleClickHandler(*) {
        LogError("Double-click detected`n")
        this.Edit()
    }

    ; for when a note gui loses focus.
    isBeingDeleted := false
    WM_ACTIVATE(wParam, lParam, msg, hwnd) {
        try {
            ; First verify we have a valid GUI
            if (!this.gui || !IsObject(this.gui)) {
                return
            }
            
            ; Then check if it's our window losing focus
            if (hwnd = this.gui.Hwnd && wParam = 0) {
                ; Get the correct position to save
                x := 0
                y := 0
                if (this.borderGui && this.borderGui.Gui1) {
                    ; If bordered, save the border GUI's position
                    this.borderGui.Gui1.GetPos(&x, &y)
                } else {
                    ; If not bordered, save the note's position
                    this.gui.GetPos(&x, &y)
                }
                
                ; Only save if the note still exists in storage
                storage := NoteStorage()
                if (storage.NoteExists(this.id)) {
                    this.currentX := x
                    this.currentY := y
                    storage.SaveNote(this)
                    LogError("Saved position for note " this.id " after losing focus: x=" x ", y=" y "`n")
                } else {
                    LogError("Note " this.id " no longer exists - skipping position save`n")
                }
            }
        } catch as err {
            LogError("Error in WM_ACTIVATE: " err.Message "`n")
        }
    }

    StartDrag(*) {
        SetWinDelay(-1)
        CoordMode("Mouse", "Screen")
        
        startWinX := 0, startWinY := 0
        startMouseX := 0, startMouseY := 0
        currentMouseX := 0, currentMouseY := 0
        
        ; Determine which window to move based on whether we have a border
        targetGui := this.borderGui ? this.borderGui.Gui1 : this.gui
        
        WinGetPos(&startWinX, &startWinY, , , targetGui)
        MouseGetPos(&startMouseX, &startMouseY)
        
        while GetKeyState("LButton", "P") {
            MouseGetPos(&currentMouseX, &currentMouseY)
            WinMove(
                startWinX + (currentMouseX - startMouseX),
                startWinY + (currentMouseY - startMouseY),
                , , targetGui
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
        ; Add alarm-related items if note has an alarm
        if (this.hasAlarm) {
            noteMenu.Add("Stop Alarm", this.StopSound.Bind(this))
            noteMenu.Add("Delete Alarm", this.DeleteAlarm.Bind(this))
            noteMenu.Add()
        }
        noteMenu.Add("Show Main Window", (*) => app.mainWindow.Show())
        noteMenu.Add("New Sticky Note", (*) => app.noteManager.CreateNote())
        noteMenu.Add()
        noteMenu.Add("Save Status", (*) => app.noteManager.SaveAllNotes())
        noteMenu.Add("Show Hidden", (*) => app.noteManager.ShowHiddenNotes())
        noteMenu.Add("ReLoad Notes", (*) => app.noteManager.LoadSavedNotes())

        noteMenu.Show()
    }
    
        StopSound(*) {
        try SoundPlay("nonexistent.wav")
    }

    SaveCheckboxState(ctrl, *) {
        ; Require modifier key to toggle checkbox?
        If !(OptionsConfig.CHECKBOX_MODIFIER_KEY = "")
            if !GetKeyState(OptionsConfig.CHECKBOX_MODIFIER_KEY) {
                ctrl.Value := !ctrl.Value
                Return
            }
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

    HandleAlarm() {
        if (!this.hasAlarm)
            return
                
        ; Show the note if it's hidden
        this.Show()
        
        ; Trigger shake effect if enabled
        if (this.visualShake)
            this.ShakeNote()
        
        ; Play the alarm sound if specified
        if (this.alarmSound && FileExist(this.alarmSound))
            SoundPlay(this.alarmSound, 1)

        ; Update lastPlayDate
        this.lastPlayDate := FormatTime(A_Now, "yyyyMMdd")
        ; Save to storage immediately
        storage := NoteStorage()
        storage.SaveNote(this)
                
        LogError("Playing alarm for note " this.id "`n")
    }

    DeleteAlarm(*) {
        if (MsgBox("Are you sure you want to delete this alarm?",, "YesNo") = "Yes") {
            try {
                this.hasAlarm := false
                this.alarmTime := ""
                this.alarmSound := ""
                this.alarmDays := ""
                this.alarmRepeatCount := 1
                
                ; Force reset of cycle state - only if key exists
                if (StickyNotes.cycleComplete && StickyNotes.cycleComplete.Has(this.id)) {
                    StickyNotes.cycleComplete.Delete(this.id)
                }
                
                ; Save changes to storage
                storage := NoteStorage()
                storage.SaveNote(this)
                
                ; If note editor is open, update the alarm button text
                if (this.editor) {
                    this.editor.addAlarmBtn.Text := "Add &Alarm..."
                }
            } catch Error as err {
                LogError("Error in Note.DeleteAlarm: " err.Message "`n")
            }
        }
    }

    ShakeNote() {
        this.isShaking := false  ; Class property to track shake state
        
        ; Create a bound function for the timer to use
        boundShake := this.PerformShake.Bind(this)
        SetTimer(boundShake, 10)  ; Start immediately
    }

    ; handle the actual shake animation:
    PerformShake() {
        static shakeStep := 0
        static originalX := 0
        static originalY := 0
        
        try {
            targetGui := this.borderGui ? this.borderGui.Gui1 : this.gui
            
            ; Initialize on first step
            if (!this.isShaking) {
                this.isShaking := true
                targetGui.GetPos(&originalX, &originalY)
                shakeStep := 0
                LogError("Starting shake animation for note " this.id "`n")
            }
            
            ; Calculate diminishing shake distance
            progress := shakeStep / OptionsConfig.SHAKE_STEPS
            range := OptionsConfig.MAX_SHAKE_DISTANCE - OptionsConfig.MIN_SHAKE_DISTANCE
            shakeDistance := Max(
                OptionsConfig.MIN_SHAKE_DISTANCE,
                OptionsConfig.MAX_SHAKE_DISTANCE - (range * progress)
            )
            
            ; Calculate new position based on step
            Switch Mod(shakeStep, 4) {
                Case 0: WinMove(originalX + shakeDistance, originalY,,, targetGui)
                Case 1: WinMove(originalX - shakeDistance, originalY,,, targetGui)
                Case 2: WinMove(originalX, originalY - shakeDistance,,, targetGui)
                Case 3: WinMove(originalX, originalY + shakeDistance,,, targetGui)
            }
            
            shakeStep++
            
            ; Stop after configured number of steps
            if (shakeStep >= OptionsConfig.SHAKE_STEPS) {
                ; Return to original position
                WinMove(originalX, originalY,,, targetGui)
                this.isShaking := false
                shakeStep := 0
                SetTimer(, 0)  ; Stop timer
                LogError("Completed shake animation for note " this.id "`n")
                return
            }
            
        } catch as err {
            LogError("Error in PerformShake: " err.Message "`n")
            SetTimer(, 0)  ; Stop timer on error
            return
        }
    }

    UpdateContent(newContent, options := "") {
        ; Store current position and border state
        x := 0
        y := 0
        if (this.borderGui && this.borderGui.Gui1) {
            this.borderGui.Gui1.GetPos(&x, &y)
        } else {
            this.gui.GetPos(&x, &y)
        }
        
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
            if (options.HasOwnProp("hasBorder"))
                this.hasBorder := options.hasBorder
        }
        
        ; Destroy old GUI components
        if (this.borderGui) {
            this.borderGui.Release()
            this.borderGui := ""
        }
        this.gui.Destroy()
        
        ; Create new GUI with current AlwaysOnTop setting
        if (this.isOnTop) {
            this.gui := Gui("-Caption +AlwaysOnTop")
        } else {
            this.gui := Gui("-Caption")
        }
        
        this.gui.BackColor := this.bgcolor
        
        ; Create drag area
        xPos := OptionsConfig.RESERVE_DRAG_AREA ? 0 : OptionsConfig.NO_RESERVE_X_OFFSET
        this.dragArea := this.gui.Add("Text", 
            "x" xPos " y0 w" (this.width - xPos) " h20 Background" this.bgcolor)

            
        ; Set default font
        this.gui.SetFont("s" this.fontSize (this.isBold ? " bold" : ""), this.font)
        
        ; Create note content
        if (InStr(this.content, "[]") || InStr(this.content, "[x]")) {
            this.CreateComplexNote()
        } else {
            this.CreateSimpleNote()
        }
        
        ; First show the note GUI to establish its size
        this.gui.Show(Format("x{} y{}", x, y))
        
        ; Then handle border if needed
        if (this.hasBorder) {
            thickness := this.isBold ? 4 : 2
            this.borderGui := BorderGui(this.gui, this.fontColor, thickness, thickness, thickness, thickness, 1, this.isOnTop)
        }
        
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
                LogError("Position out of bounds - x: " x ", y: " y " (screen: " left "," top "," right "," bottom ")`n")
                return false
                }
            }
        return true
    }

    Hide(*) {
        if (this.borderGui && this.borderGui.Gui1) {
            ; If note has a border, hide the border GUI (child will hide automatically)
            this.borderGui.Gui1.Hide()
        } else {
            ; If no border, just hide the note
            this.gui.Hide()
        }
        ; Only mark as hidden in storage if this is a user-initiated hide,
        ; not an automatic hide due to window closing
        if (!this.isStuckToWindow) {
            storage := NoteStorage()
            storage.MarkNoteHidden(this.id)
        }
    }

    Show(*) {
        if (this.borderGui && this.borderGui.Gui1) {
            ; For bordered notes, need to show both explicitly
            this.gui.Show()  ; Show child first
            this.borderGui.Gui1.Show()  ; Then show parent
        } else {
            ; If no border, just show the note
            this.gui.Show()
        }
    }
        
    Delete(*) {
        if (!OptionsConfig.WARN_ON_DELETE_NOTE || 
            MsgBox("Are you sure you want to delete this note?",, "YesNo") = "Yes") {
            try {
                ; Clean up alarm state if it exists
                if (this.hasAlarm && StickyNotes.cycleComplete && StickyNotes.cycleComplete.Has(this.id)) {
                    StickyNotes.cycleComplete.Delete(this.id)
                }

                ; Clean up border if it exists
                if (this.borderGui) {
                    this.borderGui.Release()
                    this.borderGui := ""
                }

                ; Use the app's noteManager to delete this note
                app.noteManager.DeleteNote(this.id)
            } catch Error as err {
                LogError("Error in Note.Delete: " err.Message "`n")
            }
        }
    }
    
    Destroy(*) {
        try {
            ; First clean up border if it exists
            if (this.borderGui) {
                this.borderGui.Release()
                this.borderGui := ""
            }

            ; Then clean up event bindings
            if (this.dragArea) {
                this.dragArea.OnEvent("Click", this.StartDrag.Bind(this), 0)
                this.dragArea.OnEvent("DoubleClick", this.DoubleClickHandler.Bind(this), 0)
                this.dragArea := ""
            }

            ; Clean up controls array
            if (this.controls) {
                for item in this.controls {
                    if (item.HasOwnProp("control") && item.control) {
                        if (item.type == "checkbox") {
                            item.control.OnEvent("Click", this.SaveCheckboxState.Bind(this), 0)
                        }
                        item.control := ""
                    }
                }
                this.controls := []
            }

            ; Clean up editor if it exists
            if (this.editor) {
                this.editor.Destroy()
                this.editor := ""
            }

            ; Finally destroy the GUI
            if (this.gui) {
                this.gui.OnEvent("ContextMenu", this.ShowContextMenu.Bind(this), 0)
                this.gui.Destroy()
                this.gui := ""
            }

            return true
        } catch as err {
            LogError("Error in Note.Destroy: " err.Message)
            return false
        }
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

class AlarmConfig {
    
    static ALARM_SOUNDS_FOLDER := "Sounds"  ; Folder containing .wav files
    
    ; Define days in chronological order
    static WEEKDAY_ORDER := ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]
    
    ; Helper method to sort days chronologically
    static SortDays(daysStr) {
        if (!daysStr)
            return ""
            
        ; Build string in chronological order
        orderedDays := ""
        for day in AlarmConfig.WEEKDAY_ORDER {
            if InStr(daysStr, day)
                orderedDays .= day
        }
        return orderedDays
    }

    static GetSoundFiles() {
        soundFiles := []
        if (!DirExist(AlarmConfig.ALARM_SOUNDS_FOLDER)) {
            soundFiles.Push("Default Beep")  ; Add at least one option
            return soundFiles
        }
        
        Loop Files, AlarmConfig.ALARM_SOUNDS_FOLDER "\*.wav" {
            soundFiles.Push(A_LoopFileName)
        }
        
        if (soundFiles.Length = 0) {  ; If no .wav files found
            soundFiles.Push("Default Beep")
        }
        
        return soundFiles
    }
}

class NoteManager {
    notes := Map()
    storage := NoteStorage()
    mainWindow := ""

    ; Static properties for cascade positioning
    static CASCADE_OFFSET := 75      ; Pixels to offset each new note
    static MAX_CASCADE := 10         ; Maximum number of cascaded notes before reset
    static currentCascadeCount := 0  ; Track number of cascaded notes
    static baseX := ""              ; Base X position (will be set on first use)
    static baseY := ""              ; Base Y position (will be set on first use)
    static currentColorIndex := NoteManager.InitRandomColorIndex()
    static currentFontColorIndex := NoteManager.InitRandomFontColorIndex()

    ; Helper methods for initialization
    static InitRandomColorIndex() {
        colorCount := 0
        for _ in StickyNotesConfig.COLORS
            colorCount++
        Random start := 1, colorCount
        LogError("Starting with color index: " (start - 1) "`n")
        return start - 1  ; -1 because GetNextColor increments first
    }

    static InitRandomFontColorIndex() {
        fontColorCount := 0
        for _ in StickyNotesConfig.FONT_COLORS
            fontColorCount++
        Random start := 1, fontColorCount
        LogError("Starting with font color index: " (start - 1) "`n")
        return start - 1
    }

    __New(mainWindow := "") {
        this.mainWindow := mainWindow
    }
    
    GetNextColor() {
        ; Convert color map to array for easy indexing
        colorCodes := []
        for name, code in StickyNotesConfig.COLORS
            colorCodes.Push(code)
            
        ; Increment and wrap the index
        NoteManager.currentColorIndex := Mod(NoteManager.currentColorIndex + 1, colorCodes.Length)
        
        ; Return the next color
        return colorCodes[NoteManager.currentColorIndex + 1]  ; +1 because array is 1-based
    }

    GetNextFontColor() {
            ; If cycling is disabled, return black
            if (!OptionsConfig.CYCLE_FONT_COLOR)
                return "000000"  ; Black
                
            ; Convert font color map to array for easy indexing
            fontColorCodes := []
            for name, code in StickyNotesConfig.FONT_COLORS
                fontColorCodes.Push(code)
                
            ; Increment and wrap the index
            NoteManager.currentFontColorIndex := Mod(NoteManager.currentFontColorIndex + 1, fontColorCodes.Length)
            
            ; Return the next color
            return fontColorCodes[NoteManager.currentFontColorIndex + 1]  ; +1 because array is 1-based
        }

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
                bgcolor: this.GetNextColor(),
                font: StickyNotesConfig.DEFAULT_FONT,
                fontSize: StickyNotesConfig.DEFAULT_FONT_SIZE,
                fontColor: this.GetNextFontColor(),  ; Use cycling font colors
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
            
            ; Only save initial state if this isn't a new blank note
            if (options.HasOwnProp("content") && options.content != "Right-click here, or double-click header, to edit. Drag header to move.") {
                this.storage.SaveNote(newNote)
            }
            
            if (this.mainWindow)
                this.mainWindow.PopulateNoteList()
            return newId
        } catch as err {
            MsgBox("Error creating note: " err.Message "`n" err.Stack)
            return 0
        }
    }

    SaveAllNotes() {
        try {
            LogError("`nStarting SaveAllNotes...`n")
            LogError("Current number of notes: " this.notes.Count "`n")

            for id, note in this.notes {
                try {
                    ; Get and log current position
                    x := 0
                    y := 0
                    note.gui.GetPos(&x, &y)
                    LogError("Saving note " id " at position x=" x ", y=" y "`n")
                    
                    ; Update position properties
                    note.currentX := x
                    note.currentY := y
                    
                    ; Save to storage
                    this.storage.SaveNote(note)
                    
                } catch as err {
                    LogError("Error saving note " id ": " err.Message "`n")
                }
            }
            
            LogError("SaveAllNotes complete`n")
            return true
        } catch as err {
                LogError("Error in SaveAllNotes: " err.Message "`n")
            return false
        }
    }

    LoadSavedNotes() {
        try {
            LogError("Starting LoadSavedNotes...`n")
            
            ; First destroy all existing note GUIs and editors
            LogError("Destroying existing notes...`n")
            for id, existingNote in this.notes {
                LogError("Attempting to destroy note " id "`n")
                if (!existingNote.Destroy()) {
                    LogError("Failed to destroy note " id "`n")
                }
            }
            
            ; Reset note tracking completely
            LogError("Resetting note tracking...`n")
            this.notes := Map()
            
            ; Get all saved notes from storage
            LogError("Loading notes from storage...`n")
            savedNotes := this.storage.LoadAllNotes()
            LogError("Found " savedNotes.Length " notes in storage`n")
            
            ; Process each saved note
            for noteData in savedNotes {
                try {
                    LogError("Processing note " noteData.id "`n")
                    ; Verify note exists in storage before creating
                    if (!this.storage.NoteExists(noteData.id)) {
                        LogError("Note " noteData.id " doesn't exist in storage - skipping`n")
                        continue
                    }
                    
                    ; Create GUI if note is visible OR has an alarm
                    if (!noteData.isHidden || (noteData.hasAlarm && noteData.alarmTime)) {
                        LogError("Creating GUI for note " noteData.id " (Visible: " (!noteData.isHidden ? "Yes" : "No") ", Alarmed: " (noteData.hasAlarm ? "Yes" : "No") ")`n")
                        newNote := Note(noteData.id, noteData)
                        this.notes[noteData.id] := newNote
                        
                        ; If note should be hidden but has an alarm, hide it after creation
                        if (noteData.isHidden) {
                            LogError("Hiding alarmed note " noteData.id "`n")
                            newNote.gui.Hide()
                            if (newNote.borderGui && newNote.borderGui.Gui1)
                                newNote.borderGui.Gui1.Hide()
                        }
                    }
                } catch as err {
                    LogError("Error creating note " noteData.id ": " err.Message "`n")
                    continue
                }
            }
            
            LogError("LoadSavedNotes complete`n")
            return true
        } catch as err {
            LogError("Error in LoadSavedNotes: " err.Message)
            return false
        }
    }

    DeleteNote(id) {
        LogError("DeleteNote called for id " id "`n")
        LogError("Current notes in manager: " this.notes.Count "`n")
        if !this.notes.Has(id) {
            LogError("Note " id " not found in manager`n")
            return false
        }
            
        try {
            ; Get note reference
            note := this.notes[id]
            
            ; Remove from storage first
            this.storage.DeleteNote(id)
            
            ; Destroy note and all its resources
            note.Destroy()
            
            ; Remove from collection and clear reference
            this.notes.Delete(id)
            note := ""
            
            ; Update main window list if it exists
            if (this.mainWindow)
                this.mainWindow.PopulateNoteList()
                
            LogError("Successfully deleted note " id "`n")
            return true
        } catch as err {
            LogError("Error in DeleteNote for id " id ": " err.Message "`n")
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
            IniWrite(note.hasBorder, OptionsConfig.INI_FILE, sectionName, "HasBorder")

            IniWrite(note.hasAlarm, OptionsConfig.INI_FILE, sectionName, "HasAlarm")
            if (note.hasAlarm) {
                IniWrite(note.alarmTime, OptionsConfig.INI_FILE, sectionName, "AlarmTime")
                IniWrite(note.alarmSound, OptionsConfig.INI_FILE, sectionName, "AlarmSound")
                IniWrite(note.alarmDays, OptionsConfig.INI_FILE, sectionName, "AlarmDays")
                IniWrite(note.alarmRepeatCount, OptionsConfig.INI_FILE, sectionName, "AlarmRepeatCount")
                IniWrite(note.visualShake, OptionsConfig.INI_FILE, sectionName, "VisualShake")
                IniWrite(note.lastPlayDate, OptionsConfig.INI_FILE, sectionName, "LastPlayDate")
            } 

            ; Save window sticking info
            IniWrite(note.isStuckToWindow ? 1 : 0, OptionsConfig.INI_FILE, sectionName, "IsStuckToWindow")
            if (note.isStuckToWindow) {
                IniWrite(note.stuckWindowTitle, OptionsConfig.INI_FILE, sectionName, "StuckWindowTitle")
                IniWrite(note.stuckWindowClass, OptionsConfig.INI_FILE, sectionName, "StuckWindowClass")
            }
                
            return true
        } catch as err {  ; <- This was also missing its closing brace
            MsgBox("Error saving note to INI: " err.Message)
            return false
        }
    }
        
    LoadNote(id) {
        try {
            sectionName := this.GetNoteSectionName(id)
            LogError("LoadNote: Starting to load note " id "`n")
            
            ; Check if note exists
            if !this.NoteExists(id) {
                LogError("LoadNote: Note " id " doesn't exist`n")
                return false
            }
            
            ; Load raw content from INI
            rawContent := IniRead(OptionsConfig.INI_FILE, sectionName, "Content", "")
            LogError("LoadNote: Raw content length: " StrLen(rawContent) "`n")
            
            ; Convert literal '\n' back to actual newlines
            unescapedContent := StrReplace(rawContent, "\n", "`n")
            LogError("LoadNote: Unescaped content length: " StrLen(unescapedContent) "`n")
            
            ; Load raw content from INI
            rawContent := IniRead(OptionsConfig.INI_FILE, sectionName, "Content", "")
            
            ; Convert literal '\n' back to actual newlines
            unescapedContent := StrReplace(rawContent, "\n", "`n")
            
            ; Load all note data
            noteData := {
                id: id,
                content: unescapedContent,
                bgcolor: IniRead(OptionsConfig.INI_FILE, sectionName, "Color", StickyNotesConfig.DEFAULT_BG_COLOR),
                font: IniRead(OptionsConfig.INI_FILE, sectionName, "Font", StickyNotesConfig.DEFAULT_FONT),
                fontSize: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "FontSize", StickyNotesConfig.DEFAULT_FONT_SIZE)),
                isBold: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "IsBold", "0")),
                fontColor: IniRead(OptionsConfig.INI_FILE, sectionName, "FontColor", "000000"),
                x: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "PosX", "")),
                y: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "PosY", "")),
                isOnTop: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "IsOnTop", "0")),
                width: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "Width", StickyNotesConfig.DEFAULT_WIDTH)),
                isHidden: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "Hidden", "0")),
                hasBorder: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "HasBorder", OptionsConfig.DEFAULT_BORDER)),
                hasAlarm: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "HasAlarm", "0")),
                alarmTime: IniRead(OptionsConfig.INI_FILE, sectionName, "AlarmTime", ""),
                alarmSound: IniRead(OptionsConfig.INI_FILE, sectionName, "AlarmSound", ""),
                alarmDays: IniRead(OptionsConfig.INI_FILE, sectionName, "AlarmDays", ""),
                alarmRepeatCount: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "AlarmRepeatCount", "1")),
                visualShake: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "VisualShake", "1")),
                lastPlayDate: IniRead(OptionsConfig.INI_FILE, sectionName, "LastPlayDate", ""),
                isStuckToWindow: Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "IsStuckToWindow", "0")) = 1,
                stuckWindowTitle: IniRead(OptionsConfig.INI_FILE, sectionName, "StuckWindowTitle", ""),
                stuckWindowClass: IniRead(OptionsConfig.INI_FILE, sectionName, "StuckWindowClass", "")
            }
            
            LogError("Loading note " id " with window properties:`n")
            LogError("  isStuckToWindow: " noteData.isStuckToWindow "`n")
            LogError("  stuckWindowTitle: " noteData.stuckWindowTitle "`n")
            LogError("  stuckWindowClass: " noteData.stuckWindowClass "`n")

            return noteData
        } catch as err {
            LogError("LoadNote: Error loading note " id ": " err.Message "`n")
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
            
            LogError("LoadAllNotes - Before loading sections, verifying each section exists...`n")
            validSections := ""
            Loop Parse, sections, "`n", "`r" {
                if RegExMatch(A_LoopField, "Note-(\d{14})", &match) {
                    noteId := match[1]
                    ; Explicitly check if the section still exists
                    if (IniRead(OptionsConfig.INI_FILE, A_LoopField, "Content", "") != "") {
                        validSections .= A_LoopField "`n"
                        LogError("Found valid section: " A_LoopField "`n")
                    } else {
                        LogError("Found invalid/deleted section: " A_LoopField "`n")
                    }
                }
            }
            sections := validSections
            
            LogError("Reading sections from INI file...`n")
            loop parse, sections, "`n", "`r" {
                ; Update regex to match new timestamp format
                if RegExMatch(A_LoopField, "Note-(\d{14})", &match) {
                    noteId := match[1]
                    LogError("Found section for note " noteId "`n")
                    
                    ; Verify section still exists before loading
                    if (!this.NoteExists(noteId)) {
                        LogError("Section appears deleted for note " noteId " - skipping`n")
                        continue
                    }
                    
                    if noteData := this.LoadNote(noteId) {
                        LogError("Successfully loaded note " noteId "`n")
                        notes.Push(noteData)
                    } else {
                        LogError("Failed to load note " noteId "`n")
                    }
                }
            }
            
            LogError("LoadAllNotes complete - found " notes.Length " valid notes`n")
            NoteStorage.isLoading := false
            return notes
            
        } catch as err {
            LogError("Error in LoadAllNotes: " err.Message "`n")
            NoteStorage.isLoading := false
            return []
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
    
    IsStuckNoteHidden(id) {
        try {
            sectionName := this.GetNoteSectionName(id)
            return Integer(IniRead(OptionsConfig.INI_FILE, sectionName, "Hidden", "0")) = 1
        } catch {
            return false
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
    
    DeleteNote(id) {
        try {
            sectionName := this.GetNoteSectionName(id)
            LogError("Starting deletion of section " sectionName " from INI file`n")
            
            ; Get all sections
            sections := IniRead(OptionsConfig.INI_FILE)
            
            ; Delete the section
            IniDelete(OptionsConfig.INI_FILE, sectionName)
            LogError("Attempted deletion of section " sectionName "`n")
            
            ; Verify deletion
            if (this.NoteExists(id)) {
                LogError("WARNING: First deletion attempt failed - trying again`n")
                ; Try more aggressive deletion
                Loop 3 {
                    Sleep(50)  ; Give filesystem time
                    IniDelete(OptionsConfig.INI_FILE, sectionName)
                    if (!this.NoteExists(id)) {
                        LogError("Deletion successful on attempt " A_Index "`n")
                        break
                    }
                }
            }
            
            ; Final verification
            if (this.NoteExists(id)) {
                LogError("ERROR: Failed to delete section after multiple attempts`n")
                return false
            }
            
            LogError("Successfully deleted section " sectionName "`n")
            return true
        } catch Error as err {
            LogError("Error in NoteStorage.DeleteNote for id " id ": " err.Message "`n")
            return false
        }
    }
}

class NoteEditor {
    ; Editor properties
    note := ""
    gui := ""
    editControl := ""
    alarmDialog := "" 
    editorWidth := 0 
    addAlarmBtn := ""
    widthEdit := ""
    
    __New(note) {
        this.note := note
        this.CreateGui()
    }
    
    CreateGui() {
        editorWidth := Max(this.note.width, 200)  ; Use note width but minimum of 200
        
        ; Create editor window - don't set window background color
        this.gui := Gui("+AlwaysOnTop", "Edit Note " this.note.id)
        this.gui.BackColor := formColor
        this.gui.SetFont("c" fontColor)  ; Set default text color for all controls

        ; Create edit control with matching background color and correct bold state
        this.editControl := this.gui.Add("Edit",
        "x5 y5 w" (editorWidth - 10) " h200 +Multi +WantReturn Background" this.note.bgcolor,
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
        editorWidth := Max(this.note.width, 200)  ; Use note width but minimum of 200
        ; Create format group below the edit control
        this.gui.Add("GroupBox", 
            "x5 y215 w" (editorWidth - 10) " h125 Background" formColor,
            "Formatting")


        ; Background color dropdown
        this.gui.Add("Text", "xp+10 yp+20", "Background:")
        colorDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w100", this.GetColorList())
        colorDropdown.Text := this.GetColorName(this.note.bgcolor)
        colorDropdown.OnEvent("Change", (*) => this.UpdateBackgroundColor(colorDropdown.Text))
        
        ; Font dropdown - full width
        this.gui.Add("Text", "x15 y+10", "Font:")
        fontDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w" ((editorWidth - 10) - 50),
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
        this.gui.Add("Text", "x15 y+14", "Font:")
        fontColorDropdown := this.gui.Add("DropDownList", "x+5 yp-3 w80", this.GetFontColorList())
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
        editorWidth := Max(this.note.width, 200)  ; Use note width but minimum of 200
        ; Add Note Options group
        this.gui.Add("GroupBox", 
            "x5 y+15 w" (editorWidth - 10) " h123",  ; Made taller for both controls
            "Note Options")

        ; Add alarm button
        addAlarmBtnText := this.note.hasAlarm 
            ? this.note.alarmTime  (this.note.alarmDays ?  "-" this.note.alarmDays : "") : "Add &Alarm..."
        this.addAlarmBtn := this.gui.Add("Button", "xp+6 yp+16 w" (editorWidth - 25), addAlarmBtnText)
        this.addAlarmBtn.OnEvent("Click", this.ShowAlarmDialog.Bind(this))
        
        LogError("Setting stick button text for note " this.note.id "`n")
        LogError("  isStuckToWindow: " this.note.isStuckToWindow "`n")
        LogError("  stuckWindowTitle: " this.note.stuckWindowTitle "`n")

        ; stick to window button
        stickBtnText := "Stick to &Window..."
        if (this.note.isStuckToWindow && this.note.stuckWindowTitle) {
            ; Truncate title if too long
            stickBtnText := StrLen(this.note.stuckWindowTitle) > 30 
                ? SubStr(this.note.stuckWindowTitle, 1, 27) "..."
                : this.note.stuckWindowTitle
            LogError("Setting stick button text to: " stickBtnText "`n")
        }

        this.stickToWindowBtn := this.gui.Add("Button", "xp y+5 w" (editorWidth - 25), stickBtnText)
        this.stickToWindowBtn.OnEvent("Click", this.ShowWindowStickyDialog.Bind(this))

        ; Add Width control with UpDown
        this.gui.Add("Text", "xp+4 yp+33", "Width:")
        this.widthEdit := this.gui.Add("Edit", "x+5 yp-3 w50 Number", this.note.width)
        widthUpDown := this.gui.Add("UpDown", "Range50-500", this.note.width)
        this.widthEdit.OnEvent("Change", (*) => this.UpdateWidth(this.widthEdit.Value))
        
        
        ; Add Always on Top checkbox below
        alwaysOnTopCB := this.gui.Add("Checkbox", "xp-35 yp+27", "Always on &Top")
        alwaysOnTopCB.Value := this.note.isOnTop
        alwaysOnTopCB.OnEvent("Click", (*) => this.ToggleAlwaysOnTop(alwaysOnTopCB.Value))

        
        ; Border option
        borderCB := this.gui.Add("Checkbox", "x+10", "&Border")
        borderCB.Value := this.note.hasBorder
        borderCB.OnEvent("Click", (*) => this.UpdateBorder(borderCB.Value))

        
        this.gui.Add("Button", "x5 y+15 w60", "&Save")
            .OnEvent("Click", (*) => this.Save())

        this.gui.Add("Button", "x+5 w60", "&Delete")
            .OnEvent("Click", this.DeleteNote.Bind(this))

        this.gui.Add("Button", "x+5 w60", "&Cancel")
            .OnEvent("Click", (*) => this.Hide())
    }

    ShowAlarmDialog(*) {
        ; Create alarm dialog using existing alarm clock GUI code
        if (!this.alarmDialog) {
            this.alarmDialog := AlarmDialog(this.note)
        }
        this.alarmDialog.Show()
    }

    ShowWindowStickyDialog(*) {
        if (this.note.isStuckToWindow) {
            ; If already stuck, unstick it
            this.note.isStuckToWindow := false
            this.note.stuckWindowTitle := ""
            this.note.stuckWindowClass := ""
            
            ; Update button text
            this.stickToWindowBtn.Text := "Stick to &Window..."
            
            ; Save changes
            storage := NoteStorage()
            storage.SaveNote(this.note)
            
            ; Update window following
            app.UpdateWindowFollowing()
            return
        }

        ; Create window picker dialog
        windowList := []
        stickyGui := Gui("+Owner" this.gui.Hwnd " +AlwaysOnTop", "Select Window")
        stickyGui.BackColor := formColor  ; Set background color
        stickyGui.SetFont("c" fontColor)  ; Set font color
        
        ; Add ListView for windows - with list color
        lv := stickyGui.Add("ListView", "w400 h200 Background" listColor, ["Window Title", "Class"])
        lv.ModifyCol(1, 270)
        lv.ModifyCol(2, 130)
        
        ; Populate ListView with visible windows
        ids := WinGetList(,, "Program Manager")
        for hwnd in ids {
            title := WinGetTitle(hwnd)
            class := WinGetClass(hwnd)
            
            ; Skip if no title or if it's our own GUI
            if (title = "" || hwnd = this.gui.Hwnd || hwnd = stickyGui.Hwnd)
                continue
                
            ; Skip certain system windows
            if (class ~= "i)^(Shell_TrayWnd|Windows.UI.Core.CoreWindow)$")
                continue
                
            lv.Add(, title, class)
            windowList.Push({title: title, class: class})
        }

        ; Handle selection (for both OK button and double-click)
        SaveSelection(*) {
            if (selected := lv.GetNext()) {
                ; Store current edit content before changes
                currentContent := this.editControl.Text
                
                ; Store previous AlwaysOnTop state before changing it
                this.note.previousAlwaysOnTop := this.note.isOnTop
                
                ; Set up window sticking
                this.note.isStuckToWindow := true
                this.note.stuckWindowTitle := windowList[selected].title
                this.note.stuckWindowClass := windowList[selected].class
                
                ; Force AlwaysOnTop
                this.note.isOnTop := true
                
                ; Update editor GUI
                this.gui.Destroy()
                this.CreateGui()
                
                ; Restore the edit content
                this.editControl.Text := currentContent
                
                this.Show()
                app.UpdateWindowFollowing()
                
                stickyGui.Destroy()
            }
        }

        ; Add double-click handler
        lv.OnEvent("DoubleClick", SaveSelection)

        ; Add buttons
        stickyGui.Add("Button", "Default x10 y+10 w80", "OK").OnEvent("Click", SaveSelection)
        stickyGui.Add("Button", "x+10 w80", "Cancel").OnEvent("Click", (*) => stickyGui.Destroy())
        
        stickyGui.Show()
    }

    ; And update the ToggleAlwaysOnTop method to actually set the property
    ToggleAlwaysOnTop(value) {
        this.note.isOnTop := value ? true : false
        if (this.note.borderGui) {
            this.note.borderGui.SetAlwaysOnTop(value)
        }
    }

    UpdateWidth(newWidth) {
        ; Ensure width is a positive number
        width := Max(50, Integer(newWidth))  ; Minimum width of 50 to prevent too-narrow notes
        
        ; Update stored width both in the note and editor
        this.note.width := width
        this.editorWidth := width
    }

    UpdateBorder(hasBorder) {
        this.note.hasBorder := hasBorder
    }
    
    Save(*) {
        ; Get current content
        newContent := this.editControl.Text
        
        ; Update note with properties
        this.note.UpdateContent(newContent, {
            bgcolor: this.note.bgcolor,
            font: this.note.font,
            fontSize: this.note.fontSize,
            isBold: this.note.isBold,
            fontColor: this.note.fontColor,
            isOnTop: this.note.isOnTop,  
            width: this.widthEdit.Value,
            hasBorder: this.note.hasBorder,
            hasAlarm: this.note.hasAlarm,
            alarmTime: this.note.alarmTime,
            alarmSound: this.note.alarmSound,
            alarmDays: this.note.alarmDays,
            alarmRepeatCount: this.note.alarmRepeatCount
        })

        ; Save to storage
        (NoteStorage()).SaveNote(this.note)
        
        this.gui.Destroy()
        this.CreateGui()
    }

    DeleteNote(*) {
        if (!OptionsConfig.WARN_ON_DELETE_NOTE || 
            MsgBox("Are you sure you want to delete this note?",, "YesNo 0x40000") = "Yes") {
            ; Use the app's noteManager to delete this note
            app.noteManager.DeleteNote(this.note.id)
        }
    }

    Show(*) {
        this.gui.Show("AutoSize")
    }
    
    Hide(*) {
        this.gui.Hide()
    }
    
    Destroy(*) {
        try {
            ; Clean up event bindings
            if (this.gui) {
                this.gui.OnEvent("Close", (*) => this.Hide(), 0)
                
                if (this.editControl) {
                    this.editControl := ""
                }

            if (this.alarmDialog) {
                this.alarmDialog.Destroy()
                this.alarmDialog := ""
            }
            
                this.gui.Destroy()
                this.gui := ""
            }

            return true
        } catch as err {
            LogError("Error in NoteEditor.Destroy: " err.Message)
            return false
        }
    }
}

class AlarmDialog {
    gui := ""
    note := ""
    hourEdit := ""
    minuteEdit := ""
    soundDropDown := ""
    weekdayChecks := Map()
    
    __New(note) {
        this.note := note
        this.CreateGui()
    }
    
    CreateGui() {
        ; Create alarm settings window
        this.gui := Gui("+Owner +AlwaysOnTop", "Set Alarm")
        this.gui.BackColor := formColor
        this.gui.SetFont("c" fontColor)

        ; Time inputs
        this.gui.AddText("y10", "Time: ")
        this.hourEdit := this.gui.AddEdit("x+5 yp-3 w45 Background" listColor, "8")
        hourUpDown := this.gui.AddUpDown("+wrap Range1-12", 8)

        this.hourEdit.OnEvent("Change", this.CheckAMPM.Bind(this))
        hourUpDown.OnEvent("Change", this.CheckAMPM.Bind(this))
        
        this.gui.AddText("x+5 yp+3", ":")
        this.minuteEdit := this.gui.AddEdit("x+5 yp-3 w60 Background" listColor, "00")
        minuteUpDown := this.gui.AddUpDown("+wrap Range0-59", 0)
        this.minuteEdit.OnEvent("Change", this.UpdateMinuteFormat.Bind(this))
        minuteUpDown.OnEvent("Change", this.UpdateMinuteFormat.Bind(this))
        

        ; AM/PM radio buttons
        varAM := this.gui.AddRadio("x+15 yp+3 Group Checked", "AM")
        varPM := this.gui.AddRadio("x+4 yp", "PM")

        ; Sound selection
        this.gui.AddText("xm y+10", "Alarm Sound:")
        this.soundDropDown := this.gui.AddDropDownList("x+5 yp-3 w200 Background" listColor, AlarmConfig.GetSoundFiles())
        this.soundDropDown.Choose(1)
        
        testButton := this.gui.AddButton("x+5 yp-2 w45 h24", "Test")
        testButton.OnEvent("Click", this.TestSound.Bind(this))

        ; Repeat options in AlarmDialog's CreateGui method
        this.repeatOnce := this.gui.AddRadio("Group xm", "Once")
        this.repeatGroup2 := this.gui.AddRadio("x+4", "3 times")
        this.repeatGroup3 := this.gui.AddRadio("x+4", "10 times")
        ; Shake box
        this.visualShake := this.gui.AddCheckbox("x+60 yp", "Visual Shake")
        ; Set defaults
        this.visualShake.Value := OptionsConfig.DEFAULT_ALARM_SHAKE 
        this.repeatOnce.Value := 1

        ; Weekday checkboxes
        this.gui.AddText("xm y+15", "Reoccur:")
        days := ["Su", "Mo", "Tu", "We", "Th", "Fr", "Sa"]
        for i, day in days {
            this.weekdayChecks[day] := this.gui.AddCheckbox("x+5 yp", day)
        }
        
        ; Save button
        saveBtn := this.gui.AddButton("xm y+15 w75", "&Save")
        saveBtn.OnEvent("Click", this.SaveAlarm.Bind(this))

        ; Stop button
        stopBtn := this.gui.AddButton("x+5 yp w75", "&Stop")
        try stopBtn.OnEvent("Click", this.StopSound.Bind(this))

        ; Delete button
        deleteBtn := this.gui.AddButton("x+5 yp w75", "&Delete")
        deleteBtn.OnEvent("Click", this.DeleteAlarm.Bind(this))

        ; Cancel button
        cancelBtn := this.gui.AddButton("x+5 yp w75", "&Cancel")
        cancelBtn.OnEvent("Click", (*) => this.Destroy())
        
        ; Either load existing alarm or set default time
        if (this.note.hasAlarm)
            this.LoadExistingAlarm()
        else {
            ; Set time to one minute from now
            futureTime := DateAdd(A_Now, 1, "Minutes")
            this.hourEdit.Value := FormatTime(futureTime, "h")
            this.minuteEdit.Value := FormatTime(futureTime, "mm")
            if (FormatTime(futureTime, "tt") = "PM")
                this.gui["PM"].Value := 1
            else
                this.gui["AM"].Value := 1
        }
    }

    UpdateMinuteFormat(*) { ; Pads zeros and single digits. 00, 01, ... 09.
        sleep 500 ; The sleep allows time to type e.g. "30" without getting "03".
        try {
            value := Integer(this.minuteEdit.Value)
            if (value >= 0 && value < 10)
                this.minuteEdit.Value := Format("{:02d}", value)
        } catch Error as err {
            this.minuteEdit.Value := "00"
        }
    }

    CheckAMPM(*) { ; Sets AM/PM based on hour
        val := Integer(this.hourEdit.Value)
        if (val >= 5 && val <= 11)
            this.gui["AM"].Value := 1
        else
            this.gui["PM"].Value := 1
    }

    LoadExistingAlarm() {
        ; Parse existing time
        if (this.note.alarmTime) {
            RegExMatch(this.note.alarmTime, "(\d+):(\d+)\s*(AM|PM)", &match)
            this.hourEdit.Value := match[1]
            this.minuteEdit.Value := match[2]
            if (match[3] = "PM")
                this.gui["PM"].Value := 1
            else
                this.gui["AM"].Value := 1
        }
        
        ; Set sound if exists
        if (this.note.alarmSound) {
            ; Extract just the filename from the full path
            SplitPath(this.note.alarmSound, &fileName)
            this.soundDropDown.Choose(fileName)
        }
            
        ; Set weekdays
        if (this.note.alarmDays) {
            for day, checkbox in this.weekdayChecks {
                checkbox.Value := InStr(this.note.alarmDays, day) > 0
            }
        }
        
        ; Set repeat count based on note's stored value
        if (this.note.alarmRepeatCount = 1)
            this.repeatOnce.Value := 1
        else if (this.note.alarmRepeatCount = 3)
            this.repeatGroup2.Value := 1
        else
            this.repeatGroup3.Value := 1
    }
    
    TestSound(*) {
        selectedSound := this.soundDropDown.Text
        soundPath := AlarmConfig.ALARM_SOUNDS_FOLDER "\" selectedSound
        if FileExist(soundPath) {
            try {
                SoundPlay(soundPath, 1)
                this.currentlyPlaying := true
            }
        } else
            SoundBeep()
    }

    StopSound(*) {
        try if (this.currentlyPlaying) {
            try SoundPlay("nonexistent.wav")
            this.currentlyPlaying := false
        }
    }
    
    SaveAlarm(*) {
        ; Format time
        hour := this.hourEdit.Value
        minute := this.minuteEdit.Value
        ampm := this.gui["PM"].Value ? "PM" : "AM"
        newTime := Format("{1}:{2:02d} {3}", hour, minute, ampm)
        
        ; Always reset cycle state when saving any alarm
        LogError("Resetting cycle state for note " this.note.id "`n")
        app.ResetNoteAlarmCycle(this.note.id)
        
        ; Update alarm properties
        this.note.alarmTime := newTime
        this.note.hasAlarm := true
        
        LogError("New alarm set for note " this.note.id " at " newTime "`n")
        
        ; Save sound
        this.note.alarmSound := AlarmConfig.ALARM_SOUNDS_FOLDER "\" this.soundDropDown.Text
        
        ; Get weekdays
        activeDays := ""
        for day, checkbox in this.weekdayChecks {
            if (checkbox.Value)
                activeDays .= day
        }
        this.note.alarmDays := AlarmConfig.SortDays(activeDays)
        
        ; Save repeat count
        if (this.repeatOnce.Value)
            this.note.alarmRepeatCount := 1
        else if (this.repeatGroup2.Value)
            this.note.alarmRepeatCount := 3
        else if (this.repeatGroup3.Value)
            this.note.alarmRepeatCount := 10
        else  ; Default to once if somehow none are selected
            this.note.alarmRepeatCount := 1

        this.note.visualShake := this.visualShake.Value
            
        ; Save to storage immediately
        storage := NoteStorage()
        storage.SaveNote(this.note)

        ; If note editor is open, update the alarm button text
        if (this.note.editor) {
        buttonText := this.note.hasAlarm 
            ? "&A: " this.note.alarmTime (this.note.alarmDays ? " - " this.note.alarmDays : "")
            : "Add &Alarm..."
        this.note.editor.addAlarmBtn.Text := buttonText
    }

        ; Close dialog
        this.Destroy()
    }

    DeleteAlarm(*) {
        try {
            if (MsgBox("Are you sure you want to delete this alarm?",, "YesNo 0x40000") = "Yes") {
                ; Reset alarm properties
                this.note.hasAlarm := false
                this.note.alarmTime := ""
                this.note.alarmSound := ""
                this.note.alarmDays := ""
                this.note.alarmRepeatCount := 1
                
                ; Force reset of cycle state
                if (StickyNotes.cycleComplete && StickyNotes.cycleComplete.Has(this.note.id)) {
                    StickyNotes.cycleComplete.Delete(this.note.id)
                }
                
                ; Save changes to storage
                storage := NoteStorage()
                storage.SaveNote(this.note)
                
                ; If note editor is open, update the alarm button text
                if (this.note.editor) {
                    this.note.editor.addAlarmBtn.Text := "Add &Alarm..."
                }
                
                ; Close alarm dialog
                this.Destroy()
            }
        } catch Error as err {
            LogError("Error in AlarmDialog.DeleteAlarm: " err.Message "`n")
        }
    }

    Show() {
        try {
            if (!this.gui) {
                this.CreateGui()
            }
            if (!this.gui) {
                LogError("GUI is null after CreateGui in AlarmDialog.Show")
                return
            }
            this.gui.Show()
        } catch as err {
            LogError("Error in AlarmDialog.Show: " err.Message)
        }
    }

    Destroy() {
        if (this.gui) {
            this.gui.Destroy()
            this.gui := ""
        }
    }
}

/*
Colorized rows in ListView.
Based on Justme's Class LV_Colors
https://www.autohotkey.com/boards/viewtopic.php?t=93922
Claude was instructed to remove the parts of the class that prevent 
sorting/filtering, alow alternating colors, and allow column customization.
It was told to only keep the part for colorizing individual rows.
*/
Class LV_Colors {
    __New(LV) {
        this.LV := LV
        this.HWND := LV.HWND
        
        ; Set LVS_EX_DOUBLEBUFFER style
        LV.Opt("+LV0x010000")
        
        this.Rows := Map()
        this.ShowColors()
    }
    
    __Delete() {
        this.ShowColors(false)
    }
    
    Row(Row, BkColor, TxColor) {
        if !(this.HWND)
            return false
            
        this.Rows[Row] := Map("B", this.BGR(BkColor), "T", this.BGR(TxColor))
        return true
    }
    
    ShowColors(Apply := true) {
        if (Apply) && !this.HasProp("OnNotifyFunc") {
            this.OnNotifyFunc := ObjBindMethod(this, "OnNotify")
            this.LV.OnNotify(-12, this.OnNotifyFunc)
        }
        return true
    }
    
    OnNotify(LV, L) {
        Critical -1
        
        static SizeNMHDR := A_PtrSize * 3
        static OffItem := SizeNMHDR + 16 + (A_PtrSize * 2)
        static OffCT := SizeNMHDR + 16 + (A_PtrSize * 5)
        static OffCB := OffCT + 4
        
        if (NumGet(L, "UPtr") != this.HWND)
            return
            
        DrawStage := NumGet(L + SizeNMHDR, "UInt")
        Row := NumGet(L + OffItem, "UPtr") + 1
        
        if (DrawStage = 0x1)  ; CDDS_PREPAINT
            return 0x20
            
        if (DrawStage = 0x10001) {  ; CDDS_ITEMPREPAINT
            if this.Rows.Has(Row) {
                NumPut("UInt", this.Rows[Row]["T"], L + OffCT)
                NumPut("UInt", this.Rows[Row]["B"], L + OffCB)
            }
            return 0x2  ; CDRF_NEWFONT
        }
        return 0
    }
    
    BGR(Color) {
        if IsInteger(Color)
            return ((Color >> 16) & 0xFF) | (Color & 0x00FF00) | ((Color & 0xFF) << 16)
        return 0
    }
}

class MainWindow {
    ; Window properties
    gui := ""
    noteList := ""     ; ListView control
    CLV := ""         ; Color handler instance
    noteRowMap := Map() ; Map for tracking note IDs
    filterHidden := ""  ; Checkbox control
    filterVisible := "" ; Checkbox control
    searchEdit := ""    ; Search box control
    
    __New() {
        this.CreateGui()
    }
    
    CreateGui() {
        buttonW := 140  ; button width
        spacing := 10   ; Space between buttons
        totalWidth := (buttonW * 3) + (spacing * 2)  ; Calculate total width for 3 columns

        ; Create main window
        this.gui := Gui("+AlwaysOnTop +Resize", "Sticky Notes Manager")
        this.gui.BackColor := formColor
        buttonW := 140  ; Button width
        spacing := 10   ; Space between buttons
        helpTextFont := 11

        this.gui.setFont("s" helpTextFont + 1 " w700")
        this.gui.SetFont("c" fontColor)
        this.gui.Add("Text", " x150 y10", "Sticky Notes Manager")
        this.gui.setFont("s" helpTextFont " w400")
        this.gui.Add("Button", " x+85 y10 h25 w40 ", "Tips")
            .OnEvent("Click", (*) => this.ShowTips())

        ; Add separator
        this.gui.Add("Text", "x10 y+10 w" totalWidth " h2 0x10")  ; Horizontal line

        ; Create button group 1 - Core Functions
        this.gui.Add("Button", "x10 y+10 w" buttonW " h30", "&New Note")
            .OnEvent("Click", (*) => app.CreateNewNote())
        
        this.gui.Add("Button", "x" (buttonW + spacing + 10) " yp w" buttonW " h30", "New From Clpbrd")
            .OnEvent("Click", (*) => app.CreateClipboardNote())
            
        this.gui.Add("Button", "x" (buttonW * 2 + spacing * 2 + 10) " yp w" buttonW " h30", "&Load Notes")
            .OnEvent("Click", (*) => app.noteManager.LoadSavedNotes())
            
        this.gui.Add("Button", "x10 y+5 w" buttonW " h30", "Show Hidden")
            .OnEvent("Click", (*) => app.noteManager.ShowHiddenNotes())
            
        this.gui.Add("Button", "x" (buttonW + spacing + 10) " yp w" buttonW " h30", "Save Status")
            .OnEvent("Click", (*) => app.noteManager.SaveAllNotes())
            
        this.gui.Add("Button", "x" (buttonW * 2 + spacing * 2 + 10) " yp w" buttonW " h30", "Exit App")
            .OnEvent("Click", (*) => this.ExitApp())

        this.gui.Add("Button", "x10 y+5 w" totalWidth " h30", "Hide This Window")
            .OnEvent("Click", (*) => this.Hide())
            
        ; Add separator before ListView
        this.gui.Add("Text", "x10 y+10 w" totalWidth " h2 0x10")  ; Horizontal line

        ; Add filter controls
        this.gui.Add("Text", "x10 y+10", "Show:")
        this.filterHidden := this.gui.Add("Checkbox", "x+5 yp Checked", "Hidden")
        this.filterVisible := this.gui.Add("Checkbox", "x+10 yp Checked", "Visible")
        this.searchEdit := this.gui.Add("Edit", "x+10 yp-3 w240", "")

        ; Add filter events
        this.filterHidden.OnEvent("Click", (*) => this.PopulateNoteList())
        this.filterVisible.OnEvent("Click", (*) => this.PopulateNoteList())
        this.searchEdit.OnEvent("Change", (*) => this.PopulateNoteList())

        ; Create ListView with columns
        this.noteList := this.gui.Add("ListView", 
            "x10 y+5 w" totalWidth " h200 Background" listColor, 
            ["Created", "Note Contents", "Alarm|Window"])

        ; Set column widths
        this.noteList.ModifyCol(1, 70)
        this.noteList.ModifyCol(2, totalWidth * 0.60) 
        this.noteList.ModifyCol(3, totalWidth * 0.45)

        ; Create color handler instance
        this.CLV := LV_Colors(this.noteList)

        ;  handlers
        this.noteList.OnEvent("DoubleClick", this.EditSelectedNote.Bind(this))
        this.noteList.OnEvent("ContextMenu", this.HandleClick.Bind(this))

        buttonY := "y+10"
        this.gui.Add("Button", "x10 " buttonY " w" buttonW " h30", "Edit Note")
            .OnEvent("Click", (*) => this.EditSelectedNote())
        this.gui.Add("Button", "x" (buttonW + spacing + 10) " yp w" buttonW " h30", "Bring Forward")
            .OnEvent("Click", (*) => this.ShowSelectedNote())
        this.gui.Add("Button", "x" (buttonW * 2 + spacing * 2 + 10) " yp w" buttonW " h30", "&Delete Note")
            .OnEvent("Click", (*) => this.DeleteSelectedNote())
        this.gui.Add("Button", "x10 y+5 w" buttonW " h30", "&Hide Note")
            .OnEvent("Click", (*) => this.HideSelectedNote())
        this.gui.Add("Button", "x" (buttonW + spacing + 10) " yp w" buttonW " h30", "&Unhide Note")
            .OnEvent("Click", (*) => this.UnhideSelectedNote())
        this.gui.Add("Button", "x" (buttonW * 2 + spacing * 2 + 10) " yp w" buttonW " h30", "Open ini file")
            .OnEvent("Click", (*) => Run(OptionsConfig.INI_FILE))
        
        ; Set up events
        this.gui.OnEvent("Close", (*) => this.Hide())
        this.gui.OnEvent("Escape", (*) => this.Hide())
    }

    ShowTips(*) {
        helpText :=  "~~~Customizable Hotkeys~~~`n"
            . FormatHotkeyForDisplay(OptionsConfig.NEW_NOTE) " = New Note`n"
            . FormatHotkeyForDisplay(OptionsConfig.NEW_CLIPBOARD_NOTE) " = Clipboard Note`n"
            . FormatHotkeyForDisplay(OptionsConfig.TOGGLE_MAIN_WINDOW) " = Show/Hide Window`n"
            . (OptionsConfig.CHECKBOX_MODIFIER_KEY? OptionsConfig.CHECKBOX_MODIFIER_KEY "+Click = Toggle Checkbox in sticky note`n`n" : "`n")
            . "~~~Tips for Sticky Note Manager~~~`n"
            . "Select list item, then double-click for editor.  Note item must be unhidden before editing or deleting. Right-click for 2 second preview of note content.  Let previous disappear before viewing the next one. Ctrl+Click to select muliple note items.  Notes can be bulk hidden, unhidden, or deleted.`n`n"
            . "~~~Tips for Sticky Notes~~~`n"
            . "If an alarm is playing, you can right-click the sticky note gui for a 'Stop Alarm' command.  The top/center 'title bar' area of the note is the drag area.  Reposition a note by dragging its drag area.  Open note in Note Editor by double-clicking drag area."
        
        msgbox(helpText,"Help Tips", "Owner" this.gui.Hwnd)
        
        FormatHotkeyForDisplay(hotkeyStr) {
        return (InStr(hotkeyStr, "^") ? "Ctrl+" : "")
            . (InStr(hotkeyStr, "+") ? "Shift+" : "")
            . (InStr(hotkeyStr, "!") ? "Alt+" : "")
            . (InStr(hotkeyStr, "#") ? "Win+" : "")
            . StrUpper(SubStr(hotkeyStr, RegExMatch(hotkeyStr, "[a-zA-Z0-9]")))
        }
    }

    PopulateNoteList() {
        LogError("Starting PopulateNoteList at " A_TickCount "`n")
        this.noteList.Delete()
        this.noteRowMap.Clear()
        
        storage := NoteStorage()
        savedNotes := storage.LoadAllNotes()
        searchText := this.searchEdit.Text
        
        for noteData in savedNotes {
            isHidden := Integer(noteData.isHidden)
            if ((!this.filterHidden.Value && isHidden) || 
                (!this.filterVisible.Value && !isHidden))
                continue
                
            if (searchText && !InStr(noteData.content, searchText, true))
                continue
                
            if (RegExMatch(noteData.id, "(\d{4})(\d{2})(\d{2})", &match))
                creationDate := SubStr(match[1], 3) "-" match[2] "-" match[3]
            else
                creationDate := ""
                
            preview := StrLen(noteData.content) > 50 
                ? SubStr(noteData.content, 1, 47) "..."
                : noteData.content
                
            ; Build combined alarm/window info
            combinedInfo := ""

            ; Add alarm info if present
            if (noteData.hasAlarm) {
                combinedInfo := noteData.alarmTime (noteData.alarmDays ? " (" AlarmConfig.SortDays(noteData.alarmDays) ")" : "")
            }

            ; Add window info if present
            if (noteData.isStuckToWindow && noteData.stuckWindowTitle) {
                windowInfo := StrLen(noteData.stuckWindowTitle) > 30 
                    ? SubStr(noteData.stuckWindowTitle, 1, 27) "..."
                    : noteData.stuckWindowTitle
                    
                combinedInfo := combinedInfo 
                    ? combinedInfo "|" windowInfo  ; Add window info with separator if there's already alarm info
                    : windowInfo                   ; Just window info if no alarm
            }
            
            rowNum := this.noteList.Add(, creationDate, preview, combinedInfo, "")
            this.noteRowMap[rowNum] := noteData.id
            
            this.CLV.Row(rowNum, 
                "0x" noteData.bgcolor,
                "0x" noteData.fontColor)
        }

        LogError("PopulateNoteList completed at " A_TickCount "`n")
        ; Select the last row if any rows exist
        if (lastRow := this.noteList.GetCount()) {
            this.noteList.Modify(lastRow, "Select Focus")
        }
    }

    HandleClick(ctrl, info*) {
        static tooltipShowing := false
        
        try {
            ; Get selected row
            row := this.noteList.GetNext(0)  
            if (!row || !this.noteRowMap.Has(row))
                return
                
            ; Get note data from storage
            storage := NoteStorage()
            noteData := storage.LoadNote(this.noteRowMap[row])
            if (!noteData)
                return
                
            ; Initialize tooltip system if needed
            static tooltipInitialized := false
            if (!tooltipInitialized) {
                ToolTipOpts.Init()
                tooltipInitialized := true
            }
            
            ; Clean up content
            content := StrReplace(noteData.content, "`r`n", "`n")  
            content := StrReplace(content, "`r", "`n")    
            
            ; Set tooltip formatting
            ToolTipOpts.SetColors(noteData.bgcolor, noteData.fontColor)
            ToolTipOpts.SetFont(noteData.fontSize, noteData.font, noteData.isBold)
            ToolTipOpts.SetMaxWidth(noteData.width + 40)
            
            ; Debug logging
            LogError("Showing tooltip for note " this.noteRowMap[row] "`n"
                . "  Font: " noteData.font " Size: " noteData.fontSize 
                . " Bold: " noteData.isBold "`n"
                . "  Colors - Bg: #" noteData.bgcolor " Font: #" noteData.fontColor "`n")
            
            ; Show tooltip
            ToolTip(content)
            tooltipShowing := true
            
            ; Set up timer to hide tooltip after 2 seconds
            SetTimer () => (ToolTip(), tooltipShowing := false), -2000
            
        } catch Error as err {
            LogError("Error in HandleClick: " err.Message "`n")
            ToolTip()
            tooltipShowing := false
        }
    }
            
    GetSelectedNoteIds() {
        selectedIds := []
        row := 0
        while (row := this.noteList.GetNext(row)) {
            if (this.noteRowMap.Has(row))
                selectedIds.Push(this.noteRowMap[row])
        }
        return selectedIds
    }

    EditSelectedNote(*) {
        selectedIds := this.GetSelectedNoteIds()
        if (selectedIds.Length = 0)
            return
            
        if (selectedIds.Length > 1) {
                MsgBox("Please select only one note to edit.", "Edit Note", "Icon! Owner" this.gui.Hwnd)
            return
        }
        
        if app.noteManager.notes.Has(selectedIds[1])
            app.noteManager.notes[selectedIds[1]].Edit()
    }

    ShowSelectedNote(*) {
        selectedIds := this.GetSelectedNoteIds()
        if (selectedIds.Length = 0)
            return
            
        for noteId in selectedIds {
            if app.noteManager.notes.Has(noteId)
                app.noteManager.notes[noteId].Show()
        }
    }

    DeleteSelectedNote(*) {
        selectedIds := this.GetSelectedNoteIds()
        if (selectedIds.Length = 0)
            return
            
        if (selectedIds.Length = 1) {
            ; Single note deletion - show preview
            storage := NoteStorage()
            noteData := storage.LoadNote(selectedIds[1])
            if (noteData) {
                if (!OptionsConfig.WARN_ON_DELETE_NOTE) {
                    ; Skip warning if disabled
                    this.gui.Opt("+Disabled")  ; Prevent user interaction during deletion
                    app.noteManager.DeleteNote(selectedIds[1])
                    this.gui.Opt("-Disabled")  ; Re-enable user interaction
                    this.PopulateNoteList()    ; Single refresh
                } else {
                    preview := noteData.content
                    if (StrLen(preview) > 500)
                        preview := SubStr(preview, 1, 497) "..."
                    
                    if (MsgBox("Are you sure you want to delete this note?`n`nContent:`n" preview,
                        "Delete Note", "YesNo 0x30 Owner" this.gui.Hwnd) = "Yes") {
                        this.gui.Opt("+Disabled")  ; Prevent user interaction during deletion
                        app.noteManager.DeleteNote(selectedIds[1])
                        this.gui.Opt("-Disabled")  ; Re-enable user interaction
                        this.PopulateNoteList()    ; Single refresh
                    }
                }
            }
        } else {
            ; Multiple note deletion - always show warning regardless of setting
            if (MsgBox("Are you sure you want to delete " selectedIds.Length " notes?",
                "Delete Multiple Notes", "YesNo 0x30 Owner" this.gui.Hwnd) = "Yes") {
                ; Temporarily disable list updates
                this.gui.Opt("+Disabled")  ; Prevent user interaction during deletion
                
                ; Modify NoteManager.DeleteNote to not trigger listview updates
                needsRefresh := this.noteList.Visible  ; Store listview visibility state
                if (needsRefresh)
                    this.noteList.Visible := false
                    
                for noteId in selectedIds
                    app.noteManager.DeleteNote(noteId)
                    
                ; Restore listview and do single refresh
                if (needsRefresh) {
                    this.noteList.Visible := true
                    this.PopulateNoteList()    ; Single refresh after all deletions
                }
                
                this.gui.Opt("-Disabled")  ; Re-enable user interaction
            }
        }
    }

    HideSelectedNote(*) {
        selectedIds := this.GetSelectedNoteIds()
        if (selectedIds.Length = 0)
            return
            
        for noteId in selectedIds {
            if app.noteManager.notes.Has(noteId)
                app.noteManager.notes[noteId].Hide()
        }
        this.PopulateNoteList()
    }

    UnhideSelectedNote(*) {
        selectedIds := this.GetSelectedNoteIds()
        if (selectedIds.Length = 0)
            return
            
        for noteId in selectedIds
            app.noteManager.RestoreNote(noteId)
        this.PopulateNoteList()
    }

    Show(*) {
        this.PopulateNoteList()
        ; Position window near top of screen
        this.gui.Show("AutoSize y50")
        ; Focus the ListView by default
        this.noteList.Focus()
    }
    
    Hide(*) {
        this.gui.Hide()
    }
    
    Destroy(*) {
        this.gui.Destroy()
    }
    
    ExitApp(*) {
        ; Save all notes and clean up using the noteManager
        app.noteManager.SaveAllNotes()
        
        ; Clean up all notes
        for id, note in app.noteManager.notes {
            if (!note.Destroy()) {
                LogError("Failed to destroy note " id " during exit")
            }
        }
        
        ; Clear the notes collection
        app.noteManager.notes := Map()
        
        ExitApp()
    }
}

/*
CLASS TOOLTIPOPTS
The code (much like that name) is a subset of, and 
based directly from, Justme's 'class TooltipOptions' here
https://www.autohotkey.com/boards/viewtopic.php?f=83&t=113308
Distilled by kunkel321, using Claude AI on 2-9-2025.
It keeps the parts of Justme's code that does colors and font size and style.
Adds ability to set width of tooltips and bold font.
*/
Class ToolTipOpts {
    Static HTT := DllCall("User32.dll\CreateWindowEx", "UInt", 8, "Str", "tooltips_class32", "Ptr", 0, "UInt", 3
                        , "Int", 0, "Int", 0, "Int", 0, "Int", 0, "Ptr", A_ScriptHwnd, "Ptr", 0, "Ptr", 0, "Ptr", 0) ;*HTT = Handle to ToolTip (the window handle for the tooltip control)
    Static SWP := CallbackCreate(ObjBindMethod(ToolTipOpts, "_WNDPROC_"), , 4) ;*SWP = SubclassWindowProc (the callback function for window subclassing)
    Static OWP := 0, ToolTips := Map() ; *OWP = Original Window Proc (stores the original window procedure)
    ; Properties
    Static BkgColor := "", TxtColor := "", HFONT := 0, MaxWidth := 0 ; *HFONT = Handle to Font

    Static Call(*) => False
    Static Init() {
        If (This.OWP = 0) {
            This.BkgColor := "", This.TxtColor := ""
            If (A_PtrSize = 8)
                This.OWP := DllCall("User32.dll\SetClassLongPtr", "Ptr", This.HTT, "Int", -24, "Ptr", This.SWP, "UPtr")
            Else
                This.OWP := DllCall("User32.dll\SetClassLongW", "Ptr", This.HTT, "Int", -24, "Int", This.SWP, "UInt")
            OnExit(ToolTipOpts._EXIT_, -1)
            Return This.OWP
        }
        Return False
    }
    
    Static SetColors(BkgColor := "", TxtColor := "") {
        This.BkgColor := BkgColor = "" ? "" : This.BGR(BkgColor)
        This.TxtColor := TxtColor = "" ? "" : This.BGR(TxtColor)
    }
    
    Static BGR(Color, Default := "") {
        Static HTML := {BLACK: 0x000000, WHITE: 0xFFFFFF, RED: 0x0000FF, GREEN: 0x00FF00, BLUE: 0xFF0000}
        
        If HTML.HasProp(Color)
            Return HTML.%Color%
        If (Color Is String) && IsXDigit(Color) && (StrLen(Color) = 6)
            Color := Integer("0x" . Color)
        If IsInteger(Color)
            Return ((Color >> 16) & 0xFF) | (Color & 0x00FF00) | ((Color & 0xFF) << 16)
        Return Default
    }
    
    Static SetFont(Size := "", FontName := "", Bold := False) {  ; Added Bold parameter
        Static LOGFONTW := Buffer(92, 0)
        Static HDEF := DllCall("GetStockObject", "Int", 17, "UPtr")
        
        If (Size = "") && (FontName = "") {
            If This.HFONT
                DllCall("DeleteObject", "Ptr", This.HFONT)
            This.HFONT := 0
            Return
        }
        
        HDC := DllCall("GetDC", "Ptr", 0, "UPtr") ; HDC = Handle to Device Context (for display/graphics operations)
        LOGPIXELSY := DllCall("GetDeviceCaps", "Ptr", HDC, "Int", 90, "Int")
        DllCall("ReleaseDC", "Ptr", 0, "Ptr", HDC)
        
        DllCall("GetObject", "Ptr", HDEF, "Int", 92, "Ptr", LOGFONTW) ; HDC = Handle to Device Context (for display/graphics operations).  LOGFONTW = Logical Font Windows (W stands for Wide/Unicode version.This is a Windows structure that defines the attributes of a font)
        
        If (Size != "")
            NumPut("Int", -Round(Size * LOGPIXELSY / 72), LOGFONTW)
        
        ; Set font weight (Regular = 400, Bold = 700)
        NumPut("Int", Bold ? 700 : 400, LOGFONTW, 16)
        
        If (FontName != "")
            StrPut(FontName, LOGFONTW.Ptr + 28, 32)
        
        If HFONT := DllCall("CreateFontIndirectW", "Ptr", LOGFONTW, "UPtr") {
            If This.HFONT
                DllCall("DeleteObject", "Ptr", This.HFONT)
            This.HFONT := HFONT
        }
    }
    
    Static SetMaxWidth(Width := 0) {
        This.MaxWidth := Width
    }
    
    Static _WNDPROC_(hWnd, uMsg, wParam, lParam) {
        Switch uMsg {
            Case 0x0411:  ; TTM_TRACKACTIVATE
                If This.ToolTips.Has(hWnd) && (This.ToolTips[hWnd] = 0) {
                    If (This.BkgColor != "")
                        SendMessage(0x413, This.BkgColor, 0, hWnd)  ; TTM_SETTIPBKCOLOR *TTM = ToolTip Message
                    If (This.TxtColor != "")
                        SendMessage(0x414, This.TxtColor, 0, hWnd)  ; TTM_SETTIPTEXTCOLOR
                    If This.HFONT
                        SendMessage(0x30, This.HFONT, 0, hWnd)     ; WM_SETFONT *WM = Window Message
                    If This.MaxWidth
                        SendMessage(0x418, 0, This.MaxWidth, hWnd)  ; TTM_SETMAXTIPWIDTH
                    This.ToolTips[hWnd] := 1
                }
            Case 0x0001:  ; WM_CREATE
                DllCall("UxTheme.dll\SetWindowTheme", "Ptr", hWnd, "Ptr", 0, "Ptr", StrPtr(""))
                This.ToolTips[hWnd] := 0
            Case 0x0002:  ; WM_DESTROY
                This.ToolTips.Delete(hWnd)
        }
        Return DllCall(This.OWP, "Ptr", hWnd, "UInt", uMsg, "Ptr", wParam, "Ptr", lParam, "UInt")
    }
    
    Static _EXIT_(*) {
        If (ToolTipOpts.OWP != 0) {
            For HWND In ToolTipOpts.ToolTips.Clone()
                DllCall("DestroyWindow", "Ptr", HWND)
            ToolTipOpts.ToolTips.Clear()
            If ToolTipOpts.HFONT
                DllCall("DeleteObject", "Ptr", ToolTipOpts.HFONT)
            If (A_PtrSize = 8)
                DllCall("User32.dll\SetClassLongPtrW", "Ptr", ToolTipOpts.HTT, "Int", -24, "Ptr", ToolTipOpts.OWP, "UPtr")
            Else
                DllCall("User32.dll\SetClassLongW", "Ptr", ToolTipOpts.HTT, "Int", -24, "Int", ToolTipOpts.OWP, "UInt")
            ToolTipOpts.OWP := 0
        }
    }
}

; Helper function for conditional logging
LogError(message) {
    if (OptionsConfig.ERROR_LOG) {
        FileAppend(message "`n", "error_log.txt")
    }
}
