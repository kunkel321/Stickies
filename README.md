# Sticky Notes
(from code comments)
* Project:    Sticky Notes
* Author:     kunkel321
* Tool used:  Claude AI
* Version:    (Version date will be updated in code.)
* Forum:      https://www.autohotkey.com/boards/viewtopic.php?f=83&t=135340
* Repository: https://github.com/kunkel321/Stickies     

Customizable Hotkeys:
--------------------
* Win+Shift+N - Create new note
* Win+Shift+C - Create new note from clipboard text
* Win+Shift+S - Toggle main window visibility

Features, Functionality, and Usage:
-----------------------------------
- Create sticky notes that persist between script restarts
- Notes cascade from top-left of screen for better visibility
- Cascading Positions: New note doesn't fully cover previous note
- Notes cycle through different pastel background colors automatically
- Several formatting options including fonts, colors, sizing, and borders
- Set alarms for individual notes with optional weekly recurrence
- Main window with note management, preview, and search functionality
- Access main window with Win+Shift+S (hidden by default)
- Notes are created using Win+Shift+N or from clipboard with Win+Shift+C
- Hotkeys can be changed near top of code
- Tips button in note manager shows hotkeys and other tips
- Drag notes by their top bar to reposition
- Notes auto-save position when moved
- Double-click top bar or right-click for editing options
- System tray icon provides quick access to common functions
- Start with Windows: Option available in tray menu
- Checkbox Creation: Any text line starting with [] or [x] becomes an interactive checkbox
- Checkbox Safety: Alt+Click required by default to prevent accidental toggles
- Alarm System: Set one-time or recurring alarms with custom sounds
- All note data saved to sticky_notes.ini in script directory
- Hidden notes can be restored through main window

Tips:
-----
- Use manual "Save Status" after significant changes
- "Load/Reload Notes" refreshes all notes from storage
- Check error_log.txt for troubleshooting (if enabled)
- Use main window's search to find specific notes or filter by hidden/unhidden
- Right-click menu provides quick access to several functions
- Preview notes in note manager listview via right-click


Development Note:
----------------
This script was developed primarily through AI-assisted coding, with Claude AI generating most of the base code structure and functionality. Later versions 
include additional human-written code for enhanced features and bug fixes.. The system tray context menu has a few extra items, such as "Start with Windows."

This script is unique because nearly every bit of code was created with AI prompts, then pasted into the ahk editor.  A great deal of human input was needed, but very little of the actual code was human-generated.  Edit: In later versions, more human code was added.
![Screenshot of Sticky Notes tool](https://i.imgur.com/GKEZ5er.png)
Screenshot explanation:
1.	1.	Sticky Note. Top/Center is "dragbar." Drag the dragbar to move the note. Double-click the dragbar to open the editor. Right-click the note for the context menu.  The notes have optional borders to visually differentiate them.  
2.	 Alt+Click interactive checkbox items to toggle.
3.	Note Editor. Most of the formatting options are found here. Make checkbox items as seen in image.
4.	Alarm Settings gui.  After an alarm plays, it deletes itself unless it recurs on specified weekdays.   Choose from a folder of .wav files.  There is a ‘test’ button and a ‘stop’ button.  
5.	Sticky Note context menu.  When a note is alarmed, additional items (‘Stop’, ‘Delete Alarm’) are added to the menu.
6.	Sticky Notes Manager gui form. It is hidden by default. Default hotkey to show it is Win+Shift+S.  The buttons on the top are for mostly ‘app-level’ commands.
7.	The Filter Settings for the colorized ListView of notes.  Use edit box to filter list by note text.
8.	Select a note from the list and right-click for a 2 second pop up preview.  Note: Must let preview timeout before previewing another.  Left double-click to edit note.  Note: You might need to unhide a note before editing or deleting it.
9.	List row colors match note back and font colors. 
10.	The buttons on the bottom are for mostly ‘note-level’ commands.  You can Ctrl+Click to select multiple items.  Multi-selected items can be bulk hidden/unhidden or deleted. 
11.	System tray icon right click menu has a few optons such as ‘Start with Windows’ and ‘Open Note ini File.’
12.	The Tips dialog just shows a few tips.  It also lists the (customizable) hotkey combinations. 
