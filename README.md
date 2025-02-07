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
* Ctrl+Shift+N - Create new note
* Ctrl+Shift+C - Create new note from clipboard text
* Ctrl+Shift+S - Toggle main window visibility

Features and Usage:
------------------
- Create sticky notes that persist between script restarts
- Notes cascade from top-left of screen for better visibility
- Notes cycle through different pastel background colors automatically
- Convert text lines starting with [] or [x] into interactive checkboxes
- Set alarms for individual notes with optional weekly recurrence
- Rich formatting options including fonts, colors, and sizing
- Main window with note management and search functionality

Core Functionality:
------------------
- Notes are created using Ctrl+Shift+N or from clipboard with Ctrl+Shift+S
- Drag notes by their top bar to reposition
- Double-click top bar or right-click for editing options
- Notes auto-save position when moved
- Access main window with Ctrl+Shift+S (hidden by default)
- System tray icon provides quick access to common functions

Special Features:
----------------
- Checkbox Creation: Any line starting with [] or [x] becomes a checkbox
- Checkbox Safety: Alt+Click required by default to prevent accidental toggles
- Alarm System: Set one-time or recurring alarms with custom sounds
- Note Search: Filter notes by content in main window
- Color Cycling: Each new note gets next color in palette
- Cascading Positions: New note doesn't fully cover previous note
- Start with Windows: Option available in tray menu

Data Management:
---------------
- All note data saved to sticky_notes.ini in script directory
- Notes auto-save on most changes
- Manual save available via "Save Status" button
- "Load/Reload Notes" refreshes all notes from storage
- Hidden notes can be restored through main window

Tips:
-----
- Use "Save Status" after significant changes
- "Load/Reload Notes" helps if note display issues occur
- Check error_log.txt for troubleshooting (if enabled)
- Use main window's search to find specific notes
- Right-click menu provides quick access to common functions

Development Note:
----------------
This script was developed primarily through AI-assisted coding, with Claude AI generating most of the base code structure and functionality. Later versions 
include additional human-written code for enhanced features and bug fixes.. The system tray context menu has a few extra items, such as "Start with Windows."

This script is unique because nearly every bit of code was created with AI prompts, then pasted into the ahk editor.  A great deal of human input was needed, but very little of the actual code was human-generated.  Edit: In later versions, more human code was added.
![Screenshot of Sticky Notes tool](https://i.imgur.com/6AdAEtZ.png)
Screenshot explanation:
1.	Sticky Note. Top/Center is "DragBar." Drag dragbar to move note. Double-click dragbar to open editor. Right-click note for menu. Alt+Click checkbox items to toggle.
2.	Note Editor. Most of the formatting options are found here. Make checkbox items as seen in image.
3.	New ‘Add Alarm...’ button.  
4.	New Alarm Settings gui.  After an alarm plays, it deletes itself unless it recurs on specified weekdays.  
5.	Sticky Note context menu.
6.	Sticky Notes Manager gui form. It is hidden by default. 
7.	The buttons on the top are for mostly ‘app-level’ commands.
8.	The Filter Settings for the new colorized ListView of notes.
9.	The buttons on the bottom are for mostly ‘note-level’ commands.
