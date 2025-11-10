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
- Alarm System: Set one-time or recurring alarms with custom sounds
- Set alarms for individual notes with optional weekly recurrence
- Many editor dialog options support accelerator keys (Alt+A for alarm, etc.)
- Visual alert: Notes can shake when alarms trigger
- Multiple alarm repeats: Choose between once, 3x, or 10x alarm repetitions
- Alarm sounds: Custom alarm sounds can be added to the Sounds folder
- Smart alarm management: System detects and reports missed alarms on startup
- An alarm can have a: date and/or time and/or recurrence 
- The logic for alarms is this:
- > Date + time + no weekly recurrence: plays once on given date, then deletes itself
- > Date + no time + no weekly recurrence: note appears in morning on date with no sound/shake, then alarm deletes
- > Date + time + weekly recurrence: plays on given date, doesn't delete itself, plays again on recurring days
- > Date + no time + weekly recurrence: note appears in morning on date, doesn't delete, reappears on recurring days
- > No date + no time + weekly recurrence: note appears in morning on recurring weekdays
- > No date + time + no weekly recurrence: plays at specified time
- > No date + time + weekly recurrence: plays at specified time on recurring days
- Main window with note management, preview, and search functionality
- Color-coded notes: Notes maintain their color scheme in preview list
- Rich preview: Right-click notes in manager for formatted preview with original fonts/colors
- Search functionality: Filter notes by content in main window
- Selective display: Filter note list to show only hidden or visible notes
- Include recently deleted note in hidden/visible listview filter
- Deleted notes are identified by deletion time appearing in listview
- Access main window with Win+Shift+S (hidden by default)
- Resize main window, by dragging edge/corner, to see more of note text in listview
- Notes' show alarm times and window attachments in note manager listview
- Multiple selection: Use Ctrl+Click to select multiple notes in manager
- Bulk operations: Select multiple notes to hide/unhide/delete/undelete simultaneously
- Notes are created using Win+Shift+N or from clipboard with Win+Shift+C
- Hotkeys and more can be changed near top of code
- Tips button in note manager shows current hotkeys and other tips
- Double-click top bar or right-click for editing options
- Drag notes by their top bar to reposition
- Configurable drag area: Option to maximize note space by minimizing drag area
- Notes auto-save position when moved
- Turn off note deletion warning
- Undelete notes
- Deleted notes are purged from ini file after 3 days (Configurable)
- Checkbox Creation: Any text line starting with [] or [x] becomes an interactive checkbox
- Checkbox Safety: Alt+Click required by default to prevent accidental toggles
- Hidden or deleted notes can be restored through main window or via note context menu
- All note data saved to sticky_notes.ini in script directory
- Check error_debug_log.txt for troubleshooting (if enabled; warning: system hog)
- Use manual "Save Status" after significant changes
- "Load/Reload Notes" refreshes all notes from storage
- System tray icon provides quick access to common functions
- Start with Windows: Option available in tray menu

Known Issue:
------------
- In note listview, if notes are sorted, colors will not sort with them--that is why sorting is disabled.

Development Note:
-----------------
This script was developed primarily through AI-assisted coding, with Claude AI generating most of the base code structure and functionality. Later versions 
include additional human-written code for enhanced features and bug fixes.. The system tray context menu has a few extra items, such as "Start with Windows."

This script is unique because nearly every bit of code was created with AI prompts, then pasted into the ahk editor.  A great deal of human input was needed, but very little of the actual code was human-generated.  Edit: In later versions, more human code was added.
![Screenshot of Sticky Notes tool](https://i.imgur.com/QtzJshJ.png)
Screenshot explanation:
1.	Sticky Note. The top/center of each note is a "drag area." Drag the drag area to move the note. Double-click the drag area to open the editor. Right-click the note for the context menu.  Size of drag area can be customized.  The top area of the note can be reserved. (It is not reserved in the image.)  The notes have optional borders to visually differentiate them.  Border color matches font color.
2.	 Alt+Click interactive checkbox items to toggle.  Checkboxes are typed into the note text by starting a new line with [] or [x].  Checkbox text doesn’t wrap.  Long items will override the note width setting. 
3.	Note Editor. Most of the formatting options are found here. Make checkbox items as seen in image.  The top group of dialog controls are for formatting.  Default colors of new notes cycle with each new note, but you can manually select colors or click “random” for random colors.  The bottom group of controls are note options.  The “Add Alarm” button opens the Set Alarm dialog seen in the image.  The “Stick to Window” button opens a dialog with a list of the current visible windows on your computer (not seen in image).  Stuck windows auto show/hide with their associated window.   The Width box is for the width of the note.  The editor will mostly change width too.  
4.	Alarm Settings Dialog.  Single occurrence alarms delete themselves from the note after playing.  An alarm can have a date and/or time.  Timed alarms can have custom sounds.  Choose from a folder of .wav files.  There is a sound ‘test’ button and a ‘stop’ button.  Timed alarms can also “Shake.”  The duration and amount of shake are customizable.  Untimed dated or recurring alarms silently unhide themselves the morning of the associated date or weekday. 
5.	Alarm date picker dialog. 
6.	Sticky Note context menu.  When a note is alarmed, additional items (‘Stop’, ‘Delete Alarm’) are added to the menu.  The “Show Hidden” and the “Undelete a note” items popup a submenu of the hidden or deleted items.  Deleted items get purged from the ini settings file after X days.
7.	Submenu showing recently deleted items.  Choose one to undelete, or choose multiple from the Note Manager list. 
8.	Sticky Notes Manager gui form. It is hidden by default. Default hotkey to show it is Win+Shift+S.  
9.	The buttons on the top are for mostly ‘app-level’ commands.
10.	The Filter Settings for the colorized ListView of notes.  Deleted notes are hidden by default.  Use edit box to filter list by note text.
11.	Select a note from the list and right-click for a two-second pop up preview.  List row colors match note back and font colors.  Left double-click to toggle visibility of note.  Note: You might need to unhide a note before editing or deleting it.  Drag edge of Note Manager window to resize the list and show more notes.  List is sorted by first column which is creation date.  It is really small to make more room for column two.  Column two shows note text.  It any deleted notes are shown, then column two will have Deletion date/time stamp preceding the note text (not seen in image).  Column three has the Alarm date and/or time, and lists the stuck-to window.   
12.	The buttons on the bottom are for mostly ‘note-level’ commands.  You can Ctrl+Click to select multiple items.  Multi-selected items can be bulk hidden/unhidden or deleted. 
13.	The Tips dialog just shows a few tips.  It also lists the current (customizable) hotkey combinations.
14.	System tray icon right click menu has a few options such as ‘Start with Windows’ and ‘Open Note ini File.’
15.	This is just my Desktop...  I added that marker accidentally!
