# Sticky Notes
(from code comments)
* Project:    Sticky Notes
* Author:     kunkel321
* Tool used:  Claude AI
* Version:    1-14-2025
* Forum:      https://www.autohotkey.com/boards/viewtopic.php?f=83&t=135340
* Repository: https://github.com/kunkel321/Stickies     

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
![Screenshot of Sticky Notes tool](https://github.com/kunkel321/Stickies/blob/main/sticky_note_screenshot.PNG?raw=true)
