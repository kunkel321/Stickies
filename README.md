# Sticky Notes
(from code comments)
* Project:    Sticky Notes
* Author:     kunkel321
* Tool used:  Claude AI
* Version:    (Version date will be updated in code.)
* Forum:      https://www.autohotkey.com/boards/viewtopic.php?f=83&t=135340
* Repository: https://github.com/kunkel321/Stickies     

Customizable Hotkeys:
* Ctrl+Shift+N - Create new note
* Ctrl+Shift+C - Create new note from clipboard text
* Ctrl+Shift+S - Toggle main window visibility

Notes are created in the top/left of the screen, cascading down, so they don't hide each other.  The top portion of the note is the Drag Bar.  Drag from there to move the note around.   To edit a sticky note, double-click Drag Bar or right-click, then choose edit from context menu.  

Note will open in editor where user can access most of the note options.  After moving and editing notes, click "Save Status" on main window. Main window is hidden at start up, so toggle to view it.  Notes are saved to an ini file in same location as script file. The ini file "sticky_notes.ini" will be created in same location as script.

If a note's text has multiple lines, and any of those lines start with [] or [x] then those lines of text will be made into checkbox controls.  Accidental clicked are prevented with an optional mondifier key.  (Alt+Click by default).

If you change the length of text after the note has already been made, or if you check/uncheck boxes, I recommend doing "Save Status" from the note context menu or the main form.  This will save all notes and properties to the ini file.  Then do "Load/ReLoad Notes" if the note size hasn't adjusted correctly. 

The system tray context menu has a few extra items, such as "Start with Windows."

This script is unique because nearly every bit of code was created with AI prompts, then pasted into the ahk editor.  A great deal of human input was needed, but very little of the actual code was human-generated.  Edit: In later versions, more human code was added.
![Screenshot of Sticky Notes tool](https://i.imgur.com/j6Kyled.jpeg)
Screenshot explanation:
1. Sticky Note. Top is "DragBar." Drag dbar to move note. Double-click dbar to open editor. Right-click note for menu. Alt+Click checkbox items to toggle.
2. Sticky Note menu.
3. Note Editor. Most of the formatting options are found here. Make checkbox items as seen in image.
4. Sticky Notes Manager gui form. It is hidden by default. The hotkey/accelerator key text tips are dynaminically updated if hotkey/acc key are changed.
