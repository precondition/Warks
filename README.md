# Warks
## Quick marks add-in for Microsoft Word for faster navigation

Warks makes moving around long Word documents feel instant. Set a mark, jump to it, jump back, and even bounce across documents, all without inserting anything into the file. Marks persist across sessions, work on read-only docs, and are designed for keyboard-centric reading workflows. Inspired by Vim marks, Warks brings that speed to Word.



https://github.com/user-attachments/assets/70c6e6f6-ecbb-4079-b9f1-404b6c899ee2



**Why Warks instead of native Word Bookmarks?**

- Works in read-only and leaves no trace
  - Bookmarks modify the document and cannot be added from “Viewing” mode or to restricted read-only documents.
  - Warks keeps your documents pristine. Set and jump to marks even when you can’t edit the file.
- Works in all three viewing layouts (Read Mode, Print Layout, Web Layout)
    - It is impossible to interact with bookmarks in read mode. There is no
      ribbon and you can try <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd> and
      <kbd>Ctrl</kbd>+<kbd>G</kbd> all you want, nothing will happen.
    - You can set, list, and jump to marks using keyboard shortcuts even in read mode.
- Global, cross-document navigation
  - Bookmarks are scoped to a single file.
  - Warks supports local marks and global marks (start with an uppercase letter) that jump between documents and open the target file if needed. So you can quickly jump to that one section from the specs you keep referring back to, regardless of what is the active document.
- Throwaway agility
  - Bookmarks tend to be permanent artifacts you must clean up.
  - Warks encourages quick, disposable waypoints (that still persist for you) without touching the document.

### Illustrative Use Case

User story: an engineer cross-checks a diagram and jumps to a glossary, fast and without losing his place.

Alex is a systems engineer reviewing a 170‑page technical spec in Word’s Read Mode. The document has several figures, dense tables, and long explanatory sections. While reading, Alex often needs to cross‑check the main architecture diagram. Crucially, the place from which he jumps changes every time since he’s naturally reading forward, so he must be able to return to the exact spot he just left. Later, he hits an unfamiliar acronym and wants to look it up in a “Glossary.docx” that isn’t currently open.

How this plays out with native Word Bookmarks

Goal 1: Jump to the main diagram, then back to the exact reading spot (which keeps changing)
- Every time before jumping, Alex must create or update a bookmark at his current reading position. If he forgets, he has no easy “jump back” path.
- In Read Mode, setting bookmarks is not supported, so he first switches to Print Layout:
  - <kbd>Alt</kbd>, W, E (View > Edit Document)
- Set a bookmark at the current reading position:
  - <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd> (Bookmark dialog)
  - Type READ (or a unique name, e.g., read1, read2…)
  - Important gotcha: typing an existing bookmark name and pressing <kbd>Enter</kbd> will jump to that bookmark (default action), not add a new one. To add here, he must press:
    - <kbd>Alt</kbd>+<kbd>A</kbd> (Add)
  - <kbd>Esc</kbd> or <kbd>Enter</kbd> (close)
- Jump to the main diagram:
  - <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd>
  - Arrow keys (or type) to MAIN_DIAGRAM
  - <kbd>Alt</kbd>+<kbd>G</kbd> (Go To)
  - <kbd>Esc</kbd> or <kbd>Enter</kbd>
- Return to the exact reading spot:
  - <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd>
  - Arrow keys (or type) to READ (or readN if he uses throwaway names)
  - <kbd>Alt</kbd>+<kbd>G</kbd>
  - <kbd>Esc</kbd> or <kbd>Enter</kbd>
- Go back to Read Mode:
  - <kbd>Alt</kbd>, <kbd>W</kbd>, <kbd>F</kbd> (View > Read Mode)

Notes and friction:
- If READ already exists and he reflexively types READ + <kbd>Enter</kbd>, Word will jump to the old READ instead of adding one at his current spot. He must either invent new names each time (read1, read2, read3, …) or be careful to press <kbd>Alt</kbd>+<kbd>A</kbd> to Add before jumping.
- In read‑only docs, he cannot add bookmarks at all.
- Bookmarks remain in the file as artifacts. Without cleanup, the file will grow increasingly more bloated.

Goal 2: Look up an unknown abbreviation in an external glossary (not open)
- Open the glossary:
  - <kbd>Ctrl</kbd>+<kbd>O</kbd> (Open)
  - <kbd>Tab</kbd> x 4 times to enter the Favorites section
  - Select Glossary.docx with arrow keys
  - <kbd>Enter</kbd> (open)
- Find the acronym:
  - <kbd>Ctrl</kbd>+<kbd>F,</kbd> type the term, <kbd>Enter</kbd>
- Return to the spec and exact reading spot:
  - <kbd>Alt</kbd>+<kbd>Tab</kbd> or <kbd>Ctrl</kbd>+<kbd>F6</kbd> (switch back to the spec)

Total feel: frequent mode/layout switches, repeated dialog interactions, and the cognitive overhead of “did I remember to Add a fresh READ before jumping?”. If he doesn’t, <kbd>Enter</kbd> will take him to the previous READ location, not where he just was.

How the same flow feels with Warks

Assumptions:
- Alex bound MarkSet to <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>'</kbd> and MarkJump to <kbd>Ctrl</kbd>+<kbd>'</kbd>.
- He bound MarkJumpBack to <kbd>Alt</kbd>+<kbd>Backspace</kbd> (apostrophe).
- He set a one‑time local mark “d” at the main diagram and a one‑time global mark “G” in Glossary.docx. Marks persist for Alex but don’t touch documents, and they work in Read Mode and read‑only.

Goal 1: Cross‑check the main diagram, then get back to the evolving reading position
- One‑time prep (when he first sees the diagram):
  - <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>'</kbd>, type d, <kbd>Enter</kbd> (sets “d” at the diagram)
- While reading (the start point changes every time):
  - Jump to the diagram:
    - <kbd>Ctrl</kbd>+<kbd>'</kbd>, type d, <kbd>Enter</kbd>
  - Jump back to the exact place he left:
    - <kbd>Alt</kbd>+<kbd>Backspace</kbd> (MarkJumpBack)
- He doesn’t need to pre‑mark his reading spot each time. Warks remembers the previous location automatically. No layout switching. Works even if the doc is read‑only.

Goal 2: Look up an acronym in a glossary that isn’t open
- Jump straight to the glossary and the “G” mark:
  - <kbd>Ctrl</kbd>+<kbd>'</kbd>, type G, <kbd>Enter</kbd> (uppercase = global; Warks opens the glossary if needed)
- Find the term:
  - <kbd>Ctrl</kbd>+<kbd>F,</kbd> type the acronym, <kbd>Enter</kbd>
- Return to the spec:
  - <kbd>Alt</kbd>+<kbd>Tab</kbd>

Optional no‑prompt variant for Alex’s two most‑used targets
- Bind MarkSetLocalA/MarkJumpLocalA for the main diagram mark and MarkSetGlobalA/MarkJumpGlobalA for the glossary entry point.
- Then it’s one key to set and one key to jump; press apostrophe to go back. Zero prompts.

Why Warks wins here
- It eliminates the “don’t forget to add a new READ before you jump” trap. Back‑jump just works.
- It keeps Alex in flow with minimal prompts and pure keyboard movement.
- It works in Read Mode and read‑only documents because it doesn’t modify the file.
- Global marks mean cross‑document jumps are one shortcut away, even if the target document isn’t open.

Quick sequence comparison

Cross‑check diagram and return:
- Word Bookmarks:
  - <kbd>Alt</kbd>, W, E → <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd> → type READ → <kbd>Alt</kbd>+<kbd>A</kbd> → <kbd>Esc</kbd> → <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd> → select MAIN_DIAGRAM → <kbd>Alt</kbd>+<kbd>G</kbd> → <kbd>Esc</kbd> → <kbd>Ctrl</kbd>+<kbd>Shift</kbd>+<kbd>F5</kbd> → select READ → <kbd>Alt</kbd>+<kbd>G</kbd> → <kbd>Esc</kbd> → <kbd>Alt</kbd> W, F
  - And if he typed READ + <kbd>Enter</kbd> by habit, he’ll jump to the old READ and lose his current spot unless he first created a new unique READ name.
- Warks:
  - <kbd>Ctrl</kbd>+<kbd>'</kbd>, d, <kbd>Enter</kbd> → <kbd>Alt</kbd>+<kbd>Backspace</kbd>

Jump to external glossary (not open) and return:
- Word Bookmarks:
  - <kbd>Ctrl</kbd>+<kbd>O</kbd> → navigate to Glossary.docx → <kbd>Enter</kbd> → <kbd>Ctrl</kbd>+<kbd>F</kbd> → type acronym → <kbd>Enter</kbd> → <kbd>Alt</kbd>+<kbd>Tab</kbd>/<kbd>Ctrl</kbd>+<kbd>F6</kbd>
- Warks:
  - <kbd>Ctrl</kbd>+<kbd>'</kbd>, G, <kbd>Enter</kbd> → <kbd>Ctrl</kbd>+<kbd>F</kbd> → type acronym → <kbd>Enter</kbd> → <kbd>Alt</kbd>+<kbd>Tab</kbd>

#### Bottom line
Bookmarks make Alex manage dialogs, layout changes, and the naming/adding dance before every jump. Warks makes the same workflow a couple of keystrokes: jump, back. Clean, persistent, and fast, within and across documents.

Vim marks inspiration
- Uppercase letters (A–Z) are global marks that can jump across documents.
- Lowercase letters (a–z) are local marks bound to the current document.
- The apostrophe ' is reserved for “jump back” so you can toggle between contexts.
- MarkJump goes to the exact character; MarkJumpLine goes to the start of the marked line.

### System requirements
- Microsoft Word for Windows with VBA (Word 2013 or later recommended)
- Windows only. Warks uses a Windows-specific persistence API; macOS is not supported.

### Installation

Option A: Add to Normal template (loads for all documents)
1) Open Word.
2) Press <kbd>Alt</kbd>+<kbd>F11</kbd> to open the VBA editor.
3) In the Project pane, expand Normal (Normal.dotm).
4) Right-click Normal > Insert > Module.
5) Paste the Warks code into the new module.
6) Close the VBA editor.
7) In Word, visit File > Options > Trust Center > Trust Center Settings… > Macro Settings and allow macros (e.g., “Disable all macros with notification” so you can enable them per document).
8) Save Normal.dotm if prompted.

Option B: Install as a Startup add-in (.dotm)
1) Create a new document in Word.
2) Press <kbd>Alt</kbd>+<kbd>F11,</kbd> insert a Module, and paste the Warks code.
3) Save as a Word Macro-Enabled Template (`*.dotm`), e.g., Warks.dotm.
4) Place Warks.dotm in your Word Startup folder:
   - Typically at %AppData%\Microsoft\Word\STARTUP
   - To confirm path: File > Options > Advanced > General > File Locations… > Startup.
5) Restart Word.

Quick start: make it fast
- Assign keyboard shortcuts:
  - File > Options > Customize Ribbon > Keyboard shortcuts: Customize…
  - Categories: Macros; then bind the macros below.
- Suggested bindings:
  - MarkSet: <kbd>Ctrl</kbd>+<kbd>Alt</kbd>+<kbd>'</kbd>
  - MarkJump: <kbd>Ctrl</kbd>+<kbd>'</kbd>
  - MarkJumpBack <kbd>Alt</kbd>+<kbd>Backspace</kbd>

### Available macros
- Interactive (prompt for a name):
  - MarkSet: set a mark to the current cursor position
  - MarkJump: jump to a named mark
  - MarkJumpLine: jump to the start of the line containing the mark
  - MarkList: list local marks for the current document plus global marks, with text previews
- No-prompt convenience:
  - MarkSetGlobalA / MarkJumpGlobalA
  - MarkSetLocalA  / MarkJumpLocalA
  - MarkSetGlobalB / MarkJumpGlobalB
  - MarkSetLocalB  / MarkJumpLocalB
  - MarkJumpBack: jump to the previous location (the `'` mark)
  - MarkJumpLineBack: previous location at line start

### Known limitations
- Local marks are tied to the document’s full path. If a file is moved/renamed, previously set local marks won’t show for the new path.
- Marks are anchored to character positions; edits before a mark can shift its intended location. Use MarkList previews to confirm targets.
- Due to limitations in the Word API, Warks must temporarily switch back to Print Layout in order to scroll the mark into view in Read Mode so there may be undesirable visual flashing.
- Windows only.

### Troubleshooting
- Macros don’t appear: Bind and use the 0-argument public macros listed above; Word’s Macros UI shows only those.
- Macros won’t run: Trust Center > Macro Settings must allow macros (e.g., “Disable with notification,” then enable per document).

### Uninstallation

Step 1: Remove the VBA module or add-in
- If installed in Normal.dotm:
  1) Press <kbd>Alt</kbd>+<kbd>F11</kbd> to open the VBA editor.
  2) Under Normal (Normal.dotm), locate the Warks module.
  3) Right-click > Remove. Choose No when asked to export unless you want a backup.
  4) Close the editor, exit Word, and save Normal.dotm if prompted.

- If installed as a Startup add-in:
  1) Close Word.
  2) Open your Startup folder (see Installation Option B).
  3) Delete or move out Warks.dotm.
  4) Restart Word.

Step 2: Optional cleanup to remove persisted marks
- Caution: Editing the Windows Registry affects only your own marks here, but back up the key if you want a restore point.
1) Press <kbd>Win</kbd>+<kbd>R</kbd>, type `regedit`, press <kbd>Enter</kbd>.
2) Navigate to:
   HKEY_CURRENT_USER\Software\VB and VBA Program Settings\Warks
3) Under Warks, select Marks. You can:
   - Delete just Marks (right-click Marks > Delete) to remove all saved marks.
   - Or delete the entire Warks key to remove everything Warks stored.
4) Inside Marks are values named by full document paths (local marks) and a special value named </GLOBAL\> (global marks). Deleting Marks or Warks removes them all.
5) Close Registry Editor.
