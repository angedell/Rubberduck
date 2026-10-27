# Git Sync: comparing and merging a module

When a module and its file differ and neither side is simply the newer one, the Git Sync review lets you look at both and build the result yourself, one block at a time. This works like the merge view of [WinMerge](https://winmerge.org/).

## Opening the compare window

1. Run **Sync Project > Git Sync...** and wait for the review.
2. Select **one** item that exists on both sides (status *Changed in both*, *Different, no history*, or any other that is not *Only in project/folder*).
3. Press **Compare...**, or double-click the row.

Items that exist on one side only have nothing to compare, so the button does nothing for them.

## The three panes

| Pane | What it shows |
| --- | --- |
| **Project** (left) | The module as the VBA editor exports it. Read-only. |
| **Result** (middle) | What the module and the file will both contain after **Sync selected**. You can type in it. |
| **Folder** (right) | The file from the sync folder. Read-only. |

Project and Folder are lined up row by row; where one side has fewer lines, the other side gets blank (grey) rows. The three panes scroll together. Result may have a different number of lines once you edit it, so it follows block by block rather than row by row.

### Colors

- **Black on white** - the same on all three. Nothing to decide.
- **Yellow, brown text** - a block that differs and is **not decided yet**. In Result it shows as a reminder line, `<<< click the project's or the folder's version >>>`.
- **Green** - a block that is in the Result. Shown green on the side(s) it came from, and green in Result.
- **Grey** - the side you did **not** take for a block.
- **Blue** - text you wrote yourself in Result.
- A stronger color inside a changed line marks the **words** that differ, as in WinMerge.
- The block you are on (the *current* one) is a little more orange.

Indentation and trailing spaces are not counted as differences.

## Choosing, block by block

Move the mouse over a differing block in **Project** or **Folder**: the pointer turns into a **hand**. Then:

- **Click** - put that side's lines in Result.
- **Click the same block again** - take it back out (the block is undecided again).
- **Click the other side** - swap to that side.
- **Ctrl+click** - take **both**: the project's lines first, then the folder's.

Clicking a part where the other side has no lines (a grey blank block) takes that side's *nothing*: the lines are left out of Result.

Selecting text with the mouse to copy it does not count as a click.

### Buttons

- **Previous / Next difference** - jump to the neighbouring block.
- **Use project / Use folder / Use both** - decide the current block, then move on to the next undecided one.
- **All from project / All from folder** - take one whole side. Use these first and then cherry pick the exceptions, or the other way round.
- **Merge** - accept the Result. If some blocks are still undecided it shows a message and takes you to the first one.
- **Cancel** - leave without changing anything.

If the history says which side is newer, the blocks start on that side. If it cannot say, they all start undecided.

## Editing the result

Click in the middle pane and type. Whatever is in Result when you press **Merge** is what is used.

- A block you edit turns **blue**. If you then click Project or Folder on that block, you are asked first, because it replaces your edit.
- Typing in a block that is still undecided removes the reminder line.
- **All from project/folder** also asks if you have edits, because it rebuilds the whole Result.
- Undo (Ctrl+Z) works as in any text box.

## After you press Merge

The item in the review changes to **Merge both sides**, ticked. Nothing is written until **Sync selected**. Then:

1. the old module is saved to `%LocalAppData%\Rubberduck\GitSyncBackups\<time>`, and the old file to its `folder` subfolder;
2. the Result is imported into the project through the VBA editor (so hidden attributes survive) and written to the file, so both are the same;
3. the history in `rubberduck.sync.json` is updated.

The `rubberduck.json` rules still apply: with `"onViolation": "fix"` they are repaired, with `"block"` a Result that breaks one cannot be ticked until you fix it in the Result.

To start over, open **Compare...** again, or choose **Use project version / Use folder version** in the review, which replaces the merge.

## Tips

- Resolve the easy parts with **All from ...** first, then switch single blocks with a click.
- **Ctrl+click** is the quick way to keep two versions of a procedure and sort them out by hand in Result.
- Large files: the comparison uses the same algorithm as WinMerge (diffutils), which stays fast on thousands of lines.

## Credits

The comparison is a C# port of parts of WinMerge (diffutils `analyze.c` and `stringdiffs.cpp`), licensed GPL-2.0-or-later. See the headers of `WinMergeLineDiff.cs` and `WinMergeWordDiff.cs`.
