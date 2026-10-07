// Typing into a selected cell.
//
// Univer keeps a hidden, empty contenteditable focused while a cell is
// selected and takes the first keystroke from its `beforeinput`. That editor
// sits in a 1×1 px clipping box and has no content, so its height can
// collapse to 0. Inside Haven's frame it does (in GenOffice's own window it
// happens not to), and Chrome then treats it as having no place to insert
// text: no `beforeinput`, and typing into a selected cell does nothing while
// the formula bar still works. A minimum height gives the editor a line box
// again; it stays invisible and clipped.
export function installEditorInputFix(): void {
  const style = document.createElement('style')
  style.textContent = '[data-u-comp="editor"] { min-height: 1em; min-width: 1px; }'
  document.head.append(style)
}
