window.enableMatrixDragSelection = ({selector, selected, keyFor, paintSelection}) => {
  const cells = [...document.querySelectorAll(selector)];
  let anchor = null;
  let baseSelection = new Set();
  let dragging = false;

  function selectRectangle(target) {
    if (!anchor) return;
    const firstRow = Math.min(anchor.parentElement.rowIndex, target.parentElement.rowIndex);
    const lastRow = Math.max(anchor.parentElement.rowIndex, target.parentElement.rowIndex);
    const firstColumn = Math.min(anchor.cellIndex, target.cellIndex);
    const lastColumn = Math.max(anchor.cellIndex, target.cellIndex);
    selected.clear();
    baseSelection.forEach(value => selected.add(value));
    cells.forEach(cell => {
      const row = cell.parentElement.rowIndex;
      if (row >= firstRow && row <= lastRow && cell.cellIndex >= firstColumn && cell.cellIndex <= lastColumn) {
        selected.add(keyFor(cell));
      }
    });
    paintSelection();
  }

  cells.forEach(cell => {
    cell.addEventListener("mousedown", event => {
      if (event.button !== 0 || window.matchMedia("(pointer: coarse)").matches) return;
      event.preventDefault();
      dragging = true;
      anchor = cell;
      baseSelection = event.ctrlKey || event.metaKey ? new Set(selected) : new Set();
      selectRectangle(cell);
    });
    cell.addEventListener("mouseenter", () => {
      if (dragging) selectRectangle(cell);
    });
    cell.addEventListener("click", event => {
      if (!window.matchMedia("(pointer: coarse)").matches) return;
      const key = keyFor(cell);
      selected.has(key) ? selected.delete(key) : selected.add(key);
      paintSelection();
    });
  });

  document.addEventListener("mouseup", event => {
    if (event.button === 0) {
      dragging = false;
      anchor = null;
    }
  });
};
