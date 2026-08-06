(() => {
  const planner = document.querySelector("[data-order-matrix]");
  if (!planner) return;
  const body = planner.querySelector(".user-order-body");
  const startButton = document.querySelector(".mobile-sort-start");
  const saveButton = document.querySelector(".mobile-sort-save");
  const mobile = window.matchMedia("(max-width: 700px)");
  let draggedRow = null;
  let changed = false;

  function rows() {
    return [...body.querySelectorAll(":scope > tr[data-user-id]")];
  }

  function moveAtPointer(clientY) {
    const target = document.elementFromPoint(20, clientY)?.closest("tr[data-user-id]");
    if (!target || target === draggedRow || target.parentElement !== body) return;
    const before = clientY < target.getBoundingClientRect().top + target.offsetHeight / 2;
    body.insertBefore(draggedRow, before ? target : target.nextSibling);
    changed = true;
  }

  async function persist() {
    const response = await fetch(planner.dataset.orderUrl, {
      method: "POST",
      headers: {"Content-Type": "application/json"},
      body: JSON.stringify({user_ids: rows().map(row => Number(row.dataset.userId))}),
    });
    if (!response.ok) throw new Error("Reihenfolge konnte nicht gespeichert werden.");
    changed = false;
  }

  body.querySelectorAll(".reorder-handle").forEach(handle => {
    const row = handle.closest("tr[data-user-id]");
    handle.addEventListener("dragstart", event => {
      if (mobile.matches) { event.preventDefault(); return; }
      draggedRow = row;
      row.classList.add("row-dragging");
      event.dataTransfer.effectAllowed = "move";
      event.dataTransfer.setData("text/plain", row.dataset.userId);
    });
    handle.addEventListener("dragend", async () => {
      row.classList.remove("row-dragging");
      draggedRow = null;
      if (changed) await persist();
    });
    handle.addEventListener("pointerdown", event => {
      if (!mobile.matches || !planner.classList.contains("sort-mode")) return;
      event.preventDefault();
      draggedRow = row;
      row.classList.add("row-dragging");
      handle.setPointerCapture(event.pointerId);
    });
    handle.addEventListener("pointermove", event => {
      if (draggedRow === row && mobile.matches) moveAtPointer(event.clientY);
    });
    handle.addEventListener("pointerup", () => {
      row.classList.remove("row-dragging");
      draggedRow = null;
    });
  });

  body.addEventListener("dragover", event => {
    if (!draggedRow || mobile.matches) return;
    event.preventDefault();
    moveAtPointer(event.clientY);
  });

  startButton?.addEventListener("click", () => {
    planner.classList.add("sort-mode");
    startButton.hidden = true;
    saveButton.hidden = false;
  });
  saveButton?.addEventListener("click", async () => {
    await persist();
    planner.classList.remove("sort-mode");
    saveButton.hidden = true;
    startButton.hidden = false;
  });
})();
