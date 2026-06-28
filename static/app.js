// ===========================================================
// MOI REGISTER — app logic
// ===========================================================

const API = "/api/entries";

const el = {
  entryList: document.getElementById("entryList"),
  emptyState: document.getElementById("emptyState"),
  summaryTotal: document.getElementById("summaryTotal"),
  summaryCount: document.getElementById("summaryCount"),
  searchInput: document.getElementById("searchInput"),
  sortSelect: document.getElementById("sortSelect"),
  addEntryBtn: document.getElementById("addEntryBtn"),
  sheetOverlay: document.getElementById("sheetOverlay"),
  sheet: document.getElementById("sheet"),
  sheetTitle: document.getElementById("sheetTitle"),
  sheetCloseBtn: document.getElementById("sheetCloseBtn"),
  cancelBtn: document.getElementById("cancelBtn"),
  entryForm: document.getElementById("entryForm"),
  entryId: document.getElementById("entryId"),
  guestName: document.getElementById("guestName"),
  amount: document.getElementById("amount"),
  entryDate: document.getElementById("entryDate"),
  relation: document.getElementById("relation"),
  eventName: document.getElementById("eventName"),
  notes: document.getElementById("notes"),
  confirmOverlay: document.getElementById("confirmOverlay"),
  confirmSub: document.getElementById("confirmSub"),
  confirmCancelBtn: document.getElementById("confirmCancelBtn"),
  confirmDeleteBtn: document.getElementById("confirmDeleteBtn"),
  toast: document.getElementById("toast"),
  exportExcelBtn: document.getElementById("exportExcelBtn"),
  exportPdfBtn: document.getElementById("exportPdfBtn"),
};

let pendingDeleteId = null;
let searchDebounce = null;

// ----------------------------------------------------------------
// Helpers
// ----------------------------------------------------------------
function formatCurrency(amount) {
  return "₹" + Number(amount).toLocaleString("en-IN", {
    maximumFractionDigits: 0,
  });
}

function formatDate(dateStr) {
  const d = new Date(dateStr + "T00:00:00");
  if (isNaN(d.getTime())) return dateStr;
  return d.toLocaleDateString("en-IN", { day: "numeric", month: "short", year: "numeric" });
}

function showToast(message) {
  el.toast.textContent = message;
  el.toast.classList.add("show");
  clearTimeout(showToast._t);
  showToast._t = setTimeout(() => el.toast.classList.remove("show"), 2200);
}

function escapeHtml(str) {
  const div = document.createElement("div");
  div.textContent = str;
  return div.innerHTML;
}

// ----------------------------------------------------------------
// Fetch + render entries
// ----------------------------------------------------------------
async function loadEntries() {
  const search = el.searchInput.value.trim();
  const sort = el.sortSelect.value;
  const params = new URLSearchParams();
  if (search) params.set("search", search);
  params.set("sort", sort);

  try {
    const res = await fetch(`${API}?${params.toString()}`);
    if (!res.ok) throw new Error("Failed to load entries");
    const data = await res.json();
    renderEntries(data.entries);
    el.summaryTotal.textContent = formatCurrency(data.total);
    el.summaryCount.textContent = data.count;
  } catch (err) {
    showToast("Could not load entries. Please try again.");
    console.error(err);
  }
}

function renderEntries(entries) {
  el.entryList.innerHTML = "";

  if (!entries.length) {
    el.emptyState.hidden = false;
    return;
  }
  el.emptyState.hidden = true;

  const frag = document.createDocumentFragment();
  for (const entry of entries) {
    frag.appendChild(buildEntryCard(entry));
  }
  el.entryList.appendChild(frag);
}

function buildEntryCard(entry) {
  const li = document.createElement("li");
  li.className = "entry-card";
  li.dataset.id = entry.id;

  const tagsHtml = [
    entry.relation ? `<span class="tag tag-relation">${escapeHtml(entry.relation)}</span>` : "",
    entry.event_name ? `<span class="tag tag-event">${escapeHtml(entry.event_name)}</span>` : "",
    `<span class="tag tag-date">${formatDate(entry.entry_date)}</span>`,
  ].join("");

  li.innerHTML = `
    <div class="entry-stub"></div>
    <div class="entry-body">
      <div class="entry-top-row">
        <p class="entry-name">${escapeHtml(entry.guest_name)}</p>
        <span class="entry-amount">${formatCurrency(entry.amount)}</span>
      </div>
      <div class="entry-meta">${tagsHtml}</div>
      ${entry.notes ? `<p class="entry-notes">${escapeHtml(entry.notes)}</p>` : ""}
    </div>
    <div class="entry-actions">
      <button class="icon-btn edit-btn" aria-label="Edit entry" data-action="edit">
        <svg width="15" height="15" viewBox="0 0 16 16" fill="none"><path d="M11.5 2.5l2 2-8 8-2.5.5.5-2.5 8-8z" stroke="currentColor" stroke-width="1.2" stroke-linejoin="round"/></svg>
      </button>
      <button class="icon-btn delete-btn" aria-label="Delete entry" data-action="delete">
        <svg width="15" height="15" viewBox="0 0 16 16" fill="none"><path d="M3 4.5h10M6 4.5V3h4v1.5M4.5 4.5l.5 9h6l.5-9" stroke="currentColor" stroke-width="1.2" stroke-linecap="round" stroke-linejoin="round"/></svg>
      </button>
    </div>
  `;

  li.querySelector('[data-action="edit"]').addEventListener("click", () => openEditSheet(entry));
  li.querySelector('[data-action="delete"]').addEventListener("click", () => openDeleteConfirm(entry));

  return li;
}

// ----------------------------------------------------------------
// Sheet (add/edit form)
// ----------------------------------------------------------------
function openAddSheet() {
  el.sheetTitle.textContent = "New Entry";
  el.entryForm.reset();
  el.entryId.value = "";
  el.entryDate.value = new Date().toISOString().slice(0, 10);
  openSheet();
}

function openEditSheet(entry) {
  el.sheetTitle.textContent = "Edit Entry";
  el.entryId.value = entry.id;
  el.guestName.value = entry.guest_name;
  el.amount.value = entry.amount;
  el.entryDate.value = entry.entry_date;
  el.relation.value = entry.relation || "";
  el.eventName.value = entry.event_name || "";
  el.notes.value = entry.notes || "";
  openSheet();
}

function openSheet() {
  el.sheetOverlay.classList.add("open");
  document.body.style.overflow = "hidden";
  setTimeout(() => el.guestName.focus(), 250);
}

function closeSheet() {
  el.sheetOverlay.classList.remove("open");
  document.body.style.overflow = "";
}

async function handleFormSubmit(e) {
  e.preventDefault();

  const payload = {
    guest_name: el.guestName.value.trim(),
    amount: el.amount.value,
    entry_date: el.entryDate.value,
    relation: el.relation.value.trim(),
    event_name: el.eventName.value.trim(),
    notes: el.notes.value.trim(),
  };

  if (!payload.guest_name) {
    showToast("Please enter the guest's name.");
    return;
  }
  if (payload.amount === "" || isNaN(parseFloat(payload.amount)) || parseFloat(payload.amount) < 0) {
    showToast("Please enter a valid amount.");
    return;
  }

  const id = el.entryId.value;
  const isEdit = Boolean(id);
  const url = isEdit ? `${API}/${id}` : API;
  const method = isEdit ? "PUT" : "POST";

  const saveBtn = document.getElementById("saveBtn");
  saveBtn.disabled = true;
  saveBtn.textContent = "Saving...";

  try {
    const res = await fetch(url, {
      method,
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify(payload),
    });
    const data = await res.json();
    if (!res.ok) throw new Error(data.error || "Something went wrong");

    closeSheet();
    showToast(isEdit ? "Entry updated" : "Entry added to the register");
    loadEntries();
  } catch (err) {
    showToast(err.message);
  } finally {
    saveBtn.disabled = false;
    saveBtn.textContent = "Save Entry";
  }
}

// ----------------------------------------------------------------
// Delete confirmation
// ----------------------------------------------------------------
function openDeleteConfirm(entry) {
  pendingDeleteId = entry.id;
  el.confirmSub.textContent = `${entry.guest_name} — ${formatCurrency(entry.amount)} will be removed from the register.`;
  el.confirmOverlay.classList.add("open");
}

function closeDeleteConfirm() {
  el.confirmOverlay.classList.remove("open");
  pendingDeleteId = null;
}

async function handleConfirmDelete() {
  if (!pendingDeleteId) return;
  try {
    const res = await fetch(`${API}/${pendingDeleteId}`, { method: "DELETE" });
    if (!res.ok) throw new Error("Failed to delete entry");
    showToast("Entry removed");
    closeDeleteConfirm();
    loadEntries();
  } catch (err) {
    showToast("Could not remove entry.");
    console.error(err);
  }
}

// ----------------------------------------------------------------
// Export
// ----------------------------------------------------------------
function triggerDownload(url) {
  const a = document.createElement("a");
  a.href = url;
  a.download = "";
  document.body.appendChild(a);
  a.click();
  document.body.removeChild(a);
}

// ----------------------------------------------------------------
// Event wiring
// ----------------------------------------------------------------
el.addEntryBtn.addEventListener("click", openAddSheet);
el.sheetCloseBtn.addEventListener("click", closeSheet);
el.cancelBtn.addEventListener("click", closeSheet);
el.sheetOverlay.addEventListener("click", (e) => {
  if (e.target === el.sheetOverlay) closeSheet();
});
el.entryForm.addEventListener("submit", handleFormSubmit);

el.confirmCancelBtn.addEventListener("click", closeDeleteConfirm);
el.confirmDeleteBtn.addEventListener("click", handleConfirmDelete);
el.confirmOverlay.addEventListener("click", (e) => {
  if (e.target === el.confirmOverlay) closeDeleteConfirm();
});

el.searchInput.addEventListener("input", () => {
  clearTimeout(searchDebounce);
  searchDebounce = setTimeout(loadEntries, 250);
});
el.sortSelect.addEventListener("change", loadEntries);

el.exportExcelBtn.addEventListener("click", () => triggerDownload("/api/export/excel"));
el.exportPdfBtn.addEventListener("click", () => triggerDownload("/api/export/pdf"));

document.addEventListener("keydown", (e) => {
  if (e.key === "Escape") {
    if (el.sheetOverlay.classList.contains("open")) closeSheet();
    if (el.confirmOverlay.classList.contains("open")) closeDeleteConfirm();
  }
});

// ----------------------------------------------------------------
// Init
// ----------------------------------------------------------------
loadEntries();
