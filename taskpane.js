// ─────────────────────────────────────────────────────────────
//  CPL CRM Outlook Add-in — taskpane.js
// ─────────────────────────────────────────────────────────────

const CONFIG = {
  CRM_BASE_URL: "https://cpl-marketing.base44.app",
  API_KEY: "",
};

let currentContact = null;
let selectedActivityType = "email";

Office.onReady((info) => {
  if (info.host === Office.HostType.Outlook) {
    initAddin();
  }
});

function initAddin() {
  const today = new Date().toISOString().split("T")[0];
  document.getElementById("activityDate").value = today;
  document.getElementById("taskDueDate").value = today;
  lookupCurrentContact();
}

function lookupCurrentContact() {
  showState("loading");
  setSyncStatus("loading");
  try {
    const item = Office.context.mailbox.item;
    const senderEmail = item.from?.emailAddress || item.sender?.emailAddress;
    if (!senderEmail) { showError("Geen e-mailadres gevonden in dit bericht."); return; }
    fetchContact(senderEmail);
  } catch (err) {
    showError("Fout bij lezen van e-mail: " + err.message);
  }
}

async function fetchContact(email) {
  try {
    const url = `${CONFIG.CRM_BASE_URL}/api/outlookAddinApi?action=lookupContact&email=${encodeURIComponent(email)}`;
    const response = await fetch(url, { method: "GET", headers: { "Content-Type": "application/json" } });
    if (response.status === 404) { showNotFound(email); return; }
    if (!response.ok) throw new Error(`API fout: ${response.status} ${response.statusText}`);
    const data = await response.json();
    currentContact = data;
    renderContact(data);
    setSyncStatus("ok");
  } catch (err) {
    showError(err.message);
    setSyncStatus("error");
  }
}

function retryLookup() { lookupCurrentContact(); }

function renderContact(data) {
  const name    = data.name    || data.contact_name || "Onbekend";
  const company = data.company || data.company_name || "—";
  const email   = data.email   || "—";
  const status  = data.status  || data.contact_status || "";
  document.getElementById("contactAvatar").textContent  = name.charAt(0).toUpperCase();
  document.getElementById("contactName").textContent    = name;
  document.getElementById("contactCompany").textContent = company;
  document.getElementById("contactEmail").textContent   = email;
  const badge = document.getElementById("statusBadge");
  badge.textContent = formatStatus(status);
  badge.className   = "status-badge " + getStatusClass(status);
  renderTasks(data.tasks || data.open_tasks || []);
  renderActivities(data.activities || data.recent_activities || []);
  showState("contact");
}

function formatStatus(status) {
  const map = { suspect:"Suspect",Suspect:"Suspect",prospect:"Prospect",Prospect:"Prospect",lead:"Lead",Lead:"Lead",customer:"Klant",Customer:"Klant" };
  return map[status] || status || "—";
}

function getStatusClass(status) {
  const s = (status || "").toLowerCase();
  if (s === "suspect")  return "status-suspect";
  if (s === "prospect") return "status-prospect";
  if (s === "lead")     return "status-lead";
  if (s === "customer") return "status-customer";
  return "status-suspect";
}

function renderTasks(tasks) {
  const countEl = document.getElementById("taskCount");
  countEl.textContent = tasks.length;
  countEl.style.display = tasks.length > 0 ? "inline" : "none";
  const listEl = document.getElementById("tasksList");
  if (!tasks.length) { listEl.innerHTML = '<p class="empty-hint">Geen open taken</p>'; return; }
  listEl.innerHTML = tasks.map(task => {
    const desc = task.description || task.title || task.name || "Taak";
    const dueDate = task.due_date || task.dueDate || "";
    const isOverdue = dueDate && new Date(dueDate) < new Date();
    return `<div class="task-item"><div class="task-check"></div><div class="task-body"><div class="task-description">${escapeHtml(desc)}</div>${dueDate ? `<div class="task-due ${isOverdue?"overdue":""}">${isOverdue?"⚠ ":""}${formatDate(dueDate)}</div>` : ""}</div></div>`;
  }).join("");
}

function renderActivities(activities) {
  const icons = { email:"📧", call:"📞", meeting:"📅", note:"📝" };
  const makeHtml = (list) => {
    if (!list.length) return '<p class="empty-hint">Geen activiteiten gevonden</p>';
    return list.map(act => {
      const type    = (act.type || act.activity_type || "note").toLowerCase();
      const subject = act.subject || act.title || "Activiteit";
      const date    = act.date || act.created_at || "";
      return `<div class="activity-item"><div class="activity-icon">${icons[type]||"📝"}</div><div class="activity-body"><div class="activity-subject">${escapeHtml(subject)}</div><div class="activity-meta">${formatDate(date)}</div></div></div>`;
    }).join("");
  };
  document.getElementById("recentActivity").innerHTML = makeHtml(activities.slice(0, 3));
  document.getElementById("allActivity").innerHTML = makeHtml(activities);
}

function openLogActivity() {
  if (!currentContact) return;
  document.getElementById("activitySubject").value = "";
  document.getElementById("activityNotes").value = "";
  openModal("modalActivity");
}

function selectType(btn) {
  document.querySelectorAll(".type-btn").forEach(b => b.classList.remove("active"));
  btn.classList.add("active");
  selectedActivityType = btn.dataset.type;
}

async function submitActivity() {
  const subject = document.getElementById("activitySubject").value.trim();
  const notes   = document.getElementById("activityNotes").value.trim();
  const date    = document.getElementById("activityDate").value;
  if (!subject) { showToast("Vul een onderwerp in", "error"); return; }
  const payload = { action:"logActivity", contact_id: currentContact.id||currentContact.contact_id, type:selectedActivityType, subject, notes, date };
  try {
    const response = await fetch(`${CONFIG.CRM_BASE_URL}/api/outlookAddinApi`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) });
    if (!response.ok) throw new Error(`API fout: ${response.status}`);
    closeModal("modalActivity");
    showToast("Activiteit opgeslagen ✓", "success");
    setTimeout(() => fetchContact(currentContact.email), 600);
  } catch (err) { showToast("Fout: " + err.message, "error"); }
}

function openNewTask() {
  if (!currentContact) return;
  document.getElementById("taskDescription").value = "";
  openModal("modalTask");
}

async function submitTask() {
  const description = document.getElementById("taskDescription").value.trim();
  const dueDate     = document.getElementById("taskDueDate").value;
  const assignedTo  = document.getElementById("taskAssigned").value.trim();
  if (!description) { showToast("Vul een omschrijving in", "error"); return; }
  const payload = { action:"createTask", contact_id:currentContact.id||currentContact.contact_id, description, due_date:dueDate, assigned_to:assignedTo };
  try {
    const response = await fetch(`${CONFIG.CRM_BASE_URL}/api/outlookAddinApi`, { method:"POST", headers:{"Content-Type":"application/json"}, body:JSON.stringify(payload) });
    if (!response.ok) throw new Error(`API fout: ${response.status}`);
    closeModal("modalTask");
    showToast("Taak aangemaakt ✓", "success");
    setTimeout(() => fetchContact(currentContact.email), 600);
  } catch (err) { showToast("Fout: " + err.message, "error"); }
}

function showState(state) {
  ["loadingState","notFoundState","errorState","contactPanel"].forEach(id => document.getElementById(id).classList.add("hidden"));
  const map = { loading:"loadingState", notFound:"notFoundState", error:"errorState", contact:"contactPanel" };
  if (map[state]) document.getElementById(map[state]).classList.remove("hidden");
}
function showNotFound(email) { document.getElementById("notFoundEmail").textContent = email; showState("notFound"); setSyncStatus("ok"); }
function showError(msg) { document.getElementById("errorMessage").textContent = msg; showState("error"); setSyncStatus("error"); }
function setSyncStatus(status) { document.getElementById("syncStatus").className = "sync-dot sync-" + status; }
function switchTab(name, btn) {
  document.querySelectorAll(".tab").forEach(t => t.classList.remove("active"));
  document.querySelectorAll(".tab-content").forEach(c => c.classList.add("hidden"));
  btn.classList.add("active");
  document.getElementById("tab-" + name).classList.remove("hidden");
}
function openModal(id) { document.getElementById(id).classList.remove("hidden"); }
function closeModal(id) { document.getElementById(id).classList.add("hidden"); }
let toastTimer = null;
function showToast(msg, type = "") {
  const toast = document.getElementById("toast");
  toast.textContent = msg;
  toast.className = "toast" + (type ? " " + type : "");
  toast.classList.remove("hidden");
  if (toastTimer) clearTimeout(toastTimer);
  toastTimer = setTimeout(() => toast.classList.add("hidden"), 3000);
}
function formatDate(dateStr) {
  if (!dateStr) return "";
  try { return new Date(dateStr).toLocaleDateString("nl-NL", { day:"numeric", month:"short", year:"numeric" }); }
  catch { return dateStr; }
}
function escapeHtml(str) {
  return String(str).replace(/&/g,"&amp;").replace(/</g,"&lt;").replace(/>/g,"&gt;").replace(/"/g,"&quot;");
}