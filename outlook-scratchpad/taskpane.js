// Boss Replies Scratchpad — task pane logic.
//
// Persistence today: browser localStorage, scoped to this device/profile.
// See graph.js for the TODOs that turn this into real Microsoft To Do sync.

const STORAGE_KEY = "boss-replies-scratchpad-tasks";

let tasks = [];
let activeFilter = "all";

function loadTasks() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    tasks = raw ? JSON.parse(raw) : [];
  } catch (err) {
    console.error("Failed to load scratchpad tasks", err);
    tasks = [];
  }
}

function saveTasks() {
  try {
    localStorage.setItem(STORAGE_KEY, JSON.stringify(tasks));
  } catch (err) {
    console.error("Failed to save scratchpad tasks", err);
  }
}

function createEmptyTask(overrides = {}) {
  return {
    id: `task-${Date.now()}-${Math.random().toString(36).slice(2, 8)}`,
    title: "",
    note: "",
    checklist: [],
    priority: "normal",
    dueDate: "",
    status: "open",
    sourceEmail: null,
    createdAt: new Date().toISOString(),
    ...overrides,
  };
}

function findTask(id) {
  return tasks.find((t) => t.id === id);
}

function updateTask(id, patch) {
  const task = findTask(id);
  if (!task) return;
  Object.assign(task, patch);
  saveTasks();

  // TODO(graph.js): if this task is linked to a Microsoft To Do item,
  // push the patch with GraphClient.updateTask(task.todoTaskId, patch).
}

function deleteTask(id) {
  tasks = tasks.filter((t) => t.id !== id);
  saveTasks();
  render();

  // TODO(graph.js): also GraphClient.deleteTask(task.todoTaskId).
}

function addTask(task) {
  tasks.unshift(task);
  saveTasks();
  render();

  // TODO(graph.js): GraphClient.createTask(task) and store the returned
  // Microsoft To Do id back on task.todoTaskId.
}

function matchesFilter(task, filter) {
  if (filter === "all") return true;
  if (filter === "urgent") return task.priority === "urgent";
  return task.status === filter;
}

function render() {
  const list = document.getElementById("task-list");
  const emptyState = document.getElementById("empty-state");
  const template = document.getElementById("task-card-template");

  list.innerHTML = "";

  const visible = tasks.filter((t) => matchesFilter(t, activeFilter));

  emptyState.hidden = tasks.length > 0;
  if (tasks.length > 0 && visible.length === 0) {
    emptyState.hidden = false;
    emptyState.textContent = "No tasks match this filter.";
  } else if (tasks.length === 0) {
    emptyState.textContent = 'Nothing here yet. Open an email and click "Capture this email" to start.';
  }

  for (const task of visible) {
    const node = template.content.cloneNode(true);
    const card = node.querySelector(".task-card");
    card.dataset.taskId = task.id;
    card.dataset.status = task.priority === "urgent" ? "urgent-flag" : task.status;

    const titleInput = node.querySelector(".task-title");
    titleInput.value = task.title;
    titleInput.addEventListener("change", () => updateTask(task.id, { title: titleInput.value }));

    const deleteBtn = node.querySelector(".delete-btn");
    deleteBtn.addEventListener("click", () => {
      if (confirm("Delete this task?")) deleteTask(task.id);
    });

    const sourceLink = node.querySelector(".source-link");
    if (task.sourceEmail) {
      sourceLink.hidden = false;
      sourceLink.querySelector(".source-subject").textContent = task.sourceEmail.subject || "(no subject)";
      if (task.sourceEmail.webLink) {
        sourceLink.href = task.sourceEmail.webLink;
      } else {
        sourceLink.removeAttribute("href");
        sourceLink.style.cursor = "default";
        sourceLink.title = "Open-email link requires Microsoft Graph sync (not configured yet)";
      }
    }

    const noteInput = node.querySelector(".task-note");
    noteInput.value = task.note;
    noteInput.addEventListener("change", () => updateTask(task.id, { note: noteInput.value }));

    const checklistContainer = node.querySelector(".checklist");
    renderChecklist(checklistContainer, task);

    node.querySelector(".add-checklist-item").addEventListener("click", () => {
      task.checklist.push({ id: `item-${Date.now()}`, text: "", done: false });
      saveTasks();
      render();
    });

    const prioritySelect = node.querySelector(".task-priority");
    prioritySelect.value = task.priority;
    prioritySelect.addEventListener("change", () => {
      updateTask(task.id, { priority: prioritySelect.value });
      render();
    });

    const dueInput = node.querySelector(".task-due");
    dueInput.value = task.dueDate || "";
    dueInput.addEventListener("change", () => updateTask(task.id, { dueDate: dueInput.value }));

    const statusSelect = node.querySelector(".task-status");
    statusSelect.value = task.status;
    statusSelect.addEventListener("change", () => {
      updateTask(task.id, { status: statusSelect.value });
      render();
    });

    list.appendChild(node);
  }
}

function renderChecklist(container, task) {
  const template = document.getElementById("checklist-item-template");
  container.innerHTML = "";

  task.checklist.forEach((item, index) => {
    const node = template.content.cloneNode(true);
    const row = node.querySelector(".checklist-item");
    const checkbox = node.querySelector(".checklist-check");
    const textInput = node.querySelector(".checklist-text");

    checkbox.checked = item.done;
    textInput.value = item.text;
    row.classList.toggle("checked", item.done);

    checkbox.addEventListener("change", () => {
      item.done = checkbox.checked;
      row.classList.toggle("checked", item.done);
      saveTasks();
    });

    textInput.addEventListener("change", () => {
      item.text = textInput.value;
      saveTasks();
    });

    node.querySelector(".remove-checklist-item").addEventListener("click", () => {
      task.checklist.splice(index, 1);
      saveTasks();
      render();
    });

    container.appendChild(node);
  });
}

function setupFilters() {
  const row = document.getElementById("filter-row");
  row.addEventListener("click", (e) => {
    const chip = e.target.closest(".chip");
    if (!chip) return;
    row.querySelectorAll(".chip").forEach((c) => c.classList.remove("active"));
    chip.classList.add("active");
    activeFilter = chip.dataset.filter;
    render();
  });
}

function setupAddButton() {
  document.getElementById("add-task-btn").addEventListener("click", () => {
    addTask(createEmptyTask());
  });
}

function updateSyncBanner() {
  const detail = document.getElementById("sync-banner-detail");
  if (typeof GraphClient !== "undefined" && GraphClient.isConfigured()) {
    detail.textContent = GraphClient.isSignedIn()
      ? "Signed in — syncing to Microsoft To Do."
      : "Sign-in required to sync to Microsoft To Do.";
  } else {
    detail.textContent = "Microsoft To Do sync isn't configured yet (see graph.js).";
  }
}

async function captureCurrentEmail() {
  const item = Office.context.mailbox.item;
  if (!item) return;

  const subject = item.subject || "(no subject)";
  const sender = item.from ? `${item.from.displayName} <${item.from.emailAddress}>` : "";
  const itemId = item.itemId || null;

  // TODO(graph.js): replace with GraphClient.getMessageWebLink(itemId) once
  // Graph auth is wired up, so the link opens the message on any device.
  const webLink = await GraphClient.getMessageWebLink(itemId);

  const task = createEmptyTask({
    title: subject,
    sourceEmail: {
      subject,
      sender,
      itemId,
      webLink,
      capturedAt: new Date().toISOString(),
    },
  });

  addTask(task);
}

function setupCaptureButton() {
  const btn = document.getElementById("capture-btn");
  btn.addEventListener("click", captureCurrentEmail);

  const inOutlookHost =
    typeof Office !== "undefined" &&
    Office.context &&
    Office.context.mailbox &&
    Office.context.mailbox.item;

  btn.disabled = !inOutlookHost;
  if (!inOutlookHost) {
    btn.title = "Open this task pane from a read email in Outlook to enable capture.";
  }
}

Office.onReady(() => {
  loadTasks();
  setupFilters();
  setupAddButton();
  setupCaptureButton();
  updateSyncBanner();
  render();
});
