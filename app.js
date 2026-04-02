const STORAGE_KEY = "nandos-landos-jobs";
const FINISHED_STORAGE_KEY = "nandos-landos-finished-jobs";
const HQ_STORAGE_KEY = "nandos-landos-hq-address";
const SKIPPED_OCCURRENCES_KEY = "nandos-landos-skipped-occurrences";

const form = document.getElementById("jobForm");
const jobsList = document.getElementById("jobsList");
const housesList = document.getElementById("housesList");
const finishedJobsList = document.getElementById("finishedJobsList");
const statsCards = document.getElementById("statsCards");
const searchInput = document.getElementById("searchInput");
const houseSearchInput = document.getElementById("houseSearchInput");
const clearButton = document.getElementById("clearButton");
const exportButton = document.getElementById("exportButton");
const weekBoard = document.getElementById("weekBoard");
const template = document.getElementById("jobItemTemplate");
const finishedTemplate = document.getElementById("finishedJobItemTemplate");
const houseTemplate = document.getElementById("houseItemTemplate");
const hqAddressInput = document.getElementById("hqAddressInput");
const saveHqButton = document.getElementById("saveHqButton");
const routeWeekButton = document.getElementById("routeWeekButton");
const tabButtons = document.querySelectorAll(".tab-btn[data-tab]");
const tabPanels = document.querySelectorAll(".tab-panel");
const recurrenceValues = new Set(["One-time", "Bi-weekly"]);

const statusOrder = {
  "In Progress": 0,
  Scheduled: 1,
  Delayed: 2,
  Completed: 3,
};

let jobs = loadJobs();
let finishedJobs = loadFinishedJobs();
let skippedOccurrences = loadSkippedOccurrences();
let hqAddress = loadHqAddress();
let selectedMobileJobId = null;
let editingJobId = null;
let activeTabId = "upcomingPanel";
hqAddressInput.value = hqAddress;
render();

form.addEventListener("submit", (event) => {
  event.preventDefault();

  const recurrence = normalizeRecurrence(document.getElementById("recurrence").value);
  const payload = {
    id: editingJobId || crypto.randomUUID(),
    customerName: document.getElementById("customerName").value.trim(),
    contactInfo: document.getElementById("contactInfo").value.trim(),
    address: document.getElementById("address").value.trim(),
    serviceDate: document.getElementById("serviceDate").value,
    status: document.getElementById("status").value,
    price: Number(document.getElementById("price").value),
    serviceType: document.getElementById("serviceType").value.trim(),
    recurrence,
    notes: document.getElementById("notes").value.trim(),
  };

  const existingIndex = jobs.findIndex((job) => job.id === editingJobId);

  if (existingIndex >= 0) {
    jobs[existingIndex] = payload;
  } else {
    jobs.push(payload);
  }

  editingJobId = null;
  setFormMode("add");
  setActiveTab("upcomingPanel", true);
  persist();
  render();
  form.reset();
  scrollToOperationsCenter();
});

searchInput.addEventListener("input", () => {
  render();
});

houseSearchInput.addEventListener("input", () => {
  renderHouses();
});

tabButtons.forEach((button) => {
  button.addEventListener("click", () => {
    setActiveTab(button.dataset.tab);
  });
});

clearButton.addEventListener("click", () => {
  const confirmed = window.confirm("Delete all jobs from Nando's Landos dashboard?");
  if (!confirmed) return;
  jobs = [];
  finishedJobs = [];
  skippedOccurrences = [];
  editingJobId = null;
  setFormMode("add");
  persist();
  render();
});

exportButton.addEventListener("click", () => {
  exportCsv();
});

saveHqButton.addEventListener("click", () => {
  hqAddress = hqAddressInput.value.trim();
  persistHqAddress();
});

routeWeekButton.addEventListener("click", () => {
  void openWeeklyRoute();
});

jobsList.addEventListener("click", (event) => {
  const actionTarget = event.target.closest("button");
  if (!actionTarget) return;

  const item = event.target.closest(".job-item");
  if (!item) return;

  const { id, kind, occurrenceDate } = item.dataset;
  const selectedJob = jobs.find((job) => job.id === id);
  if (!selectedJob) return;

  if (actionTarget.classList.contains("delete")) {
    if (kind === "house") {
      const skippedDate = occurrenceDate || selectedJob.serviceDate;
      if (skippedDate && !isOccurrenceSkipped(selectedJob.id, skippedDate)) {
        skippedOccurrences.push({ sourceId: selectedJob.id, serviceDate: skippedDate });
      }
    } else {
      jobs = jobs.filter((job) => job.id !== id);
    }
  }

  if (actionTarget.classList.contains("mark-complete")) {
    const completionDate = occurrenceDate || selectedJob.serviceDate;
    const alreadyFinished = finishedJobs.some(
      (job) => job.sourceId === selectedJob.id && job.serviceDate === completionDate
    );

    if (!alreadyFinished) {
      finishedJobs.push(createFinishedRecord(selectedJob, completionDate, kind === "house" ? "house" : "one-time"));
    }

    if (kind !== "house") {
      jobs = jobs.filter((job) => job.id !== id);
    }
  }

  persist();
  render();
});

housesList.addEventListener("click", (event) => {
  const actionTarget = event.target.closest("button");
  if (!actionTarget) return;

  const item = event.target.closest(".job-item");
  if (!item) return;

  const { id } = item.dataset;
  const selectedJob = jobs.find((job) => job.id === id);
  if (!selectedJob) return;

  if (actionTarget.classList.contains("delete-house")) {
    jobs = jobs.filter((job) => job.id !== id);
    skippedOccurrences = skippedOccurrences.filter((entry) => entry.sourceId !== id);
  }

  if (actionTarget.classList.contains("edit-house")) {
    editingJobId = selectedJob.id;
    fillJobForm(selectedJob);
    setFormMode("edit");
    setActiveTab("upcomingPanel", true);
    form.scrollIntoView({ behavior: "smooth", block: "end" });
  }

  persist();
  render();
});

finishedJobsList.addEventListener("click", (event) => {
  const actionTarget = event.target.closest("button");
  if (!actionTarget) return;

  const item = event.target.closest(".job-item");
  if (!item) return;

  const { id } = item.dataset;
  const selectedFinishedJob = finishedJobs.find((job) => job.id === id);
  if (!selectedFinishedJob) return;

  if (actionTarget.classList.contains("delete-finished")) {
    finishedJobs = finishedJobs.filter((job) => job.id !== id);
  }

  if (actionTarget.classList.contains("reopen")) {
    if (selectedFinishedJob.sourceType === "house") {
      finishedJobs = finishedJobs.filter((job) => job.id !== id);
    } else {
      jobs.push({
        id: crypto.randomUUID(),
        customerName: selectedFinishedJob.customerName,
        contactInfo: selectedFinishedJob.contactInfo || "",
        address: selectedFinishedJob.address,
        serviceDate: selectedFinishedJob.serviceDate,
        status: "Scheduled",
        price: selectedFinishedJob.price,
        serviceType: selectedFinishedJob.serviceType,
        recurrence: "One-time",
        notes: selectedFinishedJob.notes || "",
      });
      finishedJobs = finishedJobs.filter((job) => job.id !== id);
    }
  }

  persist();
  render();
});

weekBoard.addEventListener("dragstart", (event) => {
  const card = event.target.closest(".mini-job[data-id]");
  if (!card || !event.dataTransfer) return;
  if (card.dataset.status === "Completed" || card.dataset.kind === "house") return;
  event.dataTransfer.setData("text/plain", card.dataset.id);
  event.dataTransfer.effectAllowed = "move";
});

weekBoard.addEventListener("dragover", (event) => {
  const column = event.target.closest(".day-column[data-date]");
  if (!column) return;
  event.preventDefault();
  event.dataTransfer.dropEffect = "move";
});

weekBoard.addEventListener("dragenter", (event) => {
  const column = event.target.closest(".day-column[data-date]");
  if (!column) return;
  column.classList.add("drag-over");
});

weekBoard.addEventListener("dragleave", (event) => {
  const column = event.target.closest(".day-column[data-date]");
  if (!column) return;
  column.classList.remove("drag-over");
});

weekBoard.addEventListener("drop", (event) => {
  const column = event.target.closest(".day-column[data-date]");
  if (!column || !event.dataTransfer) return;

  event.preventDefault();
  column.classList.remove("drag-over");

  const jobId = event.dataTransfer.getData("text/plain");
  if (!jobId) return;

  jobs = jobs.map((job) =>
    job.id === jobId ? { ...job, serviceDate: column.dataset.date } : job
  );
  selectedMobileJobId = null;
  persist();
  render();
});

weekBoard.addEventListener("click", (event) => {
  const miniCard = event.target.closest(".mini-job[data-id]");
  if (miniCard) {
    if (miniCard.dataset.status === "Completed" || miniCard.dataset.kind === "house") {
      selectedMobileJobId = null;
      renderWeekBoard();
      return;
    }
    const tappedId = miniCard.dataset.id;
    selectedMobileJobId = selectedMobileJobId === tappedId ? null : tappedId;
    renderWeekBoard();
    return;
  }

  const column = event.target.closest(".day-column[data-date]");
  if (!column || !selectedMobileJobId) return;

  jobs = jobs.map((job) =>
    job.id === selectedMobileJobId ? { ...job, serviceDate: column.dataset.date } : job
  );
  selectedMobileJobId = null;
  persist();
  render();
});

function render() {
  const removedCount = pruneExpiredFinishedJobs();
  if (removedCount > 0) {
    persist();
  }

  renderStats();
  renderWeekBoard();
  renderUpcomingJobs();
  renderHouses();
  renderFinishedJobs();
  setActiveTab(activeTabId, true);
}

function setActiveTab(tabId, silent = false) {
  activeTabId = tabId;

  tabButtons.forEach((button) => {
    const isActive = button.dataset.tab === tabId;
    button.classList.toggle("active", isActive);
    button.setAttribute("aria-selected", String(isActive));
  });

  tabPanels.forEach((panel) => {
    panel.classList.toggle("active", panel.id === tabId);
  });

  if (!silent) {
    render();
  }
}

function renderWeekBoard() {
  const weekDates = getWeekDates();
  weekBoard.innerHTML = "";

  weekDates.forEach((isoDate) => {
    const dayJobs = jobs
      .filter((job) => occursOnDate(job, isoDate))
      .sort((a, b) => statusOrder[a.status] - statusOrder[b.status]);

    const column = document.createElement("article");
    column.className = "day-column";
    column.dataset.date = isoDate;

    if (selectedMobileJobId) {
      column.classList.add("tap-target");
    }

    const heading = document.createElement("p");
    heading.className = "day-label";
    heading.textContent = formatBoardDate(isoDate);

    const stack = document.createElement("div");
    stack.className = "day-jobs";

    if (!dayJobs.length) {
      const empty = document.createElement("div");
      empty.className = "mini-job empty";
      empty.textContent = "No jobs";
      stack.appendChild(empty);
    }

    dayJobs.forEach((job) => {
      const mini = document.createElement("div");
      mini.className = "mini-job";
      const recurring = isRecurringJob(job);
      const occurrenceCompleted = recurring && isOccurrenceCompleted(job.id, isoDate);
      mini.draggable = !recurring && !occurrenceCompleted;
      mini.dataset.id = job.id;
      mini.dataset.kind = recurring ? "house" : "one-time";
      mini.dataset.status = occurrenceCompleted ? "Completed" : job.status;
      mini.dataset.occurrenceDate = isoDate;
      if (job.id === selectedMobileJobId) {
        mini.classList.add("selected");
      }
      if (occurrenceCompleted) {
        mini.classList.add("completed");
      }
      if (recurring) {
        mini.classList.add("house-schedule");
      }
      mini.innerHTML = `<strong>${job.customerName}</strong><span>${job.serviceType} | ${normalizeRecurrence(job.recurrence)}</span>`;
      stack.appendChild(mini);
    });

    column.appendChild(heading);
    column.appendChild(stack);
    weekBoard.appendChild(column);
  });
}

function renderStats() {
  const today = getTodayIso();
  const isFinishedTabActive = activeTabId === "finishedPanel";
  const revenueLabel = isFinishedTabActive ? "Realized Revenue" : "Projected Revenue";
  const revenueTotal = isFinishedTabActive
    ? finishedJobs.reduce((sum, job) => sum + (Number.isFinite(job.price) ? job.price : Number(job.price) || 0), 0)
    : jobs.reduce((sum, job) => sum + (Number.isFinite(job.price) ? job.price : Number(job.price) || 0), 0);

  const stats = [
    {
      label: "Total Jobs",
      value: jobs.length,
    },
    {
      label: "Today",
      value: jobs.filter((job) => occursOnDate(job, today)).length,
    },
    {
      label: "In Progress",
      value: jobs.filter((job) => job.status === "In Progress").length,
    },
    {
      label: revenueLabel,
      value: `$${revenueTotal.toFixed(2)}`,
    },
  ];

  statsCards.innerHTML = "";

  stats.forEach((stat, index) => {
    const card = document.createElement("article");
    card.className = "stat-card";
    card.style.animationDelay = `${index * 60}ms`;
    card.innerHTML = `<h3>${stat.label}</h3><p>${stat.value}</p>`;
    statsCards.appendChild(card);
  });
}

function renderUpcomingJobs() {
  const term = searchInput.value.trim().toLowerCase();
  const today = getTodayIso();

  const filteredJobs = getUpcomingEntries(today).filter((entry) => {
    if (!term) return true;
    const haystack = `${entry.customerName} ${entry.address} ${entry.contactInfo || ""}`.toLowerCase();
    return haystack.includes(term);
  });

  jobsList.innerHTML = "";

  if (!filteredJobs.length) {
    const empty = document.createElement("li");
    empty.className = "job-item";
    empty.innerHTML =
      "<h3>No jobs found</h3><p class='notes'>Add a new service call or recurring house to get started.</p>";
    jobsList.appendChild(empty);
    return;
  }

  filteredJobs.forEach((entry, index) => {
    const node = template.content.cloneNode(true);
    const listItem = node.querySelector(".job-item");
    const title = node.querySelector("h3");
    const pill = node.querySelector(".pill");
    const address = node.querySelector(".address");
    const contact = node.querySelector(".contact");
    const meta = node.querySelector(".meta");
    const notes = node.querySelector(".notes");
    const routeLink = node.querySelector(".route-link");
    const completeButton = node.querySelector(".mark-complete");
    const deleteButton = node.querySelector(".delete");

    listItem.dataset.id = entry.sourceId;
    listItem.dataset.kind = entry.kind;
    listItem.dataset.occurrenceDate = entry.occurrenceDate;
    listItem.style.animationDelay = `${index * 35}ms`;
    title.textContent = entry.customerName;
    pill.textContent = entry.kind === "house" ? "Bi-weekly" : entry.status;
    pill.dataset.status = entry.kind === "house" ? "Scheduled" : entry.status;
    address.textContent = `${entry.address}`;
    contact.textContent = entry.contactInfo ? `Contact: ${entry.contactInfo}` : "No contact listed";
    const normalizedPrice = Number.isFinite(entry.price) ? entry.price : Number(entry.price) || 0;
    meta.textContent = `${formatDate(entry.occurrenceDate)} | ${entry.serviceType} | ${normalizeRecurrence(entry.recurrence)} | $${normalizedPrice.toFixed(2)}`;
    notes.textContent = entry.notes || "No notes";
    routeLink.href = `https://www.google.com/maps/search/?api=1&query=${encodeURIComponent(entry.address)}`;

    completeButton.textContent = entry.kind === "house" ? "Mark Complete" : "Mark Complete";
    deleteButton.textContent = "Delete";

    jobsList.appendChild(node);
  });
}

function renderHouses() {
  const term = houseSearchInput.value.trim().toLowerCase();
  const today = getTodayIso();

  const recurringJobs = jobs.filter(isRecurringJob).filter((job) => {
    if (!term) return true;
    const haystack = `${job.customerName} ${job.address} ${job.contactInfo || ""}`.toLowerCase();
    return haystack.includes(term);
  });

  housesList.innerHTML = "";

  if (!recurringJobs.length) {
    const empty = document.createElement("li");
    empty.className = "job-item house-item";
    empty.innerHTML = "<h3>No recurring houses yet</h3><p class='notes'>Recurring houses will appear here.</p>";
    housesList.appendChild(empty);
    return;
  }

  recurringJobs.forEach((job, index) => {
    const node = houseTemplate.content.cloneNode(true);
    const listItem = node.querySelector(".job-item");
    const title = node.querySelector("h3");
    const contact = node.querySelector(".contact");
    const address = node.querySelector(".address");
    const meta = node.querySelector(".meta");
    const notes = node.querySelector(".notes");
    const editButton = node.querySelector(".edit-house");
    const deleteButton = node.querySelector(".delete-house");

    listItem.dataset.id = job.id;
    listItem.style.animationDelay = `${index * 35}ms`;
    title.textContent = job.customerName;
    contact.textContent = job.contactInfo ? `Contact: ${job.contactInfo}` : "No contact listed";
    address.textContent = job.address;
    meta.textContent = `Start ${formatDate(job.serviceDate)} | Next ${formatDate(getNextRecurringOccurrenceDate(job, today))} | Bi-weekly | $${(Number.isFinite(job.price) ? job.price : Number(job.price) || 0).toFixed(2)}`;
    notes.textContent = job.notes || "No notes";
    editButton.dataset.id = job.id;
    deleteButton.dataset.id = job.id;
    deleteButton.textContent = "Delete House";

    housesList.appendChild(node);
  });
}

function renderFinishedJobs() {
  const term = searchInput.value.trim().toLowerCase();
  const today = getTodayIso();

  const filteredFinished = finishedJobs
    .filter((job) => {
      if (!term) return true;
      return (
        job.customerName.toLowerCase().includes(term) ||
        job.address.toLowerCase().includes(term) ||
        (job.contactInfo || "").toLowerCase().includes(term)
      );
    })
    .sort((a, b) => (b.serviceDate || b.completedDate || today).localeCompare(a.serviceDate || a.completedDate || today));

  finishedJobsList.innerHTML = "";

  if (!filteredFinished.length) {
    const empty = document.createElement("li");
    empty.className = "job-item finished-item";
    empty.innerHTML = "<h3>No finished jobs yet</h3><p class='notes'>Completed jobs will appear here.</p>";
    finishedJobsList.appendChild(empty);
    return;
  }

  filteredFinished.forEach((job, index) => {
    const node = finishedTemplate.content.cloneNode(true);
    const listItem = node.querySelector(".job-item");
    const title = node.querySelector("h3");
    const address = node.querySelector(".address");
    const contact = node.querySelector(".contact");
    const meta = node.querySelector(".meta");
    const notes = node.querySelector(".notes");

    listItem.dataset.id = job.id;
    listItem.style.animationDelay = `${index * 35}ms`;
    title.textContent = job.customerName;
    address.textContent = `${job.address}`;
    const normalizedPrice = Number.isFinite(job.price) ? job.price : Number(job.price) || 0;
    contact.textContent = job.contactInfo ? `Contact: ${job.contactInfo}` : "No contact listed";
    meta.textContent = `${formatDate(job.serviceDate)} | ${job.serviceType} | ${normalizeRecurrence(job.recurrence)} | $${normalizedPrice.toFixed(2)}`;
    notes.textContent = job.notes || "No notes";

    finishedJobsList.appendChild(node);
  });
}

function getUpcomingEntries(todayIso) {
  const upcoming = [];

  jobs.forEach((job) => {
    if (isRecurringJob(job)) {
      const nextOccurrenceDate = getNextRecurringOccurrenceDate(job, todayIso);
      if (!nextOccurrenceDate) return;

      upcoming.push({
        ...job,
        kind: "house",
        sourceId: job.id,
        occurrenceDate: nextOccurrenceDate,
      });
      return;
    }

    if (job.status === "Completed") return;

    upcoming.push({
      ...job,
      kind: "one-time",
      sourceId: job.id,
      occurrenceDate: job.serviceDate,
    });
  });

  return upcoming.sort((a, b) => {
    const dateDiff = a.occurrenceDate.localeCompare(b.occurrenceDate);
    if (dateDiff !== 0) return dateDiff;
    return statusOrder[a.status] - statusOrder[b.status];
  });
}

function isRecurringJob(job) {
  return normalizeRecurrence(job.recurrence) === "Bi-weekly";
}

function isOccurrenceCompleted(jobId, isoDate) {
  return finishedJobs.some(
    (entry) => entry.sourceType === "house" && entry.sourceId === jobId && entry.serviceDate === isoDate
  );
}

function isOccurrenceSkipped(jobId, isoDate) {
  return skippedOccurrences.some(
    (entry) => entry.sourceId === jobId && entry.serviceDate === isoDate
  );
}

function getNextRecurringOccurrenceDate(job, fromIsoDate) {
  const from = parseIsoDate(fromIsoDate);
  const start = parseIsoDate(job.serviceDate) || from;
  if (!from || !start) return null;

  let year = from.getFullYear();
  let month = from.getMonth();

  for (let monthOffset = 0; monthOffset < 24; monthOffset += 1) {
    const currentYear = year + Math.floor((month + monthOffset) / 12);
    const currentMonth = (month + monthOffset) % 12;

    for (const day of [1, 14]) {
      const candidate = new Date(currentYear, currentMonth, day);
      const candidateIso = toIsoDate(candidate);
      if (candidate < from || candidate < start) continue;
      if (isOccurrenceCompleted(job.id, candidateIso)) continue;
      if (isOccurrenceSkipped(job.id, candidateIso)) continue;
      return candidateIso;
    }
  }

  return null;
}

function getDateForSort(job, todayIso) {
  return isRecurringJob(job) ? getNextRecurringOccurrenceDate(job, todayIso) || job.serviceDate : job.serviceDate;
}

function createFinishedRecord(job, serviceDate, sourceType) {
  return {
    id: crypto.randomUUID(),
    sourceId: job.id,
    sourceType,
    customerName: job.customerName,
    contactInfo: job.contactInfo || "",
    address: job.address,
    serviceDate,
    completedDate: getTodayIso(),
    status: "Completed",
    price: job.price,
    serviceType: job.serviceType,
    recurrence: sourceType === "house" ? "Bi-weekly" : "One-time",
    notes: job.notes || "",
  };
}

function fillJobForm(job) {
  document.getElementById("customerName").value = job.customerName || "";
  document.getElementById("contactInfo").value = job.contactInfo || "";
  document.getElementById("address").value = job.address || "";
  document.getElementById("serviceDate").value = job.serviceDate || "";
  document.getElementById("status").value = job.status || "Scheduled";
  document.getElementById("price").value = Number.isFinite(job.price) ? job.price : Number(job.price) || 0;
  document.getElementById("serviceType").value = job.serviceType || "";
  document.getElementById("recurrence").value = normalizeRecurrence(job.recurrence);
  document.getElementById("notes").value = job.notes || "";
}

function setFormMode(mode) {
  const formTitle = document.querySelector(".form-panel h2");
  const submitButton = form.querySelector("button[type='submit']");

  if (mode === "edit") {
    formTitle.textContent = "Update Job";
    submitButton.textContent = "Update Job";
    return;
  }

  formTitle.textContent = "Add New Job";
  submitButton.textContent = "Save Job";
}

function scrollToOperationsCenter() {
  const weeklySection = document.querySelector(".week-panel");

  if (!weeklySection) {
    window.scrollTo({ top: 0, behavior: "smooth" });
    return;
  }

  requestAnimationFrame(() => {
    requestAnimationFrame(() => {
      weeklySection.scrollIntoView({ behavior: "smooth", block: "center" });
    });
  });
}

function getWeekDates() {
  const now = new Date();
  const day = now.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  const monday = new Date(now);
  monday.setHours(0, 0, 0, 0);
  monday.setDate(now.getDate() + mondayOffset);

  return Array.from({ length: 7 }, (_, index) => {
    const current = new Date(monday);
    current.setDate(monday.getDate() + index);
    return current.toISOString().slice(0, 10);
  });
}

function occursOnDate(job, isoDate) {
  const recurrence = normalizeRecurrence(job.recurrence);
  if (recurrence === "One-time") {
    return job.serviceDate === isoDate;
  }

  const target = parseIsoDate(isoDate);
  const start = parseIsoDate(job.serviceDate);
  if (!target || !start || target < start) return false;

  const dayOfMonth = target.getDate();
  if (dayOfMonth !== 1 && dayOfMonth !== 14) return false;

  return !isOccurrenceCompleted(job.id, isoDate) && !isOccurrenceSkipped(job.id, isoDate);
}

function getNextOccurrenceDate(job, fromIsoDate) {
  const recurrence = normalizeRecurrence(job.recurrence);
  if (recurrence === "One-time") {
    return job.serviceDate;
  }

  const start = parseIsoDate(job.serviceDate);
  const from = parseIsoDate(fromIsoDate);
  if (!start || !from) return job.serviceDate;

  if (from <= start) {
    return job.serviceDate;
  }

  if (recurrence === "Bi-weekly") {
    const diffDays = Math.floor((from.getTime() - start.getTime()) / dayMs);
    const intervals = Math.ceil(diffDays / 14);
    const next = new Date(start);
    next.setDate(start.getDate() + intervals * 14);
    return toIsoDate(next);
  }

  let year = from.getFullYear();
  let month = from.getMonth();
  let candidate = getMonthlyOccurrenceDate(start, year, month);

  if (candidate < start) {
    month += 1;
    if (month > 11) {
      month = 0;
      year += 1;
    }
    candidate = getMonthlyOccurrenceDate(start, year, month);
  }

  if (candidate < from) {
    month += 1;
    if (month > 11) {
      month = 0;
      year += 1;
    }
    candidate = getMonthlyOccurrenceDate(start, year, month);
  }

  return toIsoDate(candidate);
}

function getFollowingOccurrenceDate(job) {
  const recurrence = normalizeRecurrence(job.recurrence);
  if (recurrence === "One-time") {
    return null;
  }

  const current = parseIsoDate(job.serviceDate);
  if (!current) {
    return null;
  }

  if (recurrence === "Bi-weekly") {
    const next = new Date(current);
    next.setDate(current.getDate() + 14);
    return toIsoDate(next);
  }

  let year = current.getFullYear();
  let month = current.getMonth() + 1;
  if (month > 11) {
    month = 0;
    year += 1;
  }

  return toIsoDate(getMonthlyOccurrenceDate(current, year, month));
}

function pruneExpiredFinishedJobs() {
  const weekStartIso = getCurrentWeekStartIso();
  const beforeCount = finishedJobs.length;

  finishedJobs = finishedJobs.filter((job) => {
    if (typeof job.serviceDate !== "string") return true;
    return job.serviceDate >= weekStartIso;
  });

  return beforeCount - finishedJobs.length;
}

function getCurrentWeekStartIso() {
  const now = new Date();
  const day = now.getDay();
  const mondayOffset = day === 0 ? -6 : 1 - day;
  const monday = new Date(now);
  monday.setHours(0, 0, 0, 0);
  monday.setDate(now.getDate() + mondayOffset);
  return toIsoDate(monday);
}

function normalizeRecurrence(value) {
  if (value === "Monthly") return "Bi-weekly";
  return recurrenceValues.has(value) ? value : "One-time";
}

function getMonthlyOccurrenceDate(anchorDate, year, month) {
  const anchorDay = anchorDate.getDate();
  const lastDayOfMonth = new Date(year, month + 1, 0).getDate();
  const day = Math.min(anchorDay, lastDayOfMonth);
  return new Date(year, month, day);
}

function parseIsoDate(isoDate) {
  const date = new Date(`${isoDate}T00:00:00`);
  if (Number.isNaN(date.getTime())) return null;
  return date;
}

function toIsoDate(date) {
  return date.toISOString().slice(0, 10);
}

function formatDate(isoDate) {
  const date = new Date(`${isoDate}T00:00:00`);
  return date.toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
  });
}

function formatBoardDate(isoDate) {
  const date = new Date(`${isoDate}T00:00:00`);
  return date.toLocaleDateString(undefined, {
    weekday: "short",
    month: "short",
    day: "numeric",
  });
}

function exportCsv() {
  const headers = [
    "Customer",
    "Address",
    "Next Service Date",
    "Recurrence",
    "Status",
    "Price",
    "Service Type",
    "Notes",
  ];

  const today = getTodayIso();

  const rows = jobs
    .slice()
    .sort((a, b) => getNextOccurrenceDate(a, today).localeCompare(getNextOccurrenceDate(b, today)))
    .map((job) => [
      job.customerName,
      job.address,
      getNextOccurrenceDate(job, today),
      normalizeRecurrence(job.recurrence),
      job.status,
      Number.isFinite(job.price) ? job.price.toFixed(2) : (Number(job.price) || 0).toFixed(2),
      job.serviceType,
      job.notes || "",
    ]);

  const csv = [headers, ...rows]
    .map((row) => row.map((value) => escapeCsv(value)).join(","))
    .join("\n");

  const blob = new Blob([csv], { type: "text/csv;charset=utf-8;" });
  const url = URL.createObjectURL(blob);
  const anchor = document.createElement("a");
  anchor.href = url;
  anchor.download = `nandos-landos-jobs-${getTodayIso()}.csv`;
  anchor.click();
  URL.revokeObjectURL(url);
}

async function openWeeklyRoute() {
  const hq = hqAddressInput.value.trim();
  if (!hq) {
    window.alert("Please set and save an HQ address first.");
    return;
  }

  hqAddress = hq;
  persistHqAddress();

  const weeklyStops = getWeeklyRouteStops();
  if (!weeklyStops.length) {
    window.alert("No jobs are scheduled for this week.");
    return;
  }

  routeWeekButton.disabled = true;
  const originalLabel = routeWeekButton.textContent;
  routeWeekButton.textContent = "Optimizing...";

  try {
    const optimizedStops = await optimizeRouteStops(hq, weeklyStops);
    const limitedStops = optimizedStops.slice(0, 24);
    const destination = limitedStops[limitedStops.length - 1];
    const waypoints = limitedStops.slice(0, -1);
    const mapsUrl = buildDirectionsUrl(hq, destination, waypoints);
    window.open(mapsUrl, "_blank", "noopener,noreferrer");
  } catch {
    const limitedStops = weeklyStops.slice(0, 24);
    const destination = limitedStops[limitedStops.length - 1];
    const waypoints = limitedStops.slice(0, -1);
    const fallbackUrl = buildDirectionsUrl(hq, destination, waypoints);
    window.open(fallbackUrl, "_blank", "noopener,noreferrer");
    window.alert("Route optimization API unavailable. Opened best-available route using weekly order.");
  } finally {
    routeWeekButton.disabled = false;
    routeWeekButton.textContent = originalLabel;
  }
}

function getWeeklyRouteStops() {
  const weekDates = getWeekDates();
  const seenAddresses = new Set();
  const stops = [];

  weekDates.forEach((isoDate) => {
    const dayStops = jobs
      .filter((job) => occursOnDate(job, isoDate))
      .sort((a, b) => statusOrder[a.status] - statusOrder[b.status]);

    dayStops.forEach((job) => {
      const normalizedAddress = job.address.trim().toLowerCase();
      if (!normalizedAddress || seenAddresses.has(normalizedAddress)) return;
      seenAddresses.add(normalizedAddress);
      stops.push(job.address.trim());
    });
  });

  return stops;
}

function buildDirectionsUrl(origin, destination, waypoints) {
  const params = new URLSearchParams({
    api: "1",
    origin,
    destination,
    travelmode: "driving",
  });

  if (waypoints.length) {
    params.set("waypoints", waypoints.join("|"));
  }

  return `https://www.google.com/maps/dir/?${params.toString()}`;
}

async function optimizeRouteStops(hqAddress, stops) {
  const uniqueStops = Array.from(new Set(stops.map((address) => address.trim()).filter(Boolean)));
  if (uniqueStops.length <= 1) {
    return uniqueStops;
  }

  const geocoded = await geocodeAddresses([hqAddress, ...uniqueStops]);
  const hqPoint = geocoded[0];
  const stopPoints = geocoded.slice(1);

  if (!hqPoint || stopPoints.some((point) => !point)) {
    throw new Error("Failed geocoding one or more addresses");
  }

  const coordinates = [hqPoint, ...stopPoints]
    .map((point) => `${point.lon},${point.lat}`)
    .join(";");

  const response = await fetch(
    `https://router.project-osrm.org/trip/v1/driving/${coordinates}?source=first&roundtrip=false&overview=false`
  );
  if (!response.ok) {
    throw new Error("OSRM optimization request failed");
  }

  const payload = await response.json();
  if (payload.code !== "Ok" || !Array.isArray(payload.waypoints)) {
    throw new Error("Invalid OSRM optimization response");
  }

  const indexByAddress = new Map(stopPoints.map((point, index) => [point.address, index]));
  const orderedStops = payload.waypoints
    .filter((waypoint) => waypoint.waypoint_index > 0)
    .sort((a, b) => a.waypoint_index - b.waypoint_index)
    .map((waypoint) => {
      const matchedAddress = findNearestAddress(waypoint.location, stopPoints);
      return uniqueStops[indexByAddress.get(matchedAddress.address)];
    })
    .filter(Boolean);

  if (!orderedStops.length) {
    throw new Error("No optimized stops returned");
  }

  return orderedStops;
}

async function geocodeAddresses(addresses) {
  const results = [];

  for (const address of addresses) {
    const params = new URLSearchParams({
      q: address,
      format: "jsonv2",
      limit: "1",
    });

    const response = await fetch(`https://nominatim.openstreetmap.org/search?${params.toString()}`);
    if (!response.ok) {
      throw new Error("Geocoding request failed");
    }

    const data = await response.json();
    if (!Array.isArray(data) || !data.length) {
      throw new Error("No geocode match found");
    }

    results.push({
      address,
      lat: Number(data[0].lat),
      lon: Number(data[0].lon),
    });
  }

  return results;
}

function findNearestAddress(osrmLocation, candidates) {
  const [lon, lat] = osrmLocation;
  let best = candidates[0];
  let bestDistance = Number.POSITIVE_INFINITY;

  candidates.forEach((candidate) => {
    const distance = Math.abs(candidate.lat - lat) + Math.abs(candidate.lon - lon);
    if (distance < bestDistance) {
      best = candidate;
      bestDistance = distance;
    }
  });

  return best;
}

function escapeCsv(value) {
  const text = String(value ?? "");
  return `"${text.replace(/"/g, '""')}"`;
}

function persist() {
  localStorage.setItem(STORAGE_KEY, JSON.stringify(jobs));
  localStorage.setItem(FINISHED_STORAGE_KEY, JSON.stringify(finishedJobs));
  localStorage.setItem(SKIPPED_OCCURRENCES_KEY, JSON.stringify(skippedOccurrences));
}

function persistHqAddress() {
  localStorage.setItem(HQ_STORAGE_KEY, hqAddress);
}

function getTodayIso() {
  return new Date().toISOString().slice(0, 10);
}

function loadHqAddress() {
  return localStorage.getItem(HQ_STORAGE_KEY) || "";
}

function loadJobs() {
  try {
    const raw = localStorage.getItem(STORAGE_KEY);
    if (!raw) return [];

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];

    return parsed.filter((job) => {
      if (job && job.status === "Completed") return false;
      return (
        typeof job.id === "string" &&
        typeof job.customerName === "string" &&
        typeof job.address === "string" &&
        typeof job.serviceDate === "string" &&
        typeof job.status === "string" &&
        typeof job.serviceType === "string"
      );
    }).map((job) => ({
      ...job,
      contactInfo: job.contactInfo || "",
      recurrence: normalizeRecurrence(job.recurrence),
    }));
  } catch {
    return [];
  }
}

function loadFinishedJobs() {
  try {
    const rawFinished = localStorage.getItem(FINISHED_STORAGE_KEY);
    if (rawFinished) {
      const parsedFinished = JSON.parse(rawFinished);
      if (Array.isArray(parsedFinished)) {
        return parsedFinished.filter((job) => {
          return (
            typeof job.id === "string" &&
            typeof job.customerName === "string" &&
            typeof job.address === "string" &&
            typeof job.serviceDate === "string" &&
            typeof job.sourceId === "string"
          );
        });
      }
    }

    const rawLegacy = localStorage.getItem(STORAGE_KEY);
    if (!rawLegacy) return [];

    const parsedLegacy = JSON.parse(rawLegacy);
    if (!Array.isArray(parsedLegacy)) return [];

    return parsedLegacy
      .filter((job) => job && job.status === "Completed")
      .map((job) => ({
        id: job.id,
        sourceId: job.id,
        sourceType: isRecurringJob(job) ? "house" : "one-time",
        customerName: job.customerName,
        contactInfo: job.contactInfo || "",
        address: job.address,
        serviceDate: job.serviceDate,
        completedDate: job.completedDate || getTodayIso(),
        status: "Completed",
        price: job.price,
        serviceType: job.serviceType,
        recurrence: normalizeRecurrence(job.recurrence),
        notes: job.notes || "",
      }));
  } catch {
    return [];
  }
}

function loadSkippedOccurrences() {
  try {
    const raw = localStorage.getItem(SKIPPED_OCCURRENCES_KEY);
    if (!raw) return [];

    const parsed = JSON.parse(raw);
    if (!Array.isArray(parsed)) return [];

    return parsed.filter((entry) => {
      return (
        entry &&
        typeof entry.sourceId === "string" &&
        typeof entry.serviceDate === "string"
      );
    });
  } catch {
    return [];
  }
}

if ("serviceWorker" in navigator && (window.location.protocol === "https:" || window.location.hostname === "localhost")) {
  window.addEventListener("load", () => {
    navigator.serviceWorker.register("./sw.js").catch(() => {
      // Registration failures are non-fatal; app continues without offline support.
    });
  });
}
