// ============================================================
//  Workshop Registration App – script.js
//  Handles: form submission, participant table, search, download
// ============================================================

// ── DOM references ─────────────────────────────────────────
const form           = document.getElementById("registrationForm");
const submitBtn      = document.getElementById("submitBtn");
const btnText        = submitBtn.querySelector(".btn-text");
const btnLoader      = submitBtn.querySelector(".btn-loader");
const toast          = document.getElementById("toast");
const tbody          = document.getElementById("participantsTbody");
const countBadge     = document.getElementById("countBadge");
const searchInput    = document.getElementById("searchInput");

// All participants loaded from the server
let allParticipants = [];


// ════════════════════════════════════════════════════════════
//  TOAST HELPER
// ════════════════════════════════════════════════════════════
let toastTimer;

/**
 * Show a toast notification.
 * @param {string} message - Message text
 * @param {"success"|"error"} type
 * @param {number} duration - Auto-hide after ms (0 = never)
 */
function showToast(message, type = "success", duration = 6000) {
  clearTimeout(toastTimer);
  toast.textContent = message;
  toast.className = `toast ${type}`;

  if (duration > 0) {
    toastTimer = setTimeout(() => {
      toast.style.opacity = "0";
      setTimeout(() => { toast.className = "toast hidden"; toast.style.opacity = ""; }, 350);
    }, duration);
  }
}

function hideToast() {
  toast.className = "toast hidden";
}


// ════════════════════════════════════════════════════════════
//  FORM VALIDATION
// ════════════════════════════════════════════════════════════
function setError(fieldId, errorId, message) {
  const field = document.getElementById(fieldId);
  const errEl = document.getElementById(errorId);
  field.classList.toggle("invalid", !!message);
  if (errEl) errEl.textContent = message || "";
}

function clearErrors() {
  ["name","email","phone","workshop"].forEach(id => {
    setError(id, id + "Error", "");
  });
}

function validateForm(data) {
  let valid = true;

  if (!data.name.trim()) {
    setError("name", "nameError", "Full name is required.");
    valid = false;
  }

  const emailRe = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  if (!data.email.trim()) {
    setError("email", "emailError", "Email address is required.");
    valid = false;
  } else if (!emailRe.test(data.email.trim())) {
    setError("email", "emailError", "Enter a valid email address.");
    valid = false;
  }

  const phoneRe = /^[\d\s+\-().]{7,15}$/;
  if (!data.phone.trim()) {
    setError("phone", "phoneError", "Phone number is required.");
    valid = false;
  } else if (!phoneRe.test(data.phone.trim())) {
    setError("phone", "phoneError", "Enter a valid phone number.");
    valid = false;
  }

  if (!data.workshop) {
    setError("workshop", "workshopError", "Please select a workshop.");
    valid = false;
  }

  return valid;
}


// ════════════════════════════════════════════════════════════
//  FORM SUBMISSION
// ════════════════════════════════════════════════════════════
form.addEventListener("submit", async (e) => {
  e.preventDefault();
  clearErrors();
  hideToast();

  // Collect form values
  const data = {
    name:     document.getElementById("name").value,
    email:    document.getElementById("email").value,
    phone:    document.getElementById("phone").value,
    workshop: document.getElementById("workshop").value,
  };

  // Client-side validation
  if (!validateForm(data)) return;

  // UI: loading state
  submitBtn.disabled = true;
  btnText.classList.add("hidden");
  btnLoader.classList.remove("hidden");

  try {
    const res = await fetch("/register", {
      method:  "POST",
      headers: { "Content-Type": "application/json" },
      body:    JSON.stringify(data),
    });

    const json = await res.json();

    if (res.ok && json.success) {
      // ✅ Success
      showToast(`✅ ${json.message}`, "success", 8000);
      form.reset();
      // Refresh the participant list and highlight the new row
      await loadParticipants(json.participant.id);
    } else {
      // ⚠️ Server-side error (e.g. duplicate)
      showToast(`⚠️ ${json.message}`, "error", 0);
    }
  } catch (err) {
    console.error("Network error:", err);
    showToast("🔌 Could not reach the server. Is it running?", "error", 0);
  } finally {
    // Restore button
    submitBtn.disabled = false;
    btnText.classList.remove("hidden");
    btnLoader.classList.add("hidden");
  }
});


// ════════════════════════════════════════════════════════════
//  LOAD & RENDER PARTICIPANTS
// ════════════════════════════════════════════════════════════

/**
 * Fetch participants from the server and re-render the table.
 * @param {number|null} newId - If set, the new row gets a highlight animation.
 */
async function loadParticipants(newId = null) {
  try {
    const res  = await fetch("/participants");
    const json = await res.json();

    if (json.success) {
      allParticipants = json.participants;
      renderTable(allParticipants, newId);
      countBadge.textContent = allParticipants.length;
    }
  } catch (err) {
    console.error("Failed to load participants:", err);
  }
}

/**
 * Render the participant list into the table body.
 * @param {Array}       list  - Array of participant objects to display
 * @param {number|null} newId - Participant id to highlight as new
 */
function renderTable(list, newId = null) {
  tbody.innerHTML = "";

  if (list.length === 0) {
    tbody.innerHTML = `
      <tr id="emptyRow">
        <td colspan="6" class="empty-state">
          <div class="empty-icon">📋</div>
          <p>No participants yet. Be the first to register!</p>
        </td>
      </tr>`;
    return;
  }

  list.forEach((p, i) => {
    // Find original index for numbering (even after filtering)
    const globalIndex = allParticipants.findIndex(x => x.id === p.id);
    const tr = document.createElement("tr");

    if (p.id === newId) tr.classList.add("new-row");

    tr.innerHTML = `
      <td>${globalIndex + 1}</td>
      <td><strong style="color:#fff">${escHtml(p.name)}</strong></td>
      <td>${escHtml(p.email)}</td>
      <td>${escHtml(p.phone)}</td>
      <td><span class="workshop-pill">${escHtml(p.workshop)}</span></td>
      <td style="font-size:0.82rem;color:rgba(255,255,255,0.6)">${escHtml(p.registeredAt)}</td>
    `;
    tbody.appendChild(tr);
  });
}

/** Escape HTML to prevent XSS in dynamically injected content */
function escHtml(str) {
  const div = document.createElement("div");
  div.textContent = String(str);
  return div.innerHTML;
}


// ════════════════════════════════════════════════════════════
//  SEARCH / FILTER
// ════════════════════════════════════════════════════════════
function filterTable() {
  const query = searchInput.value.trim().toLowerCase();

  if (!query) {
    renderTable(allParticipants);
    return;
  }

  const filtered = allParticipants.filter(p =>
    p.name.toLowerCase().includes(query)     ||
    p.email.toLowerCase().includes(query)    ||
    p.workshop.toLowerCase().includes(query) ||
    p.phone.includes(query)
  );

  renderTable(filtered);
}


// ════════════════════════════════════════════════════════════
//  DOWNLOAD
// ════════════════════════════════════════════════════════════
function downloadDoc() {
  window.location.href = "/download";
}


// ════════════════════════════════════════════════════════════
//  INITIALISE
// ════════════════════════════════════════════════════════════
loadParticipants();
