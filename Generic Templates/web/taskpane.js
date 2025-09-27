const STORAGE_KEYS = {
  templates: "genericTemplates.templates",
  teamEmails: "genericTemplates.teamEmails"
};

let ui = {
  templateSelect: null,
  statusArea: null,
  createDraftsBtn: null,
  templateList: null,
  teamList: null,
  templateForm: null,
  toggleTemplateFormBtn: null,
  templateNameInput: null,
  templateBodyInput: null,
  cancelTemplateBtn: null,
  teamForm: null,
  teamEmailInput: null,
  referenceInput: null
};

const state = {
  templates: [],
  teamEmails: []
};

Office.onReady(async () => {
  cacheUiElements();
  bindEventHandlers();
  await loadState();
  renderTemplates();
  renderTeamEmails();
  ensureTemplateSelect();
});

function cacheUiElements() {
  ui = {
    templateSelect: document.getElementById("templateSelect"),
    statusArea: document.getElementById("statusArea"),
    createDraftsBtn: document.getElementById("createDraftsBtn"),
    templateList: document.getElementById("templateList"),
    teamList: document.getElementById("teamList"),
    templateForm: document.getElementById("templateForm"),
    toggleTemplateFormBtn: document.getElementById("toggleTemplateFormBtn"),
    templateNameInput: document.getElementById("templateNameInput"),
    templateBodyInput: document.getElementById("templateBodyInput"),
    cancelTemplateBtn: document.getElementById("cancelTemplateBtn"),
    teamForm: document.getElementById("teamForm"),
    teamEmailInput: document.getElementById("teamEmailInput"),
    referenceInput: document.getElementById("referenceInput")
  };
}

function bindEventHandlers() {
  ui.createDraftsBtn.addEventListener("click", handleCreateDrafts);
  ui.toggleTemplateFormBtn.addEventListener("click", toggleTemplateForm);
  ui.cancelTemplateBtn.addEventListener("click", hideTemplateForm);
  ui.templateForm.addEventListener("submit", handleTemplateSubmit);
  ui.teamForm.addEventListener("submit", handleTeamEmailSubmit);
}

async function loadState() {
  try {
    const templatesValue = Office.context.roamingSettings.get(STORAGE_KEYS.templates);
    const emailsValue = Office.context.roamingSettings.get(STORAGE_KEYS.teamEmails);
    state.templates = Array.isArray(templatesValue) ? templatesValue : [];
    state.teamEmails = Array.isArray(emailsValue) ? emailsValue : [];
  } catch (error) {
    console.error("Failed to load roaming settings", error);
    state.templates = [];
    state.teamEmails = [];
  }
}

async function persistState() {
  try {
    Office.context.roamingSettings.set(STORAGE_KEYS.templates, state.templates);
    Office.context.roamingSettings.set(STORAGE_KEYS.teamEmails, state.teamEmails);
    await saveRoamingSettings();
  } catch (error) {
    console.error("Failed to persist roaming settings", error);
    setStatus("Unable to save settings. Please try again.", true);
  }
}

function saveRoamingSettings() {
  return new Promise((resolve, reject) => {
    Office.context.roamingSettings.saveAsync((asyncResult) => {
      if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
        resolve();
      } else {
        reject(asyncResult.error);
      }
    });
  });
}

function renderTemplates() {
  ui.templateList.innerHTML = "";
  if (!state.templates.length) {
    ui.templateList.innerHTML = "<li>No templates saved yet.</li>";
  } else {
    state.templates.forEach((template) => {
      const li = document.createElement("li");
      li.className = "list-item";
      li.innerHTML = `<span>${escapeHtml(template.name)}</span>`;
      const button = document.createElement("button");
      button.className = "remove-btn";
      button.type = "button";
      button.textContent = "Remove";
      button.addEventListener("click", () => removeTemplate(template.id));
      li.appendChild(button);
      ui.templateList.appendChild(li);
    });
  }
  ensureTemplateSelect();
}

function renderTeamEmails() {
  ui.teamList.innerHTML = "";
  if (!state.teamEmails.length) {
    ui.teamList.innerHTML = "<li>No team mail IDs saved yet.</li>";
    return;
  }

  state.teamEmails.forEach((address) => {
    const li = document.createElement("li");
    li.className = "list-item";
    const label = document.createElement("span");
    label.textContent = address;
    const button = document.createElement("button");
    button.className = "remove-btn";
    button.type = "button";
    button.textContent = "Remove";
    button.addEventListener("click", () => removeTeamEmail(address));
    li.appendChild(label);
    li.appendChild(button);
    ui.teamList.appendChild(li);
  });
}

function ensureTemplateSelect() {
  ui.templateSelect.innerHTML = "";
  if (!state.templates.length) {
    const option = document.createElement("option");
    option.value = "";
    option.textContent = "No templates available";
    ui.templateSelect.appendChild(option);
    ui.templateSelect.disabled = true;
    ui.createDraftsBtn.disabled = true;
  } else {
    state.templates.forEach((template) => {
      const option = document.createElement("option");
      option.value = template.id;
      option.textContent = template.name;
      ui.templateSelect.appendChild(option);
    });
    ui.templateSelect.disabled = false;
    ui.createDraftsBtn.disabled = false;
  }
}

function toggleTemplateForm() {
  const isHidden = ui.templateForm.classList.contains("hidden");
  if (isHidden) {
    ui.templateForm.classList.remove("hidden");
    ui.templateNameInput.focus();
  } else {
    hideTemplateForm();
  }
}

function hideTemplateForm() {
  ui.templateForm.classList.add("hidden");
  ui.templateForm.reset();
}

async function handleTemplateSubmit(event) {
  event.preventDefault();
  const name = ui.templateNameInput.value.trim();
  const body = ui.templateBodyInput.value.trim();

  if (!name || !body) {
    setStatus("Template name and body are required.", true);
    return;
  }

  const template = {
    id: generateId(),
    name,
    body
  };

  state.templates.push(template);
  await persistState();
  hideTemplateForm();
  renderTemplates();
  setStatus(`Saved template "${template.name}".`);
}

async function handleTeamEmailSubmit(event) {
  event.preventDefault();
  const email = (ui.teamEmailInput.value || "").trim().toLowerCase();
  if (!email) {
    return;
  }
  if (!isValidEmail(email)) {
    setStatus("Please enter a valid email address.", true);
    return;
  }
  if (state.teamEmails.includes(email)) {
    setStatus("Email already exists in the list.", true);
    return;
  }

  state.teamEmails.push(email);
  ui.teamEmailInput.value = "";
  await persistState();
  renderTeamEmails();
  setStatus(`Added ${email} to team mail IDs.`);
}

async function removeTemplate(templateId) {
  state.templates = state.templates.filter((t) => t.id !== templateId);
  await persistState();
  renderTemplates();
  setStatus("Template removed.");
}

async function removeTeamEmail(email) {
  state.teamEmails = state.teamEmails.filter((addr) => addr !== email);
  await persistState();
  renderTeamEmails();
  setStatus("Team mail ID removed.");
}

async function handleCreateDrafts() {
  clearStatus();
  ui.createDraftsBtn.disabled = true;
  setStatus("Creating drafts...");

  const templateId = ui.templateSelect.value;
  const template = state.templates.find((t) => t.id === templateId);
  if (!template) {
    setStatus("Select a template before creating drafts.", true);
    ui.createDraftsBtn.disabled = false;
    return;
  }

  const referenceCode = (ui.referenceInput.value || "").trim();

  try {
    const token = await getRestAccessToken();
    const restBaseUrl = `${Office.context.mailbox.restUrl}/v2.0/me`;
    const messageIds = await getSelectedMessageIds();

    if (!messageIds.length) {
      throw new Error("Select at least one message in Outlook.");
    }

    for (const messageId of messageIds) {
      await createReplyAllDraft({
        messageId,
        template,
        referenceCode,
        teamEmails: state.teamEmails,
        restBaseUrl,
        token
      });
    }

    setStatus(`Draft${messageIds.length > 1 ? "s" : ""} created successfully.`);
  } catch (error) {
    console.error(error);
    setStatus(error.message || "Failed to create drafts.", true);
  } finally {
    ui.createDraftsBtn.disabled = false;
  }
}

function setStatus(message, isError = false) {
  ui.statusArea.textContent = message;
  ui.statusArea.style.color = isError ? "#a4262c" : "#0078d4";
}

function clearStatus() {
  ui.statusArea.textContent = "";
}

function generateId() {
  return `${Date.now()}-${Math.random().toString(16).slice(2)}`;
}

function escapeHtml(value = "") {
  return value
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#039;");
}

function isValidEmail(email) {
  return /^[^\s@]+@[^\s@]+\.[^\s@]+$/.test(email);
}

function getSelectedMessageIds() {
  return new Promise((resolve) => {
    const mailbox = Office.context.mailbox;
    if (typeof mailbox.getSelectedMessagesAsync === "function") {
      mailbox.getSelectedMessagesAsync((asyncResult) => {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded && asyncResult.value?.length) {
          const ids = asyncResult.value
            .map((message) => message.itemId)
            .filter(Boolean);
          if (ids.length) {
            resolve(ids);
            return;
          }
        }
        resolve(getFallbackMessageIds());
      });
      return;
    }
    resolve(getFallbackMessageIds());
  });
}

function getFallbackMessageIds() {
  const mailbox = Office.context.mailbox;
  const item = mailbox?.item;
  if (item && item.itemId) {
    return [item.itemId];
  }
  return [];
}

function getRestAccessToken() {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.getCallbackTokenAsync({ isRest: true }, (result) => {
      if (result.status === Office.AsyncResultStatus.Succeeded) {
        resolve(result.value);
      } else {
        reject(result.error || new Error("Unable to acquire REST token."));
      }
    });
  });
}

async function createReplyAllDraft({ messageId, template, referenceCode, teamEmails, restBaseUrl, token }) {
  const headers = {
    Authorization: `Bearer ${token}`,
    Accept: "application/json;odata.metadata=minimal",
    "Content-Type": "application/json"
  };

  const encodedId = encodeURIComponent(messageId);
  const messageUrl = `${restBaseUrl}/messages/${encodedId}`;
  const selectFields = "$select=Subject,CcRecipients,Body";

  const messageResponse = await fetch(`${messageUrl}?${selectFields}`, {
    method: "GET",
    headers
  });

  if (!messageResponse.ok) {
    throw new Error(`Unable to read message details (${messageResponse.status}).`);
  }

  const message = await messageResponse.json();
  const originalSubject = message.Subject || "";

  const originalCcEntries = (message.CcRecipients || [])
    .map((recipient) => {
      const address = recipient?.EmailAddress?.Address;
      if (!address) {
        return null;
      }
      return {
        "@odata.type": "#Microsoft.OutlookServices.Recipient",
        EmailAddress: {
          "@odata.type": "#Microsoft.OutlookServices.EmailAddress",
          Address: address,
          Name: recipient.EmailAddress.Name || address
        }
      };
    })
    .filter(Boolean);

  const ccSet = new Set(originalCcEntries.map((entry) => entry.EmailAddress.Address.toLowerCase()));

  const additionalCcEntries = (teamEmails || [])
    .map((addr) => (addr || "").toLowerCase())
    .filter((addr) => addr && !ccSet.has(addr))
    .map((addr) => ({
      "@odata.type": "#Microsoft.OutlookServices.Recipient",
      EmailAddress: {
        "@odata.type": "#Microsoft.OutlookServices.EmailAddress",
        Address: addr,
        Name: addr
      }
    }));

  const ccRecipients = [...originalCcEntries, ...additionalCcEntries];
  const originalBodyContent = message.Body?.Content || "";
  const combinedBody = buildReplyBody(template.body, originalBodyContent);

  const updatedSubject = referenceCode ? `${originalSubject} ${referenceCode}`.trim() : originalSubject;

  const payload = {
    Comment: "",
    Message: {
      "@odata.type": "#Microsoft.OutlookServices.Message",
      Subject: updatedSubject,
      Body: {
        "@odata.type": "#Microsoft.OutlookServices.ItemBody",
        ContentType: "HTML",
        Content: combinedBody
      },
      CcRecipients: ccRecipients
    }
  };

  const createResponse = await fetch(`${messageUrl}/createReplyAll`, {
    method: "POST",
    headers,
    body: JSON.stringify(payload)
  });

  if (!createResponse.ok) {
    const errorText = await createResponse.text();
    throw new Error(parseServiceError(errorText) || `Failed to create draft (${createResponse.status}).`);
  }
}

function buildReplyBody(templateBody, originalBody) {
  const normalizedTemplate = normalizeTemplateBody(templateBody);
  if (!originalBody) {
    return normalizedTemplate;
  }
  if (!normalizedTemplate) {
    return originalBody;
  }
  return `${normalizedTemplate}<br/><br/>${originalBody}`;
}

function normalizeTemplateBody(input) {
  if (!input) {
    return "";
  }
  const trimmed = input.trim();
  if (!/<[a-z][\s\S]*>/i.test(trimmed)) {
    return escapeHtml(trimmed).replace(/\r?\n/g, "<br/>");
  }
  return trimmed;
}

function parseServiceError(rawBody) {
  try {
    const parsed = JSON.parse(rawBody);
    return parsed?.error?.message;
  } catch (error) {
    return null;
  }
}
