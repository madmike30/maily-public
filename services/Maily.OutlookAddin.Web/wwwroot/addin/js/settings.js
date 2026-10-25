(function () {
  const fields = [
    "userName",
    "tone",
    "length",
    "language",
    "model",
    "customInstructions",
    "systemPromptOverride",
    "draftPromptOverride",
    "stylePersonalizationEnabled"
  ];

  const state = {
    response: null,
    values: {},
    savedValues: {},
    defaults: {},
    promptCatalog: {},
    fieldErrors: {},
    dirty: false,
    busy: false,
    catalogBusy: false,
    modelCatalog: null
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function setStatus(text, kind) {
    const el = byId("settingsStatus");
    el.textContent = text;
    el.className = "status";
    if (kind) {
      el.classList.add(kind);
    }
  }

  function setBusy(isBusy) {
    state.busy = isBusy;
    fields.forEach((field) => { const element = byId(field); if (element) element.disabled = isBusy; });
    [
      "reloadSettingsBtn",
      "refreshStyleBtn",
      "resetStyleBtn",
      "resetAllBtn"
    ].forEach((id) => {
      const element = byId(id);
      if (element) {
        element.disabled = isBusy;
      }
    });

    byId("saveSettingsBtn").disabled = isBusy || !state.dirty;
    if (byId("refreshModelsBtn")) byId("refreshModelsBtn").disabled = isBusy || state.catalogBusy;

    document.querySelectorAll("[data-reset-field]").forEach((button) => {
      button.disabled = isBusy;
    });

    document.querySelectorAll("[data-copy-baseline]").forEach((button) => {
      button.disabled = isBusy;
    });
  }

  function getMailboxDefaultName() {
    return window.MailyOffice.getMailboxDisplayName() || "User";
  }

  function toBool(value) {
    return !!value;
  }

  function normalizeIncomingValues(payload) {
    const values = payload && payload.values ? payload.values : payload || {};
    return {
      userName: values.userName || payload.userName || getMailboxDefaultName(),
      tone: values.tone || payload.tone || "professional",
      length: values.length || payload.length || "concise",
      language: values.language || payload.language || "English",
      model: values.model || payload.model || "",
      customInstructions: values.customInstructions || payload.customInstructions || "",
      systemPromptOverride: values.systemPromptOverride || payload.systemPromptOverride || "",
      draftPromptOverride: values.draftPromptOverride || payload.draftPromptOverride || "",
      stylePersonalizationEnabled: values.stylePersonalizationEnabled !== false && payload.stylePersonalizationEnabled !== false
    };
  }

  function normalizeDefaults(payload) {
    const defaults = payload && payload.defaults ? payload.defaults : {};
    return {
      userName: defaults.userName || getMailboxDefaultName(),
      tone: defaults.tone || "professional",
      length: defaults.length || "concise",
      language: defaults.language || "English",
      model: defaults.model || "",
      customInstructions: defaults.customInstructions || "",
      systemPromptOverride: defaults.systemPromptOverride || "",
      draftPromptOverride: defaults.draftPromptOverride || "",
      stylePersonalizationEnabled: defaults.stylePersonalizationEnabled !== false
    };
  }

  function valuesEqual(left, right) {
    return fields.every((field) => {
      if (field === "stylePersonalizationEnabled") {
        return toBool(left[field]) === toBool(right[field]);
      }

      return String(left[field] || "") === String(right[field] || "");
    });
  }

  function readForm() {
    return {
      userName: byId("userName").value.trim() || getMailboxDefaultName(),
      tone: byId("tone").value,
      length: byId("length").value,
      language: byId("language").value,
      model: byId("model").value.trim(),
      customInstructions: byId("customInstructions").value,
      systemPromptOverride: byId("systemPromptOverride").value,
      draftPromptOverride: byId("draftPromptOverride").value,
      stylePersonalizationEnabled: !!byId("stylePersonalizationEnabled").checked
    };
  }

  function setDefaultLabel(field, value) {
    const el = byId(`default-${field}`);
    if (!el) {
      return;
    }

    const empty = !value || String(value).toLowerCase() === "blank";
    const normalized = value === true ? "On" : value === false ? "Off" : (empty ? (field === "model" ? "Automatic" : "None") : value);
    el.textContent = `Default: ${normalized}`;
  }

  function formatTimestamp(value) {
    if (!value) {
      return "";
    }

    const parsed = new Date(value);
    if (Number.isNaN(parsed.getTime())) {
      return value;
    }

    return parsed.toLocaleString([], {
      year: "numeric",
      month: "short",
      day: "numeric",
      hour: "numeric",
      minute: "2-digit"
    });
  }

  function normalizeWhitespace(value) {
    return String(value || "").replace(/\s+/g, " ").trim();
  }

  function humanizeStyleError(message) {
    const normalized = normalizeWhitespace(message);
    if (!normalized) {
      return "";
    }

    if (/could not be converted to System\.String/i.test(normalized)) {
      return "Maily could not finish learning the style profile from Sent Items. Try Refresh style again. If it keeps failing, check the local add-in logs.";
    }

    if (/At least two sent emails are required/i.test(normalized)) {
      return "Maily needs at least two sent emails with readable body text before it can learn a style profile.";
    }

    return normalized;
  }

  function buildStyleStateChips(style) {
    const chips = [];
    if ((style && style.sampleCount) > 0) {
      chips.push(`${style.sampleCount} sample${style.sampleCount === 1 ? "" : "s"}`);
    }

    if (style && style.lastRefreshedUtc) {
      chips.push(`Last refresh ${formatTimestamp(style.lastRefreshedUtc)}`);
    }

    if (style && style.lastRefreshStatus) {
      chips.push(style.lastRefreshStatus === "ok" ? "Ready to use" : `Status: ${style.lastRefreshStatus}`);
    }

    return chips;
  }

  function buildStyleProfileMeta(style) {
    if (!style || !style.summary) {
      return "No learned style yet.";
    }

    const sampleLabel = (style.sampleCount || 0) > 0
      ? `Learned from ${style.sampleCount} sent email${style.sampleCount === 1 ? "" : "s"}`
      : "Learned style is available";
    const refreshLabel = style.lastRefreshedUtc ? ` on ${formatTimestamp(style.lastRefreshedUtc)}` : "";
    if (style.enabled === false) {
      return `${sampleLabel}${refreshLabel}. The learned style is saved, but it is currently turned off.`;
    }

    return `${sampleLabel}${refreshLabel}. This profile is active for reply generation.`;
  }

  function buildStyleRefreshMeta(style) {
    if (!style) {
      return "No refresh history yet.";
    }

    const friendlyError = humanizeStyleError(style.lastRefreshError);
    if (friendlyError) {
      return friendlyError;
    }

    if (style.lastRefreshStatus === "ok") {
      return "These are the reusable rules Maily will try to follow when learned style is enabled.";
    }

    if (style.lastRefreshStatus) {
      return `Last refresh status: ${style.lastRefreshStatus}.`;
    }

    return "No refresh history yet.";
  }

  function buildStyleSteeringState() {
    const rawInstructions = byId("customInstructions").value || "";
    const normalized = normalizeWhitespace(rawInstructions);
    if (!normalized) {
      return {
        hint: "Add a short reusable note if you want to nudge the learned style without replacing it.",
        preview: "Examples: \"be slightly warmer with customers\", \"keep commercial replies crisp\", or \"prefer short bullets for action items.\""
      };
    }

    return {
      hint: "Saved reusable guidance will be layered on top of the learned style.",
      preview: normalized
    };
  }

  function renderStyleSteeringState() {
    const steering = buildStyleSteeringState();
    const hint = byId("styleSteeringHint");
    const preview = byId("styleSteeringPreview");
    if (hint) {
      hint.textContent = steering.hint;
    }
    if (preview) {
      preview.textContent = steering.preview;
      window.MailyOffice.applyDirectionToElement(preview, steering.preview);
      preview.classList.toggle("muted-copy", !normalizeWhitespace(byId("customInstructions").value || ""));
    }
  }

  function renderOptions(payload) {
    const options = payload && payload.options ? payload.options : {};
    renderSelect("tone", options.tone || ["professional", "concise", "warm", "executive summary"]);
    renderSelect("length", options.length || ["concise", "standard", "detailed"]);
    renderSelect("language", options.language || ["English", "Hebrew", "Same as thread"]);
    renderModelSelect(state.savedValues.model || "");
  }

  function renderModelSelect(currentValue) {
    const select = byId("model");
    if (!select) return;
    const current = typeof currentValue === "string" ? currentValue : select.value;
    const catalog = state.modelCatalog;
    const models = catalog && Array.isArray(catalog.models) ? catalog.models : [];
    const options = [];
    const addOption = (value, label) => {
      const option = document.createElement("option");
      option.value = value;
      option.textContent = label;
      options.push(option);
    };
    addOption("", "Automatic (fast)");
    if (current && !models.some((model) => model.id === current)) addOption(current, `${current} (saved; not listed)`);
    const seen = new Set();
    models.forEach((model) => {
      if (!model || typeof model.id !== "string" || !model.id || seen.has(model.id)) return;
      seen.add(model.id);
      addOption(model.id, model.displayName || model.id);
    });
    select.replaceChildren(...options);
    select.value = current;
  }

  async function loadModels() {
    if (state.catalogBusy || state.busy) return;
    state.catalogBusy = true;
    byId("refreshModelsBtn").disabled = true;
    byId("modelStatus").textContent = "Checking models for your connection...";
    try {
      const catalog = await window.MailyOffice.callJson("/api/outlook/models");
      if (!catalog || catalog.status !== "ready" || !Array.isArray(catalog.models) || catalog.models.length === 0) throw new Error("catalog-unavailable");
      state.modelCatalog = catalog;
      renderModelSelect();
      const recommended = catalog.models.find((model) => model.id === catalog.recommendedModel);
      byId("modelStatus").textContent = (recommended ? `Automatic currently uses ${recommended.displayName || recommended.id}. ` : "Automatic uses the connection's default. ") + "Availability and speed can change with your account. The list refreshes at most every 5 minutes.";
    } catch (_) {
      byId("modelStatus").textContent = "Could not refresh models. Your selection is kept; try again in 30 seconds.";
    } finally {
      state.catalogBusy = false;
      byId("refreshModelsBtn").disabled = state.busy;
    }
  }
  function renderSelect(id, values) {
    const select = byId(id);
    const current = select.value;
    select.innerHTML = values.map((value) => `<option value="${escapeHtml(value)}">${escapeHtml(toTitle(value))}</option>`).join("");
    if (values.includes(current)) {
      select.value = current;
    }
  }

  function toTitle(value) {
    if (!value) {
      return "";
    }

    if (value === "Same as thread") {
      return value;
    }

    return value
      .split(" ")
      .map((segment) => segment.charAt(0).toUpperCase() + segment.slice(1))
      .join(" ");
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function applyForm(values) {
    byId("profileId").value = state.response && state.response.profileId ? state.response.profileId : "default";
    byId("userName").value = values.userName || getMailboxDefaultName();
    byId("tone").value = values.tone || "professional";
    byId("length").value = values.length || "concise";
    byId("language").value = values.language || "English";
    byId("model").value = values.model || "";
    byId("customInstructions").value = values.customInstructions || "";
    byId("systemPromptOverride").value = values.systemPromptOverride || "";
    byId("draftPromptOverride").value = values.draftPromptOverride || "";
    byId("stylePersonalizationEnabled").checked = values.stylePersonalizationEnabled !== false;
    ["userName", "model", "customInstructions", "systemPromptOverride", "draftPromptOverride"].forEach((field) => {
      const element = byId(field);
      window.MailyOffice.applyDirectionToElement(element, element.value);
    });
  }

  function applyMetadata(payload) {
    const metadata = payload && payload.fieldMetadata ? payload.fieldMetadata : {};
    Object.keys(metadata).forEach((field) => {
      const help = byId(`help-${field}`);
      if (help) {
        help.textContent = metadata[field].description || "";
      }
      if (metadata[field].isAdvanced) {
        const card = byId(`card-${field}`);
        if (card) {
          card.dataset.advanced = "true";
        }
      }
      setDefaultLabel(field, metadata[field].recommended || (state.defaults && state.defaults[field]));
    });

    const warnings = payload && payload.advancedWarnings ? payload.advancedWarnings : [];
    byId("advancedWarnings").innerHTML = warnings.map((warning) => `<li>${escapeHtml(warning)}</li>`).join("");
  }

  function applyPromptCatalog(payload) {
    const promptCatalog = payload && payload.promptCatalog ? payload.promptCatalog : {};
    state.promptCatalog = {
      systemPromptBaseline: promptCatalog.systemPromptBaseline || "",
      draftPromptBaseline: promptCatalog.draftPromptBaseline || "",
      systemPromptSourceLabel: promptCatalog.systemPromptSourceLabel || "Current shipped prompt",
      draftPromptSourceLabel: promptCatalog.draftPromptSourceLabel || "Current shipped prompt"
    };

    const baselineSystemPromptSource = byId("baselineSystemPromptSource");
    const baselineDraftPromptSource = byId("baselineDraftPromptSource");
    const baselineSystemPrompt = byId("baselineSystemPrompt");
    const baselineDraftPrompt = byId("baselineDraftPrompt");
    if (baselineSystemPromptSource) {
      baselineSystemPromptSource.textContent = state.promptCatalog.systemPromptSourceLabel;
    }
    if (baselineDraftPromptSource) {
      baselineDraftPromptSource.textContent = state.promptCatalog.draftPromptSourceLabel;
    }
    if (baselineSystemPrompt) {
      baselineSystemPrompt.value = state.promptCatalog.systemPromptBaseline;
      window.MailyOffice.applyDirectionToElement(baselineSystemPrompt, state.promptCatalog.systemPromptBaseline);
    }
    if (baselineDraftPrompt) {
      baselineDraftPrompt.value = state.promptCatalog.draftPromptBaseline;
      window.MailyOffice.applyDirectionToElement(baselineDraftPrompt, state.promptCatalog.draftPromptBaseline);
    }
  }

  function applyStyleState(payload) {
    const style = payload && payload.style ? payload.style : payload || {};
    const styleStateChips = byId("styleStateChips");
    if (styleStateChips) {
      const chips = buildStyleStateChips(style);
      styleStateChips.innerHTML = chips.length === 0
        ? ""
        : chips.map((chip) => `<span class="chip-soft">${escapeHtml(chip)}</span>`).join("");
    }

    byId("styleProfileMeta").textContent = buildStyleProfileMeta(style);
    byId("styleProfileSummary").textContent = style.summary || "Refresh style from Sent Items to build a reusable style profile.";
    byId("analyzedStyle").textContent = style.analyzedStyle || "Maily will store reusable guidance here after style analysis.";
    byId("styleRefreshMeta").textContent = buildStyleRefreshMeta(style);
    window.MailyOffice.applyDirectionToElement(byId("styleProfileSummary"), byId("styleProfileSummary").textContent);
    window.MailyOffice.applyDirectionToElement(byId("analyzedStyle"), byId("analyzedStyle").textContent);
    renderStyleSteeringState();
  }

  function renderFieldErrors() {
    fields.forEach((field) => {
      const error = state.fieldErrors[field] || "";
      const errorEl = byId(`error-${field}`);
      const card = byId(`card-${field}`);
      if (errorEl) {
        errorEl.textContent = error;
      }
      if (card) {
        card.classList.toggle("has-error", !!error);
      }
    });
  }

  function renderDirtyState() {
    state.values = readForm();
    state.dirty = !valuesEqual(state.values, state.savedValues);
    fields.forEach((field) => {
      const card = byId(`card-${field}`);
      if (!card) {
        return;
      }

      const current = state.values[field];
      const saved = state.savedValues[field];
      const isDirty = field === "stylePersonalizationEnabled"
        ? toBool(current) !== toBool(saved)
        : String(current || "") !== String(saved || "");
      card.classList.toggle("has-dirty", isDirty);
    });

    document.querySelectorAll("[data-reset-field]").forEach(button => {
      const field = button.getAttribute("data-reset-field");
      const current = state.values[field];
      const baseline = state.defaults[field];
      button.hidden = field === "stylePersonalizationEnabled"
        ? toBool(current) === toBool(baseline)
        : String(current || "") === String(baseline || "");
      button.setAttribute("aria-label", `Reset ${field.replace(/([a-z])([A-Z])/g, "$1 $2")} to default`);
    });
    byId("saveSettingsBtn").disabled = state.busy || !state.dirty;
    byId("settingsSaveNote").textContent = state.dirty
      ? "Unsaved changes"
      : "Settings saved";
  }

  function applyPayload(payload) {
    const edits = {};
    if (state.dirty) {
      const current = readForm();
      fields.forEach((field) => { if (String(current[field]) !== String(state.savedValues[field])) edits[field] = current[field]; });
    }
    state.response = payload || {};
    state.defaults = normalizeDefaults(payload || {});
    state.savedValues = normalizeIncomingValues(payload || {});
    state.fieldErrors = payload && payload.fieldErrors ? payload.fieldErrors : {};
    renderOptions(payload || {});
    const values = Object.assign({}, state.savedValues, edits);
    renderModelSelect(values.model);
    applyForm(values);
    applyMetadata(payload || {});
    applyPromptCatalog(payload || {});
    applyStyleState(payload && payload.style ? payload.style : payload || {});
    renderFieldErrors();
    renderDirtyState();
  }

  function resetField(field) {
    const defaults = state.defaults || {};
    if (field === "stylePersonalizationEnabled") {
      byId("stylePersonalizationEnabled").checked = defaults.stylePersonalizationEnabled !== false;
    } else {
      const element = byId(field);
      if (element) {
        element.value = defaults[field] || "";
        if (element.tagName === "TEXTAREA" || element.tagName === "INPUT") {
          window.MailyOffice.applyDirectionToElement(element, element.value);
        }
      }
    }

    state.fieldErrors[field] = "";
    renderFieldErrors();
    renderDirtyState();
  }

  function resetAllFields() {
    fields.forEach(resetField);
  }

  async function loadSettings() {
    setBusy(true);
    setStatus("Loading settings...", "");
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
      const payload = await window.MailyOffice.callJson(`/api/outlook/settings?profileId=default&mailboxAddress=${mailboxAddress}`);
      applyPayload(payload || {});
      setStatus("Settings loaded.", "ok");
      await window.MailyOffice.logClientEvent("settings-pane-load", "ok", "Loaded settings values.", "settings");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Could not load settings.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("settings-pane-load", "error", message, "settings");
    } finally {
      setBusy(false);
    }
  }

  async function saveSettings() {
    setBusy(true);
    setStatus("Saving settings...", "");
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const payload = {
        profileId: byId("profileId").value.trim() || "default",
        values: readForm()
      };
      const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
      const response = await window.MailyOffice.callJson(`/api/outlook/settings?mailboxAddress=${mailboxAddress}`, {
        method: "PUT",
        body: JSON.stringify(payload)
      });
      state.dirty = false;
      applyPayload(response || {});
      setStatus("Settings saved.", "ok");
      await window.MailyOffice.logClientEvent("settings-save", "ok", "Saved settings values.", "settings");
    } catch (error) {
      let message = error instanceof Error ? error.message : "Could not save settings.";
      try {
        const parsed = JSON.parse(message);
        if (parsed && parsed.fieldErrors) {
          state.fieldErrors = parsed.fieldErrors;
          renderFieldErrors();
          renderDirtyState();
          message = parsed.error || "Please review the highlighted settings.";
        }
      } catch (_) {
      }
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("settings-save", "error", message, "settings");
    } finally {
      setBusy(false);
    }
  }

  async function refreshStyle() {
    setBusy(true);
    setStatus("Refreshing style from Sent Items...", "");
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
      await window.MailyOffice.callJson(`/api/outlook/style-profile/refresh?profileId=default&mailboxAddress=${mailboxAddress}`, {
        method: "POST"
      });
      await loadSettings();
      setStatus("Style profile refreshed.", "ok");
      await window.MailyOffice.logClientEvent("style-refresh", "ok", "Refreshed style profile from Sent Items.", "settings");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Could not refresh style.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("style-refresh", "error", message, "settings");
    } finally {
      setBusy(false);
    }
  }

  function wireInputTracking() {
    fields.forEach((field) => {
      const element = byId(field);
      if (!element) {
        return;
      }

      const eventName = element.tagName === "SELECT" || element.type === "checkbox" ? "change" : "input";
      element.addEventListener(eventName, () => {
        state.fieldErrors[field] = "";
        if (element.tagName === "TEXTAREA" || element.tagName === "INPUT") {
          window.MailyOffice.applyDirectionToElement(element, element.type === "checkbox" ? "" : element.value);
        }
        if (field === "customInstructions") {
          renderStyleSteeringState();
        }
        renderFieldErrors();
        renderDirtyState();
      });
    });
  }

  function wireActions() {
    byId("saveSettingsBtn").addEventListener("click", saveSettings);
    byId("reloadSettingsBtn").addEventListener("click", loadSettings);
    byId("refreshStyleBtn").addEventListener("click", refreshStyle);
    byId("refreshModelsBtn").addEventListener("click", loadModels);
    byId("openStyleGuidanceBtn").addEventListener("click", () => {
      const advancedSection = byId("advancedSection");
      const customInstructions = byId("customInstructions");
      if (advancedSection) {
        advancedSection.open = true;
      }
      if (customInstructions) {
        customInstructions.focus();
        customInstructions.scrollIntoView({ behavior: "smooth", block: "center" });
      }
      setStatus("Reusable guidance is open below. Use it to nudge Maily's learned style without replacing the baseline.", "ok");
    });
    byId("resetAllBtn").addEventListener("click", () => {
      resetAllFields();
      setStatus("All form values were restored to the recommended defaults. Save to persist them.", "ok");
    });
    document.querySelectorAll("[data-reset-field]").forEach((button) => {
      button.addEventListener("click", () => {
        resetField(button.getAttribute("data-reset-field"));
      });
    });
    document.querySelectorAll("[data-copy-baseline]").forEach((button) => {
      button.addEventListener("click", () => {
        const field = button.getAttribute("data-copy-baseline");
        if (!field) {
          return;
        }

        if (field === "systemPromptOverride") {
          byId(field).value = state.promptCatalog.systemPromptBaseline || "";
        } else if (field === "draftPromptOverride") {
          byId(field).value = state.promptCatalog.draftPromptBaseline || "";
        }

        window.MailyOffice.applyDirectionToElement(byId(field), byId(field).value);
        state.fieldErrors[field] = "";
        renderFieldErrors();
        if (field === "customInstructions") {
          renderStyleSteeringState();
        }
        renderDirtyState();
      });
    });
    wireInputTracking();
  }

  function getOfficeBridge() {
    if (!window.MailyOffice) {
      throw new Error("The Maily Outlook bootstrap script did not load. Close and reopen the pane.");
    }

    return window.MailyOffice;
  }

  async function start() {
    const office = getOfficeBridge();
    await office.officeReady();
    if (office.pruneLegacyPaneButtons) {
      office.pruneLegacyPaneButtons();
    }
    await office.logClientEvent("settings-pane-open", "ok", "Settings pane opened.", "settings");
    wireActions();
    byId("advancedSection").open = false;
    setStatus("Checking Maily setup...", "");
    await office.ensureBackendAvailable();
    await loadSettings();
    await loadModels();
  }

  start().catch((error) => {
    const message = error instanceof Error ? error.message : "Settings pane failed to start.";
    setStatus(message, "error");
    if (window.MailyOffice && window.MailyOffice.logClientEvent) {
      window.MailyOffice.logClientEvent("settings-pane-start-failed", "error", message, "settings");
    }
  });
})();
