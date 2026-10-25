(function () {
  const state = {
    draft: null,
    pending: null,
    busy: false,
    composeState: null,
    draftContext: null,
    selectedOutputLanguage: "Auto",
    languageChosenByUser: false,
    defaults: {
      profileId: "default",
      tone: "professional",
      length: "concise",
      language: "English"
    }
  };

  function mergeObjects(base, overrides) {
    const result = {};
    Object.keys(base || {}).forEach((key) => {
      result[key] = base[key];
    });
    Object.keys(overrides || {}).forEach((key) => {
      result[key] = overrides[key];
    });
    return result;
  }

  function byId(id) {
    return document.getElementById(id);
  }

  function setStatus(text, kind) {
    const el = byId("status");
    el.textContent = text;
    el.className = "status";
    if (kind) {
      el.classList.add(kind);
    }
  }

  function setBusy(isBusy) {
    state.busy = isBusy;
    ["generateBtn", "insertBtn", "shorterBtn", "clearerBtn", "warmerBtn", "replyLanguageEnglishBtn", "replyLanguageHebrewBtn"].forEach(id => { byId(id).disabled = isBusy; });
    const automaticButton = byId("replyLanguageAutoBtn");
    if (automaticButton) automaticButton.disabled = isBusy;
    byId("generateBtn").classList.toggle("primary", !state.pending);
    byId("generateBtn").textContent = state.draft ? "Try another reply" : "Prepare reply";
    byId("applyBtn").classList.toggle("primary", !!state.pending);
    byId("applyBtn").disabled = isBusy || !state.pending;
    byId("undoBtn").disabled = isBusy || !window.MailyOffice.canUndoGeneratedDraft || !window.MailyOffice.canUndoGeneratedDraft();
    byId("undoBtn").hidden = byId("undoBtn").disabled;
  }
  function normalizeLanguage(value) {
    const language = String(value || "").trim().toLowerCase();
    return language === "hebrew" ? "Hebrew" : language === "english" ? "English" : "Auto";
  }

  function renderLanguageButtons() {
    const summary = byId("replyLanguageSummary");
    if (summary) summary.textContent = "Language: " + (state.languageChosenByUser ? state.selectedOutputLanguage : "Automatic");
    const auto = byId("replyLanguageAutoBtn");
    if (auto) {
      auto.classList.toggle("is-selected", !state.languageChosenByUser);
      auto.setAttribute("aria-pressed", state.languageChosenByUser ? "false" : "true");
    }
    ["English", "Hebrew"].forEach((language) => {
      const button = byId(`replyLanguage${language}Btn`);
      const isSelected = state.languageChosenByUser && state.selectedOutputLanguage === language;
      button.classList.toggle("is-selected", isSelected);
      button.setAttribute("aria-pressed", isSelected ? "true" : "false");
    });
  }

  function setOutputLanguage(language, options) {
    state.selectedOutputLanguage = normalizeLanguage(language);
    if (options && options.userSelected) {
      state.languageChosenByUser = true;
    }
    renderLanguageButtons();
  }

  function syncLanguageSelection(snapshot) {
    if (state.languageChosenByUser) {
      return;
    }

    const composeState = (snapshot && snapshot.composeState) || state.composeState || {};
    const request = snapshot && snapshot.request ? snapshot.request : {};
    const compose = request.compose || {};
    const inferred = window.MailyOffice.detectPreferredLanguage(
      composeState.userDraftText || "",
      compose.bodyText || "",
      compose.bodyHtml || "",
      compose.subject || "");
    setOutputLanguage(inferred);
  }

  function applyLanguageOverride(request) {
    if (!request || !request.operation) {
      return;
    }

    if (state.languageChosenByUser) {
      request.operation.outputLanguageOverride = state.selectedOutputLanguage;
      return;
    }

    delete request.operation.outputLanguageOverride;
  }

  function renderSeed() {
    const panel = byId("seedPanel");
    const seedText = state.composeState && state.composeState.userDraftText ? state.composeState.userDraftText : "";
    const normalized = seedText.trim();
    const seedMeta = byId("seedMeta");
    const seedPreview = byId("seedText");

    if (!normalized) {
      panel.classList.add("seed-panel-empty");
      seedMeta.textContent = "No typed reply text detected yet.";
      seedPreview.textContent = "Maily will use the text you type in the Outlook reply area as the starting point for generation.";
      window.MailyOffice.applyDirectionToElement(seedPreview, seedPreview.textContent);
      return;
    }

    panel.classList.remove("seed-panel-empty");
    seedMeta.textContent = `Detected ${normalized.length} characters from the current Outlook reply draft.`;
    seedPreview.textContent = normalized;
    window.MailyOffice.applyDirectionToElement(seedPreview, normalized);
  }

  function renderDraft() {
    const text = state.draft && state.draft.draft && state.draft.draft.text ? state.draft.draft.text : "No generated draft yet.";
    const model = state.draft && state.draft.meta && state.draft.meta.model ? state.draft.meta.model : "-";
    const requestId = state.draft && state.draft.meta && state.draft.meta.requestId ? state.draft.meta.requestId : "-";
    byId("draftText").textContent = text;
    window.MailyOffice.applyDirectionToElement(byId("draftText"), text);
    byId("draftMeta").textContent = `Model: ${model} | Request: ${requestId}`;
  }

  async function loadDefaults() {
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
      const payload = await window.MailyOffice.callJson(`/api/outlook/settings?profileId=default&mailboxAddress=${mailboxAddress}`);
      const values = payload && payload.values ? payload.values : payload;
      state.defaults = {
        profileId: payload.profileId || "default",
        userName: values.userName || payload.userName || window.MailyOffice.getMailboxDisplayName() || "User",
        tone: values.tone || payload.tone || "professional",
        length: values.length || payload.length || "concise",
        language: values.language || payload.language || "English",
        model: values.model || payload.model || null
      };
      await window.MailyOffice.logClientEvent("settings-load", "ok", "Loaded compose defaults.", "taskpane");
    } catch (_) {
      state.defaults = {
        profileId: "default",
        userName: window.MailyOffice.getMailboxDisplayName() || "User",
        tone: "professional",
        length: "concise",
        language: "English",
        model: null
      };
      await window.MailyOffice.logClientEvent("settings-load", "fallback", "Using built-in compose defaults.", "taskpane");
    }
  }

  async function generate(mode, options) {
    if (state.busy) return;
    state.pending = null;
    setBusy(true);
    setStatus("Reading your current draft…", "");
    try {
      const generationContext = window.MailyOffice.captureComposeContext();
      await window.MailyOffice.ensureBackendAvailable();
      window.MailyOffice.assertComposeContext(generationContext);
      const snapshot = await window.MailyOffice.getComposeSnapshot(state.defaults, mode);
      window.MailyOffice.assertComposeContext(generationContext);
      state.composeState = snapshot.composeState;
      syncLanguageSelection(snapshot);
      renderSeed();
      if (snapshot.composeState.userDraftText && snapshot.composeState.canRewriteUserDraft === false) {
        throw new Error("Maily could not safely separate your reply from the signature or email history. Your draft has not been changed.");
      }
      const request = snapshot.request;
      applyLanguageOverride(request);
      const refinement = options && options.refinement;
      if (refinement) {
        request.operation.userInstruction = [request.operation.userInstruction || "", `Make the existing reply ${refinement}. Preserve its language, meaning, names, numbers, dates and commitments. Do not invent facts or commitments. Treat the current Outlook draft as authoritative.`].filter(Boolean).join("\n");
      }
      setStatus("Preparing reply preview...", "");
      const payload = await window.MailyOffice.callJson("/api/outlook/generate-reply", { method: "POST", body: JSON.stringify(request) });
      window.MailyOffice.assertComposeContext(generationContext);
      if (!payload || !payload.draft || !String(payload.draft.text || "").trim()) throw new Error("No reply was generated. Please retry.");
      state.draftContext = generationContext;
      state.draft = payload;
      state.pending = { snapshot, context: generationContext };
      byId("seedPanel").open = !!snapshot.composeState.userDraftText;
      renderDraft();
      setStatus("Ready to review. Nothing has changed in Outlook.", "ok");
      await window.MailyOffice.logClientEvent("generate-complete", "ok", "Reply preview ready; no draft mutation.", "taskpane");
    } catch (error) {
      state.pending = null;
      if (state.draftContext) {
        try { window.MailyOffice.assertComposeContext(state.draftContext); }
        catch (_) { state.draft = null; state.draftContext = null; state.composeState = null; renderDraft(); renderSeed(); }
      }
      const message = error instanceof Error ? error.message : "Failed to generate reply.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("generate-failed", "error", message, "taskpane");
    } finally { setBusy(false); }
  }

  async function applyReply() {
    if (state.busy || !state.pending || !state.draft) return;
    const pending = state.pending;
    // Consume before the mutation, including ambiguous/partially successful writes.
    state.pending = null;
    setBusy(true);
    try {
      window.MailyOffice.assertComposeContext(pending.context);
      const snapshot = pending.snapshot;
      const compose = snapshot.composeState;
      const rewrite = !!compose.userDraftText;
      await window.MailyOffice.insertGeneratedDraft(state.draft.draft, {
        expectedContext: pending.context,
        replaceExisting: !rewrite && !!compose.hasGeneratedBlock,
        rewriteComposeBody: rewrite,
        composeDefaults: state.defaults,
        subject: snapshot.request.compose.subject,
        seedText: compose.userDraftText || "",
        composeType: compose.composeType || "unknown",
        expectedBody: compose.body,
        expectedBodyType: compose.bodyType,
        originalUserDraftSegment: compose.userDraftSegment || "",
        preservedSuffix: compose.preservedSuffix || ""
      });
      const canUndo = window.MailyOffice.canUndoGeneratedDraft && window.MailyOffice.canUndoGeneratedDraft();
      setStatus("Reply updated in Outlook. Review it before sending." + (canUndo ? " Undo is available while the draft stays unchanged." : ""), "ok");
    } catch (error) {
      state.draft = null;
      state.draftContext = null;
      renderDraft();
      setStatus(error instanceof Error ? error.message : "Could not apply the reply. Prepare a fresh preview.", "error");
    } finally { setBusy(false); }
  }

  async function undoReply() {
    if (state.busy || !window.MailyOffice.undoGeneratedDraft) return;
    state.pending = null;
    setBusy(true);
    try {
      await window.MailyOffice.undoGeneratedDraft();
      state.draft = null;
      state.draftContext = null;
      const snapshot = await window.MailyOffice.getComposeSnapshot(state.defaults, "generate-initial");
      state.composeState = snapshot.composeState;
      renderSeed(); renderDraft();
      setStatus("Your previous reply was restored. Nothing was sent.", "ok");
    } catch (error) { setStatus(error instanceof Error ? error.message : "Could not undo. Your newer edits have been kept.", "error"); }
    finally { setBusy(false); }
  }

  async function insertAgain() {
    // Compatibility with older panes: always prepare a fresh preview, never repeat a write.
    await generate(getRequestedGenerateMode());
  }
  function getRequestedGenerateMode() {
    return state.draft ? "regenerate" : "generate-initial";
  }

  function wireActions() {
    const auto = byId("replyLanguageAutoBtn");
    if (auto) auto.addEventListener("click", () => {
      if (state.busy) return;
      state.pending = null;
      state.languageChosenByUser = false;
      state.selectedOutputLanguage = "Auto";
      renderLanguageButtons();
      setBusy(false);
      setStatus("Language will follow your draft or email. Prepare a new reply to use it.", "ok");
    });
    byId("generateBtn").addEventListener("click", () => generate(getRequestedGenerateMode()));
    byId("insertBtn").addEventListener("click", insertAgain);
    byId("applyBtn").addEventListener("click", applyReply);
    byId("undoBtn").addEventListener("click", undoReply);
    ["shorter", "clearer", "warmer"].forEach(refinement => {
      byId(`${refinement}Btn`).addEventListener("click", () => generate(getRequestedGenerateMode(), { refinement }));
    });
    ["English", "Hebrew"].forEach((language) => {
      byId(`replyLanguage${language}Btn`).addEventListener("click", () => {
        if (state.busy) return;
        state.pending = null;
        setBusy(false);
        setOutputLanguage(language, { userSelected: true });
        setStatus(`Reply language set to ${language}. Click Prepare reply to refresh the draft.`, "ok");
      });
    });
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
    await office.logClientEvent("taskpane-open", "ok", "Task pane opened.", "taskpane");
    wireActions();
    renderDraft();
    renderSeed();
    renderLanguageButtons();
    setStatus("Checking Maily setup...", "");
    await office.ensureBackendAvailable();
    await loadDefaults();
    const snapshot = await office.getComposeSnapshot(mergeObjects(state.defaults, {
      currentDraft: null
    }), "generate-initial");
    state.composeState = snapshot.composeState;
    syncLanguageSelection(snapshot);
    renderSeed();
    const composeBodyText = snapshot.request && snapshot.request.compose && snapshot.request.compose.bodyText
      ? snapshot.request.compose.bodyText.trim()
      : "";
    const subjectText = snapshot.request && snapshot.request.compose && snapshot.request.compose.subject
      ? snapshot.request.compose.subject.trim()
      : "";
    const hasMeaningfulComposeContext =
      !!(composeBodyText || subjectText || (snapshot.composeState && snapshot.composeState.userDraftText));

    if (!hasMeaningfulComposeContext) {
      setStatus("No draft text or thread context is available yet. Type your reply, then click Prepare reply when you're ready.", "");
      return;
    }

    await generate(
      snapshot.composeState && snapshot.composeState.hasGeneratedBlock ? "regenerate" : "generate-initial",
      { autoHidePaneAfterInsert: true });
  }

  start().catch((error) => {
    const message = error instanceof Error ? error.message : "Task pane failed to start.";
    setStatus(message, "error");
    if (window.MailyOffice && window.MailyOffice.logClientEvent) {
      window.MailyOffice.logClientEvent("taskpane-start-failed", "error", message, "taskpane");
    }
  });
})();
