(function () {
  const state = {
    selectedOutputLanguage: "Auto",
    languageChosenByUser: false,
    defaults: {
      profileId: "default",
      tone: "professional",
      length: "concise",
      language: "English"
    }
  };
  let replyFlow = null;
  let summaryBound = false;
  let replyReady = false;

  function byId(id) {
    return document.getElementById(id);
  }

  function getSurface() {
    const params = new URLSearchParams(window.location.search || "");
    return (params.get("surface") || "compose").toLowerCase() === "read" ? "read" : "compose";
  }

  function setStatus(text, kind) {
    const el = byId("summaryStatus");
    el.textContent = text;
    el.className = "status";
    if (kind) {
      el.classList.add(kind);
    }
  }

  function setBusy(isBusy) {
    const auto = byId("summaryLanguageAutoBtn");
    if (auto) auto.disabled = isBusy;
    ["summarizeBtn", "copySummaryBtn"].forEach((id) => {
      byId(id).disabled = isBusy;
    });
    byId("prepareSummaryReplyBtn").disabled = isBusy || !summaryBound;
    byId("applySummaryReplyBtn").disabled = isBusy || !replyReady;
    byId("prepareSummaryReplyBtn").classList.toggle("primary", !replyReady);
    ["English", "Hebrew"].forEach(language => { byId(`summaryLanguage${language}Btn`).disabled = isBusy; });
    byId("replyActionChoices").querySelectorAll("input").forEach(input => { input.disabled = isBusy; });
  }

  function normalizeLanguage(value) {
    const language = String(value || "").trim().toLowerCase();
    return language === "hebrew" ? "Hebrew" : language === "english" ? "English" : "Auto";
  }

  function renderLanguageButtons() {
    const summary = byId("summaryLanguageSummary");
    if (summary) summary.textContent = "Language: " + (state.languageChosenByUser ? state.selectedOutputLanguage : "Automatic");
    const auto = byId("summaryLanguageAutoBtn");
    if (auto) {
      auto.classList.toggle("is-selected", !state.languageChosenByUser);
      auto.setAttribute("aria-pressed", state.languageChosenByUser ? "false" : "true");
    }
    ["English", "Hebrew"].forEach((language) => {
      const button = byId(`summaryLanguage${language}Btn`);
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

  function syncLanguageSelection(payload) {
    if (state.languageChosenByUser) {
      return;
    }

    const inferred = window.MailyOffice.detectPreferredLanguage(
      payload && payload.bodyText ? payload.bodyText : "",
      payload && payload.subject ? payload.subject : "");
    setOutputLanguage(inferred);
  }

  function applyLanguageOverride(payload) {
    if (!payload) {
      return;
    }

    if (state.languageChosenByUser) {
      payload.outputLanguageOverride = state.selectedOutputLanguage;
      return;
    }

    delete payload.outputLanguageOverride;
  }

  function toDisplayText(value) {
    if (typeof value === "string") {
      return value;
    }
    if (Array.isArray(value)) {
      const items = value.map((item) => toDisplayText(item)).filter(Boolean);
      if (items.length === 0) {
        return "";
      }
      if (items.length === 1) {
        return items[0];
      }
      return items.map((item) => `- ${item}`).join("\n");
    }
    if (value && typeof value === "object") {
      const lines = Object.entries(value)
        .map(([key, entry]) => {
          const text = toDisplayText(entry);
          if (!text) {
            return "";
          }
          const label = key.replace(/([a-z])([A-Z])/g, "$1 $2");
          return `${label.charAt(0).toUpperCase()}${label.slice(1)}: ${text}`;
        })
        .filter(Boolean);
      return lines.join("\n");
    }
    if (value === null || typeof value === "undefined") {
      return "";
    }
    return String(value);
  }

  function renderSummary(payload) {
    const sections = payload && payload.sections ? payload.sections : {};
    const executiveSummary = toDisplayText(sections.executiveSummary) || "No executive summary available.";
    const mainPoints = toDisplayText(sections.mainPoints) || "No main points available.";
    const actionItems = toDisplayText(sections.actionItems) || "No action items available.";
    byId("executiveSummary").textContent = executiveSummary;
    byId("mainPoints").textContent = mainPoints;
    byId("actionItems").textContent = actionItems;
    window.MailyOffice.applyDirectionToElement(byId("executiveSummary"), executiveSummary);
    window.MailyOffice.applyDirectionToElement(byId("mainPoints"), mainPoints);
    window.MailyOffice.applyDirectionToElement(byId("actionItems"), actionItems);
    const choices = byId("replyActionChoices");
    choices.replaceChildren();
    window.MailySummaryReply.actions(sections.actionItems).forEach(text => {
      const label = document.createElement("label");
      const checkbox = document.createElement("input");
      checkbox.type = "checkbox";
      checkbox.value = text;
      checkbox.addEventListener("change", () => {
        replyReady = false;
        byId("summaryReplyPanel").hidden = true;
        byId("applySummaryReplyBtn").disabled = true;
          setBusy(false);
          setStatus("Selection changed. Prepare a new reply to include it.", "");
      });
      label.append(checkbox, document.createTextNode(` ${text}`));
      choices.append(label);
    });
  }

  async function loadDefaults() {
    try {
      const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
      const payload = await window.MailyOffice.callJson(`/api/outlook/settings?profileId=default&mailboxAddress=${mailboxAddress}`);
      const values = payload && payload.values ? payload.values : payload || {};
      state.defaults = Object.assign({}, state.defaults, values, {
        profileId: payload && payload.profileId ? payload.profileId : state.defaults.profileId
      });
    } catch (_) {
    }
  }

  async function buildRequest() {
    const surface = getSurface();
    const mailboxAddress = window.MailyOffice.getMailboxAddress();

    if (surface === "read") {
      const snapshot = await window.MailyOffice.getReadSnapshot(state.defaults);
      byId("summarySurfaceChip").textContent = "Read";
      return {
        surface,
        mailboxAddress,
        profileId: state.defaults.profileId || "default",
        subject: snapshot.subject || "",
        fromDisplayName: snapshot.fromDisplayName || "",
        fromAddress: snapshot.fromAddress || "",
        to: snapshot.to || [],
        cc: snapshot.cc || [],
        bodyText: snapshot.bodyText || "",
        userDraftText: ""
      };
    }

    const composeSnapshot = await window.MailyOffice.getComposeSnapshot(state.defaults, "generate-initial");
    const composeState = composeSnapshot.composeState || {};
    const request = composeSnapshot.request || {};
    const compose = request.compose || {};
    byId("summarySurfaceChip").textContent = "Compose";
    return {
      surface,
      mailboxAddress,
      profileId: state.defaults.profileId || "default",
      subject: compose.subject || "",
      fromDisplayName: compose.fromDisplayName || "",
      fromAddress: compose.fromAddress || "",
      to: compose.to || [],
      cc: compose.cc || [],
      bodyText: typeof composeState.summaryBodyText === "string" ? composeState.summaryBodyText : (compose.bodyText || ""),
      userDraftText: composeState.userDraftText || ""
    };
  }

  async function summarize() {
    summaryBound = false;
    replyReady = false;
    if (replyFlow) replyFlow.clear();
    byId("summaryReplyPanel").hidden = true;
    setBusy(true);
    setStatus("Reading Outlook context...", "");
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const sourceItem = Office.context.mailbox.item;
      const sourceId = sourceItem.itemId || "";
      const payload = await buildRequest();
      if (Office.context.mailbox.item !== sourceItem || (sourceItem.itemId || "") !== sourceId) throw new Error("The Outlook item changed. Refresh the summary.");
      syncLanguageSelection(payload);
      applyLanguageOverride(payload);
      const hasContext = (payload.bodyText || "").trim() || (payload.userDraftText || "").trim();
      if (!hasContext) {
        throw new Error("There is not enough message or thread content to summarize yet.");
      }

      setStatus("Building summary...", "");
      const response = await window.MailyOffice.callJson("/api/outlook/summary", {
        method: "POST",
        body: JSON.stringify(payload)
      });
      if (Office.context.mailbox.item !== sourceItem || (sourceItem.itemId || "") !== sourceId) throw new Error("The Outlook item changed. Refresh the summary.");
      renderSummary(response || {});
      replyFlow = window.MailySummaryReply.create({ office: window.MailyOffice, getItem: () => Office.context.mailbox.item, surface: getSurface(), defaults: state.defaults });
      replyFlow.bind(payload);
      summaryBound = true;
      setStatus("Summary ready.", "ok");
      await window.MailyOffice.logClientEvent("summary-complete", "ok", `Surface: ${payload.surface}; Language: ${state.selectedOutputLanguage}; LanguageMode: ${state.languageChosenByUser ? "manual" : "auto"}`, "summary");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Could not summarize this message.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("summary-failed", "error", message, "summary");
    } finally {
      setBusy(false);
    }
  }

  async function prepareReply() {
    if (!summaryBound || !replyFlow) return;
    replyReady = false;
    byId("summaryReplyPanel").hidden = true;
    setBusy(true);
    setStatus("Preparing reply preview...", "");
    try {
      const selected = Array.from(byId("replyActionChoices").querySelectorAll("input:checked"), input => input.value);
      const text = await replyFlow.preview(selected);
      byId("summaryReplyText").textContent = text;
      window.MailyOffice.applyDirectionToElement(byId("summaryReplyText"), text);
      byId("summaryReplyPanel").hidden = false;
      byId("summaryReplyPanel").scrollIntoView({ block: "nearest" });
      replyReady = true;
      setStatus("Review your reply, then use it in Outlook. Nothing has been changed or sent.", "ok");
    } catch (error) { setStatus(error.message || "Could not prepare the reply.", "error"); }
    finally { setBusy(false); }
  }

  async function applyReply() {
    if (!replyReady) return;
    replyReady = false;
    setBusy(true);
    try {
      await replyFlow.apply();
      setStatus(getSurface() === "read" ? "Reply opened in Outlook. Review and edit it before sending." : "Your Outlook draft has been updated. Review it before sending.", "ok");
    } catch (error) { setStatus(error.message || "Could not apply the reply. Prepare a fresh preview.", "error"); }
    finally { setBusy(false); }
  }

  function copySummary() {
    const text = [
      "Executive summary",
      byId("executiveSummary").textContent,
      "",
      "Main points",
      byId("mainPoints").textContent,
      "",
      "Action items",
      byId("actionItems").textContent
    ].join("\n");
    window.MailyOffice.copyText(text);
    setStatus("Summary copied.", "ok");
  }

  async function start() {
    await window.MailyOffice.officeReady();
    if (window.MailyOffice.pruneLegacyPaneButtons) {
      window.MailyOffice.pruneLegacyPaneButtons();
    }
    await window.MailyOffice.logClientEvent("summary-open", "ok", `Surface: ${getSurface()}`, "summary");
    byId("summarizeBtn").addEventListener("click", summarize);
    byId("copySummaryBtn").addEventListener("click", copySummary);
    byId("prepareSummaryReplyBtn").addEventListener("click", prepareReply);
    byId("applySummaryReplyBtn").addEventListener("click", applyReply);
    if (getSurface() === "read") {
      byId("applySummaryReplyBtn").textContent = "Open reply to sender";
      byId("summaryReplyHelp").textContent = "Opens a reply to the sender. If you already started a reply, use Maily in that compose window to preserve your edits. Nothing is sent.";
    }
    ["English", "Hebrew"].forEach((language) => {
      byId(`summaryLanguage${language}Btn`).addEventListener("click", async () => {
        setOutputLanguage(language, { userSelected: true });
        await summarize();
      });
    });
    const auto = byId("summaryLanguageAutoBtn");
    if (auto) auto.addEventListener("click", async () => {
      state.languageChosenByUser = false;
      renderLanguageButtons();
      await summarize();
    });
    renderLanguageButtons();
    setStatus("Checking Maily setup...", "");
    await window.MailyOffice.ensureBackendAvailable();
    await loadDefaults();
    await summarize();
  }

  start().catch((error) => {
    const message = error instanceof Error ? error.message : "Summary pane failed to start.";
    setStatus(message, "error");
    if (window.MailyOffice && window.MailyOffice.logClientEvent) {
      window.MailyOffice.logClientEvent("summary-start-failed", "error", message, "summary");
    }
  });
})();
