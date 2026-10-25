(function () {
  const state = {
    selectedOutputLanguage: "English",
    languageChosenByUser: false,
    defaults: {
      profileId: "default",
      language: "English"
    }
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function setStatus(text, kind) {
    const el = byId("briefingStatus");
    el.textContent = text;
    el.className = "status";
    if (kind) {
      el.classList.add(kind);
    }
  }

  function setBusy(isBusy) {
    const auto = byId("briefingLanguageAutoBtn");
    if (auto) auto.disabled = isBusy;
    ["refreshBriefingBtn", "copyBriefingBtn"].forEach((id) => {
      byId(id).disabled = isBusy;
    });
  }

  function normalizeLanguage(value) {
    return String(value || "").trim().toLowerCase() === "hebrew" ? "Hebrew" : "English";
  }

  function renderLanguageButtons() {
    const summary = byId("briefingLanguageSummary");
    if (summary) summary.textContent = "Language: " + (state.languageChosenByUser ? state.selectedOutputLanguage : "Automatic");
    const auto = byId("briefingLanguageAutoBtn");
    if (auto) {
      auto.classList.toggle("is-selected", !state.languageChosenByUser);
      auto.setAttribute("aria-pressed", state.languageChosenByUser ? "false" : "true");
    }
    ["English", "Hebrew"].forEach((language) => {
      const button = byId(`briefingLanguage${language}Btn`);
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

  function syncLanguageSelection() {
    if (state.languageChosenByUser) {
      return;
    }

    setOutputLanguage(state.defaults.language || "English");
  }

  function getOutputLanguageOverride() {
    return state.languageChosenByUser ? state.selectedOutputLanguage : undefined;
  }

  function renderSection(id, text) {
    const value = text || "No briefing content available.";
    const el = byId(id);
    el.textContent = value;
    window.MailyOffice.applyDirectionToElement(el, value);
  }

  function renderBriefing(payload) {
    const sections = payload && payload.sections ? payload.sections : {};
    const meta = payload && payload.meta ? payload.meta : {};

    renderSection("briefingRightNow", sections.rightNow);
    renderSection("briefingMeetingsToday", sections.meetingsToday);
    renderSection("briefingOpenActionItems", sections.openActionItems);
    renderSection("briefingImportantUnread", sections.importantUnread);
    renderSection("briefingNextFocus", sections.nextFocus);

    byId("briefingMeta").textContent =
      `Meetings: ${meta.meetingCount || 0} | Open actions: ${meta.openActionItemCount || 0} | Unread: ${meta.importantUnreadCount || 0} | Model: ${meta.model || "-"} | Request: ${meta.requestId || "-"}`;
  }

  async function loadDefaults() {
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
      const payload = await window.MailyOffice.callJson(`/api/outlook/settings?profileId=default&mailboxAddress=${mailboxAddress}`);
      const values = payload && payload.values ? payload.values : payload || {};
      state.defaults = Object.assign({}, state.defaults, values, {
        profileId: payload && payload.profileId ? payload.profileId : "default"
      });
    } catch (_) {
      state.defaults = {
        profileId: "default",
        language: "English"
      };
    }
  }

  async function refreshBriefing() {
    setBusy(true);
    setStatus("Collecting meetings and mailbox priorities...", "");
    try {
      await window.MailyOffice.ensureBackendAvailable();
      const response = await window.MailyOffice.callJson("/api/outlook/daily-briefing", {
        method: "POST",
        body: JSON.stringify({
          mailboxAddress: window.MailyOffice.getMailboxAddress(),
          profileId: state.defaults.profileId || "default",
          outputLanguageOverride: getOutputLanguageOverride()
        })
      });
      renderBriefing(response || {});
      setStatus("Daily briefing ready.", "ok");
      await window.MailyOffice.logClientEvent("daily-briefing-complete", "ok", `Language: ${state.selectedOutputLanguage}; LanguageMode: ${state.languageChosenByUser ? "manual" : "auto"}`, "daily-briefing");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Could not build the daily briefing.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("daily-briefing-failed", "error", message, "daily-briefing");
    } finally {
      setBusy(false);
    }
  }

  function copyBriefing() {
    const text = [
      "Right now",
      byId("briefingRightNow").textContent,
      "",
      "Meetings today",
      byId("briefingMeetingsToday").textContent,
      "",
      "Open action items",
      byId("briefingOpenActionItems").textContent,
      "",
      "Important unread",
      byId("briefingImportantUnread").textContent,
      "",
      "Next focus",
      byId("briefingNextFocus").textContent
    ].join("\n");
    window.MailyOffice.copyText(text);
    setStatus("Briefing copied.", "ok");
  }

  async function start() {
    await window.MailyOffice.officeReady();
    if (window.MailyOffice.pruneLegacyPaneButtons) {
      window.MailyOffice.pruneLegacyPaneButtons();
    }
    await window.MailyOffice.logClientEvent("daily-briefing-open", "ok", "Daily briefing pane opened.", "daily-briefing");
    byId("refreshBriefingBtn").addEventListener("click", refreshBriefing);
    byId("copyBriefingBtn").addEventListener("click", copyBriefing);
    ["English", "Hebrew"].forEach((language) => {
      byId(`briefingLanguage${language}Btn`).addEventListener("click", async () => {
        setOutputLanguage(language, { userSelected: true });
        await refreshBriefing();
      });
    });
    const auto = byId("briefingLanguageAutoBtn");
    if (auto) auto.addEventListener("click", async () => {
      state.languageChosenByUser = false;
      renderLanguageButtons();
      await refreshBriefing();
    });
    renderLanguageButtons();
    setStatus("Checking Maily setup...", "");
    await window.MailyOffice.ensureBackendAvailable();
    await loadDefaults();
    syncLanguageSelection();
    await refreshBriefing();
  }

  start().catch((error) => {
    const message = error instanceof Error ? error.message : "Daily briefing pane failed to start.";
    setStatus(message, "error");
    if (window.MailyOffice && window.MailyOffice.logClientEvent) {
      window.MailyOffice.logClientEvent("daily-briefing-start-failed", "error", message, "daily-briefing");
    }
  });
})();
