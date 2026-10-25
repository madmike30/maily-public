(function () {
  const state = {
    defaults: {
      profileId: "default"
    },
    sessionId: "",
    activeQuery: "",
    availableRefinements: null
  };

  function byId(id) {
    return document.getElementById(id);
  }

  function setStatus(text, kind) {
    const el = byId("searchStatus");
    el.textContent = text;
    el.className = "status";
    if (kind) {
      el.classList.add(kind);
    }
  }

  function setBusy(isBusy) {
    byId("searchBtn").disabled = isBusy;
    byId("clearSearchBtn").disabled = isBusy;
    byId("refineSearchBtn").disabled = isBusy || !state.sessionId;
  }

  function escapeHtml(value) {
    return String(value || "")
      .replace(/&/g, "&amp;")
      .replace(/</g, "&lt;")
      .replace(/>/g, "&gt;")
      .replace(/"/g, "&quot;");
  }

  function renderSelect(id, options, selectedValue, emptyLabel) {
    const select = byId(id);
    const values = Array.isArray(options) ? options : [];
    const items = [];
    if (typeof emptyLabel === "string") {
      items.push(`<option value="">${escapeHtml(emptyLabel)}</option>`);
    }
    values.forEach((option) => {
      items.push(`<option value="${escapeHtml(option.value || "")}">${escapeHtml(option.label || option.value || "")}</option>`);
    });
    select.innerHTML = items.join("");
    select.value = selectedValue || "";
  }

  function currentFilters() {
    return {
      scope: byId("searchScope").value || "all-mail",
      dateRange: byId("searchDateRange").value || "90d",
      sender: byId("searchSender").value || "",
      focus: byId("searchFocus").value || "all"
    };
  }

  function renderResults(items) {
    const container = byId("searchResults");
    if (!items || items.length === 0) {
      container.textContent = "No ranked matches were returned for this search state.";
      window.MailyOffice.applyDirectionToElement(container, container.textContent);
      return;
    }

    container.innerHTML = items.map((item) => `
      <article class="result-card">
        <div class="result-meta">
          <span class="chip">${escapeHtml(item.folderName || "Mailbox")}</span>
          <span class="meta">${escapeHtml(item.receivedDateTime || "")}</span>
        </div>
        <h3 ${window.MailyOffice.getDirectionAttributes(item.subject || "")}>${escapeHtml(item.subject || "(No subject)")}</h3>
        <p ${window.MailyOffice.getDirectionAttributes(item.fromLine || "")}>${escapeHtml(item.fromLine || "")}</p>
        <strong ${window.MailyOffice.getDirectionAttributes(item.whyItMatches || "")}>${escapeHtml(item.whyItMatches || "")}</strong>
        <div class="long-copy" ${window.MailyOffice.getDirectionAttributes(item.bodyPreview || "")}>${escapeHtml(item.bodyPreview || "")}</div>
      </article>
    `).join("");
  }

  function renderActiveChips(filters) {
    const chips = [
      `Scope: ${labelForSelect("searchScope", filters.scope)}`,
      `Time: ${labelForSelect("searchDateRange", filters.dateRange)}`,
      `Focus: ${labelForSelect("searchFocus", filters.focus)}`
    ];
    if (filters.sender) {
      chips.push(`Sender: ${labelForSelect("searchSender", filters.sender)}`);
    }

    byId("activeSearchChips").innerHTML = chips.map((chip) => `<span class="chip-soft">${escapeHtml(chip)}</span>`).join("");
  }

  function labelForSelect(id, value) {
    const select = byId(id);
    const option = Array.from(select.options).find((entry) => entry.value === value);
    return option ? option.textContent : value;
  }

  function renderResponse(payload) {
    state.sessionId = payload && payload.meta ? payload.meta.sessionId || "" : "";
    state.activeQuery = payload.query || "";
    state.availableRefinements = payload.availableRefinements || null;
    const filters = payload.filters || currentFilters();
    renderActiveChips(filters);

    const answerSummary = payload.answerSummary || "No answer summary returned.";
    const commitments = payload.commitments || "No commitments detected.";
    byId("answerSummary").textContent = answerSummary;
    byId("commitments").textContent = commitments;
    window.MailyOffice.applyDirectionToElement(byId("answerSummary"), answerSummary);
    window.MailyOffice.applyDirectionToElement(byId("commitments"), commitments);
    renderResults(payload.results || []);

    const meta = payload.meta || {};
    byId("searchMeta").textContent = `Stage: ${meta.queryStage || "-"} | Candidates: ${meta.candidateCount || 0} | Results: ${meta.resultCount || 0} | Model: ${meta.model || "-"} | Request: ${meta.requestId || "-"}`;
    byId("searchEvidenceChip").textContent = meta.evidenceMode || "unknown";
    byId("searchTrustNote").textContent = meta.evidenceMode === "hydrated-top-results"
      ? `Search narrowed to ${meta.candidateCount || 0} candidates and deepened the strongest ${meta.hydratedCount || 0} messages with full-message evidence.`
      : `Search is currently based on ranked mailbox previews across ${meta.scopeLabel || "the selected scope"}. Apply filters to deepen the strongest messages.`;

    if (meta.warning) {
      byId("searchTrustNote").textContent += " " + meta.warning;
    }

    renderSelect("searchSender", state.availableRefinements && state.availableRefinements.sender ? state.availableRefinements.sender : [], filters.sender || "", "Any sender");
    byId("searchSender").disabled = !(state.availableRefinements && state.availableRefinements.sender && state.availableRefinements.sender.length);
    byId("refineSearchBtn").disabled = !state.sessionId;
  }

  async function loadDefaults() {
    const mailboxAddress = encodeURIComponent(window.MailyOffice.getMailboxAddress());
    try {
      const settings = await window.MailyOffice.callJson(`/api/outlook/settings?profileId=default&mailboxAddress=${mailboxAddress}`);
      state.defaults.profileId = settings && settings.profileId ? settings.profileId : "default";
    } catch (_) {
      state.defaults.profileId = "default";
    }

    renderSelect("searchScope", [
      { value: "all-mail", label: "All mail" },
      { value: "inbox", label: "Inbox" },
      { value: "sent-items", label: "Sent Items" }
    ], "all-mail");
    renderSelect("searchDateRange", [
      { value: "30d", label: "Last 30 days" },
      { value: "90d", label: "Last 90 days" },
      { value: "180d", label: "Last 180 days" },
      { value: "365d", label: "Last year" },
      { value: "all", label: "Any time" }
    ], "90d");
    renderSelect("searchFocus", [
      { value: "all", label: "All matches" },
      { value: "commitments", label: "Commitments" },
      { value: "decisions", label: "Decisions" }
    ], "all");
    renderSelect("searchSender", [], "", "Any sender");
    byId("searchSender").disabled = true;
  }

  async function startSearch() {
    const query = byId("searchQuery").value.trim();
    if (!query) {
      setStatus("Enter a search query first.", "error");
      return;
    }

    setBusy(true);
    setStatus("Scanning mailbox metadata...", "");
    try {
      const response = await window.MailyOffice.callJson("/api/outlook/search/start", {
        method: "POST",
        body: JSON.stringify({
          mailboxAddress: window.MailyOffice.getMailboxAddress(),
          profileId: state.defaults.profileId || "default",
          query,
          filters: currentFilters()
        })
      });
      renderResponse(response || {});
      setStatus("Initial search ready. Apply filters to deepen the strongest evidence.", "ok");
      await window.MailyOffice.logClientEvent("search-start", "ok", `Session: ${(response.meta && response.meta.sessionId) || "-"}`, "search");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Search failed.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("search-start", "error", message, "search");
    } finally {
      setBusy(false);
    }
  }

  async function refineSearch() {
    if (!state.sessionId) {
      setStatus("Start a search first.", "error");
      return;
    }

    setBusy(true);
    setStatus("Refining and deepening the strongest candidates...", "");
    try {
      const response = await window.MailyOffice.callJson("/api/outlook/search/refine", {
        method: "POST",
        body: JSON.stringify({
          sessionId: state.sessionId,
          filters: currentFilters()
        })
      });
      renderResponse(response || {});
      setStatus("Refined search complete.", "ok");
      await window.MailyOffice.logClientEvent("search-refine", "ok", `Session: ${state.sessionId}`, "search");
    } catch (error) {
      const message = error instanceof Error ? error.message : "Refine failed.";
      setStatus(message, "error");
      await window.MailyOffice.logClientEvent("search-refine", "error", message, "search");
    } finally {
      setBusy(false);
    }
  }

  function clearSearch() {
    state.sessionId = "";
    state.activeQuery = "";
    state.availableRefinements = null;
    byId("searchQuery").value = "";
    byId("answerSummary").textContent = "No search run yet.";
    byId("commitments").textContent = "No search run yet.";
    byId("searchResults").textContent = "Run a search to see relevant emails.";
    byId("searchMeta").textContent = "No ranked results yet.";
    byId("searchEvidenceChip").textContent = "Waiting";
    byId("searchTrustNote").textContent = "Maily will show whether the answer came from ranked previews or hydrated full-message evidence.";
    byId("activeSearchChips").innerHTML = '<span class="chip-soft">No active search yet.</span>';
    byId("refineSearchBtn").disabled = true;
    renderSelect("searchSender", [], "", "Any sender");
    byId("searchSender").disabled = true;
    window.MailyOffice.applyDirectionToElement(byId("answerSummary"), byId("answerSummary").textContent);
    window.MailyOffice.applyDirectionToElement(byId("commitments"), byId("commitments").textContent);
    window.MailyOffice.applyDirectionToElement(byId("searchResults"), byId("searchResults").textContent);
    setStatus("Search cleared.", "ok");
  }

  async function start() {
    await window.MailyOffice.officeReady();
    if (window.MailyOffice.pruneLegacyPaneButtons) {
      window.MailyOffice.pruneLegacyPaneButtons();
    }
    byId("searchBtn").addEventListener("click", startSearch);
    byId("refineSearchBtn").addEventListener("click", refineSearch);
    byId("clearSearchBtn").addEventListener("click", clearSearch);
    await window.MailyOffice.logClientEvent("search-open", "ok", "Search pane opened.", "search");
    setStatus("Checking Maily setup...", "");
    await window.MailyOffice.ensureBackendAvailable();
    setStatus("Loading mailbox context...", "");
    await loadDefaults();
    setStatus("Search is ready.", "ok");
  }

  start().catch((error) => {
    const message = error instanceof Error ? error.message : "Search pane failed to start.";
    setStatus(message, "error");
    if (window.MailyOffice && window.MailyOffice.logClientEvent) {
      window.MailyOffice.logClientEvent("search-start-failed", "error", message, "search");
    }
  });
})();
