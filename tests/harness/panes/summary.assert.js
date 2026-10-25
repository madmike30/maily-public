setTimeout(async () => {
      const result = {
        scenarioName: window.__harness.scenarioName,
        errors: window.__harness.errors,
        status: document.getElementById("summaryStatus").textContent || "",
        exec: document.getElementById("executiveSummary").textContent || "",
        actionItems: document.getElementById("actionItems").textContent || "",
        selectedLanguage: document.querySelector(".language-pill.is-selected") ? document.querySelector(".language-pill.is-selected").textContent || "" : "",
        lastSummaryRequestLanguage: window.__harness.lastSummaryRequestLanguage || "",
        summaryRequestLanguages: window.__harness.summaryRequestLanguages || []
      };
      const isHebrewScenario = result.scenarioName === "hebrew";
      const checkboxes = Array.from(document.querySelectorAll("#replyActionChoices input"));
      result.actionsInitiallyUnchecked = checkboxes.length === 2 && checkboxes.every(input => !input.checked);
      if (checkboxes.length) checkboxes[0].checked = true;
      document.getElementById("prepareSummaryReplyBtn").click();
      for (let i = 0; i < 40 && document.getElementById("summaryReplyPanel").hidden; i++) await new Promise(resolve => setTimeout(resolve, 25));
      result.replyPreview = document.getElementById("summaryReplyText").textContent;
      result.selectedActionCarried = !!window.__harness.replyRequest && window.__harness.replyRequest.operation.userInstruction.includes(checkboxes[0].value);
      result.unselectedActionOmitted = !!window.__harness.replyRequest && !window.__harness.replyRequest.operation.userInstruction.includes(checkboxes[1].value);
      result.typedDraftPreserved = !!window.__harness.replyRequest && window.__harness.replyRequest.operation.currentDraft.includes("We can move the target");
      checkboxes[0].click();
      result.selectionInvalidatesPreview = document.getElementById("summaryReplyPanel").hidden &&
        document.getElementById("applySummaryReplyBtn").disabled &&
        document.getElementById("prepareSummaryReplyBtn").classList.contains("primary");
      document.getElementById("summaryLanguageAutoBtn").click();
      for (let i = 0; i < 40 && document.getElementById("summarizeBtn").disabled; i++) await new Promise(resolve => setTimeout(resolve, 25));
      result.automaticRestored = !window.__harness.lastSummaryRequestLanguage &&
        document.getElementById("summaryLanguageAutoBtn").getAttribute("aria-pressed") === "true";
      const pass = result.selectionInvalidatesPreview && result.automaticRestored && result.actionsInitiallyUnchecked && result.selectedActionCarried && result.unselectedActionOmitted && result.typedDraftPreserved && result.replyPreview === "Synthetic reply preview." && result.errors.length === 0 &&
        result.status.indexOf("Summary ready.") >= 0 &&
        (isHebrewScenario ? result.exec.indexOf("התוכנית זזה") >= 0 : result.exec.indexOf("moving to next week") >= 0) &&
        (isHebrewScenario ? result.actionItems.indexOf("לאשר את לוח הזמנים") >= 0 : result.actionItems.indexOf("Confirm the updated schedule.") >= 0) &&
        (isHebrewScenario ? result.selectedLanguage === "Hebrew" : result.selectedLanguage === "Automatic") &&
        result.lastSummaryRequestLanguage === (isHebrewScenario ? "Hebrew" : "") &&
        !document.getElementById("summaryContext") &&
        !document.getElementById("summaryMeta");
      document.getElementById("result").textContent = JSON.stringify(result, null, 2);
      document.title = pass ? "PASS" : "FAIL";
    }, 2000);
