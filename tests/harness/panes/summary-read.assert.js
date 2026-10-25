history.replaceState(null, "", location.pathname + "?surface=read");
    setTimeout(() => {
      const result = {
        errors: window.__harness.errors,
        status: document.getElementById("summaryStatus").textContent || "",
        chip: document.getElementById("summarySurfaceChip").textContent || "",
        exec: document.getElementById("executiveSummary").textContent || "",
        mainPoints: document.getElementById("mainPoints").textContent || "",
        actionItems: document.getElementById("actionItems").textContent || "",
        selectedLanguage: document.querySelector(".language-pill.is-selected") ? document.querySelector(".language-pill.is-selected").textContent || "" : "",
        lastSummaryRequestLanguage: window.__harness.lastSummaryRequestLanguage || ""
      };
      const pass = result.errors.length === 0 &&
        result.status.indexOf("Summary ready.") >= 0 &&
        result.chip.indexOf("Read") >= 0 &&
        result.exec.indexOf("synthetic training workshop") >= 0 &&
        result.actionItems.indexOf("Review the training resources.") >= 0 &&
        result.selectedLanguage === "Automatic" &&
        result.lastSummaryRequestLanguage === "";
      document.getElementById("result").textContent = JSON.stringify(result, null, 2);
      document.title = pass ? "PASS" : "FAIL";
    }, 2000);
