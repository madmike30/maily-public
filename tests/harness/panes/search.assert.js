setTimeout(() => {
      document.getElementById("searchQuery").value = "What did Alex commit to?";
      document.getElementById("searchBtn").click();
    }, 400);
    setTimeout(() => {
      document.getElementById("searchSender").value = "sam@example.com";
      document.getElementById("searchFocus").value = "commitments";
      document.getElementById("refineSearchBtn").click();
    }, 1200);
    setTimeout(() => {
      const result = {
        errors: window.__harness.errors,
        status: document.getElementById("searchStatus").textContent || "",
        answer: document.getElementById("answerSummary").textContent || "",
        commitments: document.getElementById("commitments").textContent || "",
        results: document.getElementById("searchResults").textContent || "",
        evidence: document.getElementById("searchEvidenceChip").textContent || "",
        trust: document.getElementById("searchTrustNote").textContent || "",
        chips: document.getElementById("activeSearchChips").textContent || ""
      };
      const pass = result.errors.length === 0 &&
        result.status.indexOf("Refined search complete.") >= 0 &&
        result.answer.indexOf("update the schedule next week") >= 0 &&
        result.commitments.indexOf("Update the schedule") >= 0 &&
        result.results.indexOf("Project schedule") >= 0 &&
        result.evidence.indexOf("hydrated-top-results") >= 0 &&
        result.trust.indexOf("limited subset") >= 0 &&
        result.chips.indexOf("Sender") >= 0;
      document.getElementById("result").textContent = JSON.stringify(result, null, 2);
      document.title = pass ? "PASS" : "FAIL";
    }, 2600);
