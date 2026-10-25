setTimeout(() => {
      document.getElementById("openStyleGuidanceBtn")?.click();
      const result = {
        errors: window.__harness.errors,
        userName: document.getElementById("userName")?.value || "",
        systemPromptOverride: document.getElementById("systemPromptOverride")?.value || "",
        styleSummary: document.getElementById("styleProfileSummary")?.textContent || "",
        styleMeta: document.getElementById("styleProfileMeta")?.textContent || "",
        styleRefreshMeta: document.getElementById("styleRefreshMeta")?.textContent || "",
        styleSteeringPreview: document.getElementById("styleSteeringPreview")?.textContent || "",
        advancedOpen: !!document.getElementById("advancedSection")?.open,
        focusedElementId: document.activeElement && document.activeElement.id ? document.activeElement.id : "",
        saveDisabled: !!document.getElementById("saveSettingsBtn")?.disabled,
        saveNote: document.getElementById("settingsSaveNote")?.textContent || ""
      };
      const model = document.getElementById("model");
      const automaticSelected = model.value === "";
      model.value = "synthetic-quick";
      document.getElementById('tourStart').click();
      const tourStarts = !document.getElementById('tourPanel').hidden && document.getElementById('tourProgress').textContent === '1 of 6';
      for (let step = 0; step < 5; step++) document.getElementById('tourNext').click();
      const tourExplainsConsent = document.getElementById('personalStyleSection').open && document.getElementById('tourText').textContent.includes('only when you choose');
      document.getElementById('tourDismiss').click();
      result.tourPassed = tourStarts && tourExplainsConsent && document.getElementById('tourPanel').hidden;
      const pass = result.tourPassed && model.tagName === "SELECT" && automaticSelected && model.value === "synthetic-quick" && model.options.length === 3 && document.getElementById("modelStatus").textContent.includes("Quick reply model") && result.errors.length === 0 &&
        result.userName === "Alex Example" &&
        result.systemPromptOverride === "SYSTEM OVERRIDE" &&
        result.styleSummary.indexOf("Warm, concise, pragmatic.") >= 0 &&
        result.styleMeta.indexOf("Learned from 8 sent emails") >= 0 &&
        result.styleRefreshMeta.indexOf("reusable rules") >= 0 &&
        result.styleSteeringPreview.indexOf("Keep it short.") >= 0 &&
        result.advancedOpen &&
        result.focusedElementId === "customInstructions" &&
        result.saveDisabled &&
        result.saveNote === "Settings saved" && document.getElementById("default-model").textContent === "Default: Automatic";
      document.getElementById("result").textContent = JSON.stringify(result, null, 2);
      document.title = pass ? "PASS" : "FAIL";
    }, 1800);
