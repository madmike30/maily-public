function finalizeHarness() {
      const result = window.__harness.snapshot();
      const pass = window.__harness.evaluate(result);

      document.getElementById("result").textContent = JSON.stringify(result, null, 2);
      document.title = pass ? "PASS" : "FAIL";
    }

    setTimeout(finalizeHarness, 4200);
