(function () {
      const url = new URL(window.location.href);
      const scenarioName = String(url.searchParams.get("scenario") || "english").trim().toLowerCase();
      const errors = [];
      window.__harness = { errors };
      window.__harness.scenarioName = scenarioName;
      window.__harness.generateRequestLanguages = [];
      window.__harness.recovery = {};
      window.__harness.metrics = {
        setBodyCalls: 0,
        prependCalls: 0,
        rewriteBridgeCalls: 0,
        lastSetBody: "",
        lastPrependedBody: "",
        lastRewriteBody: ""
      };
      window.addEventListener("error", (event) => {
        errors.push(`error:${event.message}`);
      });
      window.addEventListener("unhandledrejection", (event) => {
        const reason = event.reason && event.reason.message ? event.reason.message : String(event.reason);
        errors.push(`rejection:${reason}`);
      });

      function asyncSuccess(value, callback) {
        setTimeout(() => callback({
          status: "succeeded",
          value
        }), 0);
      }

      function getInitialBodyHtml() {
        if (scenarioName === "auto-hebrew") {
          return [
            "<div><br></div>",
            "<div>Regards,</div>",
            "<div>Alex Example | Project Lead</div>",
            "<div><a href='mailto:alex@example.com'>alex@example.com</a></div>",
            "<div id='divRplyFwdMsg'><div>From: Team</div><div>Subject: תוכנית</div><div>שלום אלכס, אפשר להתקדם עם זה?</div></div>"
          ].join("");
        }

        return [
          "<div>Yes please do it.</div>",
          "<div>Regards,</div>",
          "<div>Alex Example | Project Lead</div>",
          "<div><a href='mailto:alex@example.com'>alex@example.com</a></div>",
          "<div id='divRplyFwdMsg'><div>From: Example Meeting Team</div><div>Subject: Meeting</div></div>"
        ].join("");
      }

      let bodyHtml = getInitialBodyHtml();
      let setBodyCalls = 0;
      let prependCalls = 0;
      let rewriteBridgeCalls = 0;
      let lastSetBody = "";
      let lastPrependedBody = "";
      let lastRewriteBody = "";

      function countOccurrences(value, token) {
        return (String(value || "").match(new RegExp(token.replace(/[.*+?^${}()|[\]\\]/g, "\\$&"), "g")) || []).length;
      }

      function htmlToPlainText(html) {
        return String(html || "")
          .replace(/<br\s*\/?>/gi, "\n")
          .replace(/<\/div>/gi, "\n")
          .replace(/<[^>]+>/g, "")
          .replace(/\n{3,}/g, "\n\n")
          .trim();
      }

      function getScenarioConfig() {
        if (scenarioName === "hebrew") {
          return {
            expectedSelectedLanguage: "Hebrew",
            expectedFinalRequestLanguage: "Hebrew",
            expectedBodyText: "כן, אפשר להתקדם עם זה.",
            expectedGenerateLanguages: ["", "Hebrew"],
            expectedSetBodyCalls: 2,
            expectedPrependCalls: 0,
            expectedRewriteBridgeCalls: 0,
            expectedSeedText: null,
            requireRtl: true,
            requireQuotedSuffix: true
          };
        }

        if (scenarioName === "auto-hebrew") {
          return {
            // The mixed signature/thread snapshot is ambiguous to the UI; the
            // backend still receives Auto and returns the Hebrew reply below.
            expectedSelectedLanguage: "",
            expectedFinalRequestLanguage: "",
            expectedBodyText: "כן, אפשר להתקדם עם זה.",
            expectedGenerateLanguages: [""],
            expectedSetBodyCalls: 0,
            expectedPrependCalls: 1,
            expectedRewriteBridgeCalls: 0,
            expectedSeedText: null,
            requireRtl: true,
            requireQuotedSuffix: true
          };
        }

        return {
          // Latin script is intentionally Auto: it does not establish English.
          expectedSelectedLanguage: "",
          expectedFinalRequestLanguage: "",
          expectedBodyText: "Yes, please go ahead with it.",
          expectedGenerateLanguages: [""],
          expectedSetBodyCalls: 1,
          expectedPrependCalls: 0,
          expectedRewriteBridgeCalls: 0,
          expectedSeedText: "Yes please do it.",
          requireRtl: false,
          requireQuotedSuffix: true
        };
      }

      function arraysEqual(left, right) {
        if (!Array.isArray(left) || !Array.isArray(right) || left.length !== right.length) {
          return false;
        }

        for (let index = 0; index < left.length; index += 1) {
          if (left[index] !== right[index]) {
            return false;
          }
        }

        return true;
      }

      function buildHarnessResult() {
        const activeBody = lastSetBody || lastRewriteBody || lastPrependedBody || "";
        return {
          scenarioName,
          previewBeforeApply: window.__harness.previewBeforeApply,
          applyClicks: window.__harness.applyClicks || 0,
          errors: window.__harness.errors,
          buttons: {
            generate: !!document.getElementById("generateBtn"),
            regenerate: !!document.getElementById("regenerateBtn"),
            replace: !!document.getElementById("replaceBtn"),
            insert: !!document.getElementById("insertBtn"),
            copy: !!document.getElementById("copyBtn")
          },
          status: document.getElementById("status") ? document.getElementById("status").textContent || "" : "",
          seed: document.getElementById("seedText") ? document.getElementById("seedText").textContent || "" : "",
          selectedLanguage: document.querySelector(".language-pill.is-selected") ? document.querySelector(".language-pill.is-selected").textContent || "" : "",
          setBodyCalls,
          prependCalls,
          rewriteBridgeCalls,
          lastSetBody,
          lastPrependedBody,
          lastRewriteBody,
          finalBody: activeBody,
          generateRequestLanguages: Array.isArray(window.__harness.generateRequestLanguages) ? window.__harness.generateRequestLanguages.slice() : [],
          lastGenerateRequestLanguage: window.__harness.lastGenerateRequestLanguage || "",
          signatureMailtoCount: countOccurrences(activeBody, "mailto:alex@example.com"),
          signatureTitleCount: countOccurrences(activeBody, "Project Lead"),
          recovery: window.__harness.recovery,
          preview: document.getElementById('draftText')?.textContent || ''
        };
      }

      function evaluateHarnessPass(result) {
        if (scenarioName === 'failed-regenerate') {
          const recovery = result.recovery;
          return result.previewBeforeApply === true && result.errors.length === 0 &&
            recovery.previousPreview === 'Yes, please go ahead with it.' &&
            recovery.failureStatus.length > 0 &&
            recovery.restoredPreview === recovery.previousPreview &&
            recovery.bodyUnchangedAfterFailure === true &&
            recovery.insertClicked === true &&
            result.prependCalls === 0 && result.rewriteBridgeCalls === 0 && result.setBodyCalls === 2 &&
            result.generateRequestLanguages.length === 3 &&
            result.finalBody.includes(recovery.previousPreview) &&
            countOccurrences(result.finalBody, '<!-- MAILY_REPLY_START -->') === 1 &&
            result.signatureMailtoCount === 1 && result.signatureTitleCount === 1 &&
            !result.finalBody.includes('UNFINISHED UNSAFE TEXT') &&
            result.preview === recovery.previousPreview;
        }
        const scenario = getScenarioConfig();
        const hasExpectedText = result.finalBody.includes(scenario.expectedBodyText);
        const generateLanguagesMatch = arraysEqual(result.generateRequestLanguages, scenario.expectedGenerateLanguages);
        const generatedBody = new DOMParser().parseFromString(result.finalBody, "text/html");
        const hasRtlMarkup = Array.from(generatedBody.querySelectorAll('[dir="rtl"]'))
          .some(element => element.style.textAlign === "right");
        const seedMatches = typeof scenario.expectedSeedText === "string"
          ? result.seed.trim() === scenario.expectedSeedText
          : true;

        return result.previewBeforeApply === true && result.errors.length === 0 &&
          result.buttons.generate &&
          result.buttons.insert &&
          !result.buttons.regenerate &&
          !result.buttons.replace &&
          !result.buttons.copy &&
          result.selectedLanguage === scenario.expectedSelectedLanguage &&
          seedMatches &&
          result.rewriteBridgeCalls === scenario.expectedRewriteBridgeCalls &&
          result.setBodyCalls === scenario.expectedSetBodyCalls &&
          result.prependCalls === scenario.expectedPrependCalls &&
          result.lastGenerateRequestLanguage === scenario.expectedFinalRequestLanguage &&
          generateLanguagesMatch &&
          hasExpectedText &&
          (!scenario.requireQuotedSuffix || result.finalBody.includes("id='divRplyFwdMsg'")) &&
          result.signatureMailtoCount === 1 &&
          result.signatureTitleCount === 1 &&
          (!scenario.requireRtl || hasRtlMarkup);
      }

      function updateHarnessResult() {
        const result = buildHarnessResult();
        window.__harness.metrics = {
          setBodyCalls,
          prependCalls,
          rewriteBridgeCalls,
          lastSetBody,
          lastPrependedBody,
          lastRewriteBody
        };
        const pass = evaluateHarnessPass(result);

        const resultNode = document.getElementById("result");
        if (resultNode) {
          resultNode.textContent = JSON.stringify(result, null, 2);
        }

        document.title = pass ? "PASS" : "FAIL";
      }

      function runScenarioFlow() {
        let checks = 0;
        const applyPreview = () => {
          const apply = document.getElementById('applyBtn');
          if (apply && !apply.disabled) {
            if (window.__harness.previewBeforeApply === undefined) {
              window.__harness.previewBeforeApply = setBodyCalls === 0 && prependCalls === 0 && rewriteBridgeCalls === 0 && bodyHtml === getInitialBodyHtml();
            }
            window.__harness.applyClicks = (window.__harness.applyClicks || 0) + 1;
            apply.click();
          }
          if (++checks < 38) setTimeout(applyPreview, 100);
        };
        setTimeout(applyPreview, 100);
        if (scenarioName === 'failed-regenerate') {
          const afterSuccess = () => {
            const generate = document.getElementById('generateBtn');
            if (!generate || generate.disabled || setBodyCalls !== 1) {
              setTimeout(afterSuccess, 100);
              return;
            }
            const recovery = window.__harness.recovery;
            recovery.previousPreview = document.getElementById('draftText').textContent;
            const previousBody = bodyHtml;
            generate.click();
            const afterFailure = () => {
              if (generate.disabled || window.__harness.generateRequestLanguages.length !== 2) {
                setTimeout(afterFailure, 100);
                return;
              }
              recovery.failureStatus = document.getElementById('status').textContent;
              recovery.restoredPreview = document.getElementById('draftText').textContent;
              recovery.bodyUnchangedAfterFailure = bodyHtml === previousBody;
              const insert = document.getElementById('insertBtn');
              recovery.insertClicked = !!insert && !insert.disabled;
              insert?.click();
            };
            setTimeout(afterFailure, 100);
          };
          setTimeout(afterSuccess, 600);
          return;
        }
        if (scenarioName !== "hebrew") {
          return;
        }

        const trySwitchLanguage = () => {
          const generateButton = document.getElementById("generateBtn");
          const hebrewButton = document.getElementById("replyLanguageHebrewBtn");
          if (!generateButton || !hebrewButton || generateButton.disabled) {
            setTimeout(trySwitchLanguage, 200);
            return;
          }

          hebrewButton.click();
          setTimeout(() => {
            if (!generateButton.disabled) {
              generateButton.click();
            }
          }, 150);
        };

        setTimeout(trySwitchLanguage, 1400);
      }

      window.Office = {
        AsyncResultStatus: {
          Succeeded: "succeeded",
          Failed: "failed"
        },
        CoercionType: {
          Html: "html",
          Text: "text"
        },
        onReady(callback) {
          setTimeout(callback, 0);
        },
        context: {
          platform: 'PC',
          diagnostics: {
            hostName: "Outlook",
            hostVersion: "mock"
          },
          mailbox: {
            userProfile: {
              displayName: "Alex Example",
              emailAddress: "alex@example.com"
            },
            item: {
              subject: {
                getAsync(callback) {
                  asyncSuccess("RE: Meeting", callback);
                }
              },
              to: {
                getAsync(callback) {
                  asyncSuccess([{ emailAddress: "meeting@example.invalid" }], callback);
                }
              },
              cc: {
                getAsync(callback) {
                  asyncSuccess([], callback);
                }
              },
              from: {
                getAsync(callback) {
                  asyncSuccess({
                    displayName: "Alex Example",
                    emailAddress: "alex@example.com"
                  }, callback);
                }
              },
              body: {
                getTypeAsync(callback) {
                  asyncSuccess("html", callback);
                },
                getAsync(type, callback) {
                  asyncSuccess(type === "text" ? htmlToPlainText(bodyHtml) : bodyHtml, callback);
                },
                prependAsync(content, options, callback) {
                  prependCalls += 1;
                  bodyHtml = `${String(content || "")}${bodyHtml}`;
                  lastPrependedBody = bodyHtml;
                  updateHarnessResult();
                  asyncSuccess(null, callback);
                },
                setAsync(content, options, callback) {
                  setBodyCalls += 1;
                  lastSetBody = String(content || "");
                  bodyHtml = String(content || "");
                  updateHarnessResult();
                  asyncSuccess(null, callback);
                }
              }
            }
          }
        }
      };

      window.fetch = async function (url, options) {
        const textUrl = String(url);
        if (textUrl.includes("/api/health")) {
          return new Response(JSON.stringify({ status: "ok" }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/readiness")) {
          return new Response(JSON.stringify({ status: "ready", ready: true, codexReady: true, userMessage: "Maily is ready.", recommendedAction: "", checks: { codex: { ok: true, details: {} } } }), { status: 200, headers: { "Content-Type": "application/json" } });
        }

        if (textUrl.includes("/api/outlook/settings")) {
          return new Response(JSON.stringify({
            profileId: "default",
            userName: "Alex Example",
            tone: "warm",
            length: "concise",
            language: "English",
            model: "gpt-5.4-nano",
            customInstructions: "",
            systemPromptOverride: "",
            draftPromptOverride: ""
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }

        if (textUrl.includes("/api/outlook/generate-reply")) {
          const payload = options && options.body ? JSON.parse(options.body) : {};
          window.__harness.lastGenerateRequestLanguage = payload && payload.operation ? payload.operation.outputLanguageOverride || "" : "";
          window.__harness.generateRequestLanguages.push(window.__harness.lastGenerateRequestLanguage);
          if (scenarioName === 'failed-regenerate' && window.__harness.generateRequestLanguages.length === 2) {
            return new Response(JSON.stringify({ error: 'Synthetic generation failure' }), { status: 503, headers: { 'Content-Type': 'application/json' } });
          }
          const contextText = [
            payload && payload.operation ? payload.operation.currentDraft || "" : "",
            payload && payload.compose ? payload.compose.bodyText || "" : "",
            payload && payload.compose ? payload.compose.subject || "" : ""
          ].join("\n");
          const hasHebrewContext = /[\u0590-\u05FF]/.test(contextText);
          const isHebrew = window.__harness.lastGenerateRequestLanguage === "Hebrew" ||
            (!window.__harness.lastGenerateRequestLanguage && hasHebrewContext);
          const response = {
            draft: {
              html: isHebrew
                ? "<div style='margin:0; line-height:1.4;'>כן, אפשר להתקדם עם זה.</div>"
                : "<div style='margin:0; line-height:1.4;'>Yes, please go ahead with it.</div>",
              text: isHebrew
                ? "כן, אפשר להתקדם עם זה."
                : "Yes, please go ahead with it."
            },
            meta: {
              model: "gpt-5.4-nano",
              requestId: "mock-request"
            }
          };
          if (textUrl.endsWith('/stream')) {
            return new Response([
              JSON.stringify({ type: 'delta', text: response.draft.text }),
              JSON.stringify({ type: 'complete', result: response }),
              ''
            ].join('\n'), { status: 200, headers: { 'Content-Type': 'application/x-ndjson' } });
          }
          return new Response(JSON.stringify(response), { status: 200, headers: { "Content-Type": "application/json" } });
        }

        if (textUrl.includes("/api/outlook/rewrite-active-compose")) {
          const payload = options && options.body ? JSON.parse(options.body) : {};
          rewriteBridgeCalls += 1;
          bodyHtml = `<!-- MAILY_REPLY_START -->${payload.html || ""}<!-- MAILY_REPLY_END --><div>Regards,</div><div>Alex Example | Project Lead</div><div><a href='mailto:alex@example.com'>alex@example.com</a></div><div id='divRplyFwdMsg'><div>From: Example Meeting Team</div><div>Subject: Meeting</div></div>`;
          lastRewriteBody = bodyHtml;
          updateHarnessResult();
          return new Response(JSON.stringify({
            ok: true,
            replaceStart: 0,
            replaceEnd: 18,
            detectedDraftText: "Yes please do it."
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }

        if (textUrl.includes("/api/outlook/client-log")) {
          return new Response("", { status: 202 });
        }

        return new Response("not found", { status: 404 });
      };

      window.Response = window.Response || class Response {
        constructor(body, init) {
          this._body = body;
          this.status = init.status;
          this.ok = this.status >= 200 && this.status < 300;
          this.headers = init.headers || {};
        }
        async json() { return JSON.parse(this._body); }
        async text() { return this._body; }
      };

      window.__harness.snapshot = buildHarnessResult;
      window.__harness.evaluate = evaluateHarnessPass;
      runScenarioFlow();
    })();
