(function () {
      const errors = [];
      window.__harness = { errors };
      window.addEventListener("error", (event) => {
        errors.push(`error:${event.message}`);
      });
      window.addEventListener("unhandledrejection", (event) => {
        const reason = event.reason && event.reason.message ? event.reason.message : String(event.reason);
        errors.push(`rejection:${reason}`);
      });

      window.Office = {
        onReady(callback) {
          setTimeout(callback, 0);
        },
        context: {
          mailbox: {
            userProfile: {
              displayName: "Alex Example",
              emailAddress: "alex@example.com"
            }
          }
        }
      };

      window.fetch = async function (url, options) {
        const textUrl = String(url);
        if (textUrl.includes("/api/outlook/models")) {
          return new Response(JSON.stringify({ status: "ready", models: [{ id: "synthetic-quick", displayName: "Quick reply model" }, { id: "synthetic-careful", displayName: "Careful reply model" }], recommendedModel: "synthetic-quick" }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/health")) {
          return new Response(JSON.stringify({ status: "ok" }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/readiness")) {
          return new Response(JSON.stringify({ status: "ready", ready: true, codexReady: true, userMessage: "Maily is ready.", recommendedAction: "", checks: { codex: { ok: true, details: {} } } }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/settings") && (!options || !options.method || options.method === "GET")) {
          return new Response(JSON.stringify({
            profileId: "default",
            updatedUtc: "2026-03-30T12:00:00Z",
            values: {
              userName: "Alex Example",
              tone: "warm",
              length: "concise",
              language: "English",
              model: "",
              customInstructions: "Keep it short.",
              systemPromptOverride: "SYSTEM OVERRIDE",
              draftPromptOverride: "DRAFT OVERRIDE",
              stylePersonalizationEnabled: true
            },
            defaults: {
              userName: "Alex Example",
              tone: "professional",
              length: "concise",
              language: "English",
              model: "",
              customInstructions: "",
              systemPromptOverride: "",
              draftPromptOverride: "",
              stylePersonalizationEnabled: true
            },
            options: {
              model: [],
              tone: ["professional", "warm", "concise", "executive summary"],
              length: ["concise", "standard", "detailed"],
              language: ["English", "Hebrew", "Same as thread"]
            },
            fieldMetadata: {
              userName: { description: "Used as mailbox-owner context for reply generation.", recommended: "Alex Example" },
              tone: { description: "Sets the default voice Maily should prefer when rewriting or drafting.", recommended: "professional" },
              length: { description: "Controls the default depth of generated replies.", recommended: "concise" },
              language: { description: "Used only when the thread itself does not clearly establish the language.", recommended: "English" },
              model: { description: "Optional expert override for the reply-generation model.", recommended: "Blank", isAdvanced: true },
              customInstructions: { description: "Reusable guidance for all replies. Keep this short and stable.", recommended: "Blank", isAdvanced: true },
              systemPromptOverride: { description: "Full expert override for the system prompt. Leave blank for the production baseline.", recommended: "Blank", isAdvanced: true },
              draftPromptOverride: { description: "Expert override for how Maily rewrites an existing user draft.", recommended: "Blank", isAdvanced: true },
              stylePersonalizationEnabled: { description: "When enabled, Maily applies the reusable style profile inferred from Sent Items.", recommended: "Enabled" }
            },
            advancedWarnings: [
              "Advanced AI controls are expert-only.",
              "Prompt overrides can change reply quality."
            ],
            style: {
              enabled: true,
              summary: "Warm, concise, pragmatic.",
              analyzedStyle: "Short paragraphs. Direct asks. Light warmth.",
              sampleCount: 8,
              lastRefreshedUtc: "2026-03-30T10:00:00Z",
              lastRefreshStatus: "ok",
              lastRefreshError: "",
              refreshStateSummary: "8 sent emails analyzed | Last refresh: 2026-03-30T10:00:00Z | Status: ok"
            }
          }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/settings") && options && options.method === "PUT") {
          return new Response(options.body || "{}", { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/style-profile/refresh")) {
          return new Response(JSON.stringify({ ok: true }), { status: 200, headers: { "Content-Type": "application/json" } });
        }
        if (textUrl.includes("/api/outlook/client-log")) {
          return new Response("", { status: 202 });
        }
        return new Response("", { status: 200 });
      };

      window.Response = window.Response || class Response {
        constructor(body, init) {
          this._body = body;
          this.status = init.status;
          this.ok = this.status >= 200 && this.status < 300;
          this.headers = init.headers || {};
        }
        async json() { return JSON.parse(this._body || "{}"); }
        async text() { return this._body; }
      };
    })();
