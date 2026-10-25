(function () {
  const HTML_START = "<!-- MAILY_REPLY_START -->";
  const HTML_END = "<!-- MAILY_REPLY_END -->";
  const TEXT_START = "[Maily Reply Start]";
  const TEXT_END = "[Maily Reply End]";
  const SIGNATURE_CLOSINGS = [
    "regards",
    "best regards",
    "kind regards",
    "thanks",
    "thank you",
    "cheers",
    "sincerely",
    "warm regards"
  ];
  const HEBREW_REGEX = /[\u0590-\u05FF]/;
  const REPLY_PREFIX_REGEX = /\b(?:re|fw|fwd)\s*:/gi;

  function getContext() {
    return typeof Office !== "undefined" ? Office.context : null;
  }

  function getMailbox() {
    const context = getContext();
    return context && context.mailbox ? context.mailbox : null;
  }

  let composeEpoch = 0;
  let observedMailbox = null;
  let undoRecord = null;
  let composeMutationBusy = false;
  function captureComposeContext() {
    const mailbox = getMailbox();
    if (mailbox && mailbox !== observedMailbox && mailbox.addHandlerAsync && Office.EventType && Office.EventType.ItemChanged) {
      observedMailbox = mailbox;
      mailbox.addHandlerAsync(Office.EventType.ItemChanged, () => { composeEpoch++; undoRecord = null; });
    }
    const item = ensureComposeItem();
    return { item, itemId: item.itemId || "", epoch: composeEpoch };
  }
  function assertComposeContext(expected) {
    const item = ensureComposeItem();
    if (!expected || expected.item !== item || expected.epoch !== composeEpoch ||
        (expected.itemId && expected.itemId !== (item.itemId || ""))) {
      throw new Error("The Outlook item changed while Maily was working. Generate again from the current draft.");
    }
  }

  function getDiagnostics() {
    const context = getContext();
    return context && context.diagnostics ? context.diagnostics : {};
  }

  function getOfficePlatformName() {
    const context = getContext();
    const value = context && context.platform ? String(context.platform) : "";
    return value.trim();
  }

  function getHostInfo() {
    if (typeof window === "undefined" || !window.location || !window.location.search) {
      return {
        raw: "",
        host: "",
        platform: "",
        version: "",
        locale: ""
      };
    }

    const params = new URLSearchParams(window.location.search || "");
    const raw = (params.get("_host_Info") || "").trim();
    const parts = raw.split("$");
    return {
      raw,
      host: (parts[0] || "").trim(),
      platform: (parts[1] || "").trim(),
      version: (parts[2] || "").trim(),
      locale: (parts[3] || "").trim()
    };
  }

  function getRuntimeHostDetails() {
    const diagnostics = getDiagnostics();
    const hostInfo = getHostInfo();
    const diagnosticsHostName = (diagnostics.hostName || "").trim();
    const officePlatform = getOfficePlatformName();
    const userAgent = typeof navigator !== "undefined" && navigator.userAgent ? navigator.userAgent : "";
    const normalizedHostName = diagnosticsHostName.toLowerCase();
    const normalizedHostInfoHost = hostInfo.host.toLowerCase();
    const normalizedHostInfoPlatform = hostInfo.platform.toLowerCase();
    const normalizedOfficePlatform = officePlatform.toLowerCase();
    const isOutlookHost =
      normalizedHostName.includes("outlook") ||
      normalizedHostInfoHost === "outlook";
    const isDesktopPlatform =
      normalizedHostInfoPlatform === "win32" ||
      normalizedHostInfoPlatform === "mac" ||
      normalizedOfficePlatform === "pc" ||
      normalizedOfficePlatform === "mac" ||
      (typeof window !== "undefined" && typeof window.external === "object" && window.external !== null) ||
      (/windows/i.test(userAgent) && !/officeon(line|web)|outlook-ios|outlook-android/i.test(userAgent));

    return {
      diagnosticsHostName,
      officePlatform,
      hostInfoRaw: hostInfo.raw,
      hostInfoHost: hostInfo.host,
      hostInfoPlatform: hostInfo.platform,
      isOutlookHost,
      isDesktopPlatform,
      resolvedPlatform: isOutlookHost && isDesktopPlatform ? "outlook-classic" : "outlook-web"
    };
  }

  function getErrorMessage(result, fallbackMessage) {
    if (result && result.error && result.error.message) {
      return result.error.message;
    }

    return fallbackMessage;
  }

  function copyObject(source) {
    const target = {};
    Object.keys(source || {}).forEach((key) => {
      target[key] = source[key];
    });
    return target;
  }

  function mergeObjects(base, overrides) {
    const target = copyObject(base || {});
    Object.keys(overrides || {}).forEach((key) => {
      target[key] = overrides[key];
    });
    return target;
  }

  function getElementText(node) {
    return node && typeof node.textContent === "string" ? node.textContent : "";
  }

  function matchesSelector(node, selector) {
    return !!(node && typeof node.matches === "function" && node.matches(selector));
  }

  function querySelectorSafe(node, selector) {
    return node && typeof node.querySelector === "function" ? node.querySelector(selector) : null;
  }

  function querySelectorAllSafe(node, selector) {
    return node && typeof node.querySelectorAll === "function" ? node.querySelectorAll(selector) : [];
  }

  function reverseCopy(items) {
    return (items || []).slice().reverse();
  }

  function officeReady() {
    return new Promise((resolve) => {
      Office.onReady(() => resolve());
    });
  }

  function containsHebrew(value) {
    return HEBREW_REGEX.test(value || "");
  }

  function countMatches(value, regex) {
    const matches = String(value || "").match(regex);
    return matches ? matches.length : 0;
  }

  function normalizeLanguageName(value) {
    const normalized = String(value || "").trim().toLowerCase();
    if (normalized === "hebrew") {
      return "Hebrew";
    }

    if (normalized === "english") {
      return "English";
    }

    return "";
  }

  function stripLanguageNoise(value) {
    const text = String(value || "")
      .replace(/<(style|script)\b[^>]*>[\s\S]*?<\/\1\s*>/gi, " ")
      .replace(/<(?:br\b[^>]*|\/p|\/div)>/gi, "\n")
      .replace(/<[^>]*>/g, " ")
      .replace(/&(?:[a-z]+|#\d+|#x[0-9a-f]+);/gi, " ")
      .replace(/\r\n/g, "\n");
    const lines = [];
    for (const line of text.split("\n")) {
      if (/^\s*(?:--\s*$|On .+wrote:|-----Original Message-----)/i.test(line)) break;
      if (/^\s*(?:From|Sent|Subject|To|Cc):/i.test(line)) {
        if (lines.some(value => value.trim())) break;
        continue;
      }
      if (!/^\s*>/.test(line)) lines.push(line);
    }
    return lines.join("\n").replace(REPLY_PREFIX_REGEX, " ")
      .replace(/\b(?:mailto:|https?:\/\/|www\.)\S*|\b[^\s@]+@[^\s@]+\b/gi, " ").trim();
  }

  function detectPreferredLanguage() {
    for (const sample of arguments) {
      const normalized = stripLanguageNoise(sample);
      const letters = countMatches(normalized, /\p{L}/gu);
      if (!letters) continue;
      const hebrew = countMatches(normalized, /[\u0590-\u05FF]/g);
      // Latin text could be French, Spanish, English, etc. Let generation identify it.
      return hebrew > letters / 2 ? "Hebrew" : "Auto";
    }
    return "Auto";
  }
  function applyDirectionToElement(element, text) {
    if (!element) {
      return;
    }

    const isHebrew = containsHebrew(text);
    element.dir = isHebrew ? "rtl" : "ltr";
    element.classList.toggle("rtl-copy", isHebrew);
    element.classList.toggle("ltr-copy", !isHebrew);
  }

  function getDirectionalText(text) {
    const value = text || "";
    if (!containsHebrew(value)) {
      return value;
    }

    return value
      .replace(/\r\n/g, "\n")
      .split("\n")
      .map((line) => line ? `\u200F${line}` : line)
      .join("\n");
  }

  function getDirectionAttributes(text) {
    return containsHebrew(text)
      ? `dir="rtl" class="rtl-copy"`
      : `dir="ltr" class="ltr-copy"`;
  }

  function applyHtmlDirection(html, fallbackText) {
    const source = html || "";
    if (!containsHebrew(fallbackText || source)) {
      return source;
    }

    if (/dir\s*=\s*["']rtl["']/i.test(source)) {
      return source;
    }

    const wrapperMatch = source.match(/^(?<prefix>\s*(?:<!-- MAILY_REPLY_START -->\s*)?)(?<tag><div\b[^>]*>)/i);
    if (wrapperMatch && wrapperMatch.groups) {
      const originalTag = wrapperMatch.groups.tag;
      let updatedTag = originalTag.replace(/<div\b/i, '<div dir="rtl"');
      if (/style\s*=\s*["']/i.test(updatedTag)) {
        updatedTag = updatedTag.replace(/style\s*=\s*(['"])(.*?)\1/i, (match, quote, styles) => `style=${quote}${styles}; text-align:right;${quote}`);
      } else {
        updatedTag = updatedTag.replace(/>$/, ' style="text-align:right;">');
      }

      return `${wrapperMatch.groups.prefix}${updatedTag}${source.slice(wrapperMatch[0].length)}`;
    }

    return `<div dir="rtl" style="text-align:right;">${source}</div>`;
  }

  function isLocalOrigin() {
    if (typeof window === "undefined" || !window.location) {
      return false;
    }

    const host = (window.location.hostname || "").toLowerCase();
    return host === "localhost" || host === "127.0.0.1" || host === "::1";
  }

  function describeFetchError(error) {
    const rawMessage = error instanceof Error ? error.message : String(error || "");
    const message = (rawMessage || "").trim();

    if (!message || message === "Failed to fetch" || message.includes("NetworkError")) {
      if (isLocalOrigin()) {
        return "The local Maily add-in host is unavailable. Close and reopen Outlook so the Maily background helper can restart it. If this keeps happening, run the Maily installer again and choose Install to repair the local setup.";
      }

      return "The hosted Maily service is unavailable. Check your network connection and reopen this pane. If this keeps happening, contact the Maily service owner.";
    }

    return message;
  }

  function ensureComposeItem() {
    const mailbox = getMailbox();
    const item = mailbox ? mailbox.item : null;
    if (!item || !item.body || !item.subject) {
      throw new Error("This add-in works only in Outlook message compose.");
    }

    return item;
  }

  function ensureMessageItem() {
    const mailbox = getMailbox();
    const item = mailbox ? mailbox.item : null;
    if (!item || !item.body) {
      throw new Error("This add-in needs an Outlook email message.");
    }

    return item;
  }

  function getAsync(target, method, arg) {
    return new Promise((resolve, reject) => {
      if (!target || typeof target[method] !== "function") {
        reject(new Error(`Outlook API method unavailable: ${method}`));
        return;
      }

      const callback = (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(new Error(getErrorMessage(result, "Outlook API call failed.")));
        }
      };

      if (typeof arg === "undefined") {
        target[method](callback);
      } else {
        target[method](arg, callback);
      }
    });
  }

  async function getRecipients(field) {
    const item = ensureComposeItem();
    if (!item[field]) {
      return [];
    }

    const values = await getAsync(item[field], "getAsync");
    return (values || []).map((entry) => entry.emailAddress).filter(Boolean);
  }

  function normalizeRecipientEntries(value) {
    if (!value) {
      return [];
    }

    const entries = Array.isArray(value) ? value : [value];
    return entries
      .map((entry) => {
        if (!entry) {
          return "";
        }

        const displayName = entry.displayName || entry.name || "";
        const emailAddress = entry.emailAddress || entry.address || "";
        if (!displayName && !emailAddress) {
          return "";
        }

        return displayName ? `${displayName} <${emailAddress}>` : emailAddress;
      })
      .filter(Boolean);
  }

  async function readValue(target) {
    if (!target) {
      return null;
    }

    if (typeof target.getAsync === "function") {
      try {
        return await getAsync(target, "getAsync");
      } catch (_) {
        return null;
      }
    }

    return target;
  }

  async function readRecipientsField(target) {
    const value = await readValue(target);
    return normalizeRecipientEntries(value);
  }

  function getUserIdentity() {
    const mailbox = getMailbox();
    const profile = mailbox && mailbox.userProfile ? mailbox.userProfile : {};
    return {
      displayName: (profile.displayName || "").trim(),
      emailAddress: (profile.emailAddress || "").trim()
    };
  }

  async function getComposeSenderIdentity() {
    const item = ensureComposeItem();
    const fallback = getUserIdentity();
    const fromField = item.from;

    if (!fromField || typeof fromField.getAsync !== "function") {
      return fallback;
    }

    try {
      const value = await getAsync(fromField, "getAsync");
      return {
        displayName: ((value && value.displayName) || fallback.displayName || "").trim(),
        emailAddress: ((value && value.emailAddress) || fallback.emailAddress || "").trim()
      };
    } catch (_) {
      return fallback;
    }
  }

  async function getBodyType() {
    const item = ensureComposeItem();
    return await getAsync(item.body, "getTypeAsync");
  }

  async function getComposeType() {
    const item = ensureComposeItem();
    if (typeof item.getComposeTypeAsync !== "function") {
      return "unknown";
    }

    try {
      const value = await getAsync(item, "getComposeTypeAsync");
      if (!value) {
        return "unknown";
      }

      if (typeof value === "string") {
        return value;
      }

      if (typeof value.composeType === "string") {
        return value.composeType;
      }
    } catch (_) {
      return "unknown";
    }

    return "unknown";
  }

  async function getBodyByType(type) {
    const item = ensureComposeItem();
    return await getAsync(item.body, "getAsync", type);
  }

  async function prependBody(content, type, capturedItem) {
    const item = capturedItem || ensureComposeItem();
    await new Promise((resolve, reject) => {
      item.body.prependAsync(content, { coercionType: type }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(getErrorMessage(result, "Failed to prepend content.")));
        }
      });
    });
  }

  async function setBody(content, type, capturedItem) {
    const item = capturedItem || ensureComposeItem();
    await new Promise((resolve, reject) => {
      item.body.setAsync(content, { coercionType: type }, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(new Error(getErrorMessage(result, "Failed to set content.")));
        }
      });
    });
  }

  async function rewriteActiveComposeWithBridge(draft, options) {
    return await callJson("/api/outlook/rewrite-active-compose", {
      method: "POST",
      body: JSON.stringify({
        subject: options && options.subject ? options.subject : "",
        userName: options && options.userName ? options.userName : "",
        mailboxAddress: options && options.mailboxAddress ? options.mailboxAddress : getMailboxAddress(),
        seedText: options && options.seedText ? options.seedText : "",
        composeType: options && options.composeType ? options.composeType : "unknown",
        text: draft && draft.text ? draft.text : "",
        html: draft && draft.html ? draft.html : ""
      })
    });
  }

  function wrapForHtml(html) {
    const directed = applyHtmlDirection(html, html);
    if ((directed || "").includes(HTML_START)) {
      return directed;
    }
    return `${HTML_START}${directed}${HTML_END}`;
  }

  function wrapForText(text) {
    const directed = getDirectionalText(text);
    if ((directed || "").includes(TEXT_START)) {
      return directed;
    }
    return `${TEXT_START}\n${directed}\n${TEXT_END}\n`;
  }

  function replaceBetween(content, startToken, endToken, replacement) {
    const startIndex = content.indexOf(startToken);
    const endIndex = content.indexOf(endToken);
    if (startIndex < 0 || endIndex < 0 || endIndex < startIndex) {
      return { replaced: false, output: content };
    }

    let suffixStart = endIndex + endToken.length;
    // The text wrapper owns one trailing line ending; preserve any user spacing after it.
    if (endToken === TEXT_END) {
      if (content.slice(suffixStart, suffixStart + 2) === "\r\n") suffixStart += 2;
      else if (content[suffixStart] === "\n") suffixStart++;
    }
    const output = `${content.slice(0, startIndex)}${replacement}${content.slice(suffixStart)}`;
    return { replaced: true, output };
  }

  function cleanDraftMarkers(value) {
    let text = value || "";
    for (const marker of [HTML_START, HTML_END, TEXT_START, TEXT_END]) {
      text = text.split(marker).join("");
    }
    return collapseWhitespace(text);
  }

  function collapseWhitespace(value) {
    return (value || "").replace(/\r\n/g, "\n").replace(/\u00a0/g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function collapseInlineWhitespace(value) {
    return (value || "").replace(/\r\n/g, "\n").replace(/\u00a0/g, " ").replace(/[ \t]+/g, " ").replace(/[ \t]+\n/g, "\n").replace(/\n{3,}/g, "\n\n").trim();
  }

  function findFirstMatchIndex(value, patterns) {
    let cutIndex = (value || "").length;

    patterns.forEach((pattern) => {
      const match = pattern.exec(value || "");
      if (match && typeof match.index === "number") {
        cutIndex = Math.min(cutIndex, match.index);
      }
    });

    return cutIndex >= (value || "").length ? -1 : cutIndex;
  }

  function findQuotedThreadIndex(text) {
    const normalized = (text || "").replace(/\r\n/g, "\n");
    const markers = [
      /^\s*From:\s/im,
      /^\s*Sent:\s/im,
      /^\s*Subject:\s/im,
      /^\s*To:\s/im,
      /^\s*Cc:\s/im,
      /^\s*On .+wrote:\s/im,
      /^\s*-{2,}\s*Original Message\s*-{2,}\s*$/im
    ];

    return findFirstMatchIndex(normalized, markers);
  }

  function stripTrailingSignature(text, userName, mailboxAddress) {
    const lines = (text || "").replace(/\r\n/g, "\n").split("\n");
    while (lines.length && !lines[lines.length - 1].trim()) {
      lines.pop();
    }

    const normalizedUserName = (userName || "").trim().toLowerCase();
    const normalizedMailbox = (mailboxAddress || "").trim().toLowerCase();

    for (let i = lines.length - 1; i >= 0; i--) {
      const line = lines[i].trim().toLowerCase();
      if (!line) {
        continue;
      }

      if (isSignatureLine(line, normalizedUserName, normalizedMailbox)) {
        let start = i;
        while (start > 0) {
          const prior = lines[start - 1].trim().toLowerCase();
          if (!prior) {
            start -= 1;
            continue;
          }

          if (isClosingLine(prior) || isSignatureLine(prior, normalizedUserName, normalizedMailbox)) {
            start -= 1;
            continue;
          }

          break;
        }

        return lines.slice(0, start).join("\n").trim();
      }
    }

    return lines.join("\n").trim();
  }

  function stripTrailingSignatureSingleLine(text, userName, mailboxAddress) {
    const normalizedText = (text || "").trim();
    const normalizedUserName = (userName || "").trim();
    const normalizedMailbox = (mailboxAddress || "").trim();

    if (!normalizedText) {
      return "";
    }

    const candidates = [
      normalizedMailbox,
      normalizedUserName
    ].filter(Boolean);

    let cutIndex = -1;
    candidates.forEach((candidate) => {
      const index = normalizedText.toLowerCase().indexOf(candidate.toLowerCase());
      if (index > 0) {
        cutIndex = cutIndex < 0 ? index : Math.min(cutIndex, index);
      }
    });

    if (cutIndex > 0) {
      return normalizedText.slice(0, cutIndex).trim().replace(/[|,-]+$/, "").trim();
    }

    return normalizedText;
  }

  function getMeaningfulTextLines(text) {
    return ((text || "").replace(/\r\n/g, "\n").split("\n"))
      .map((line) => line.trim())
      .filter(Boolean);
  }

  function startsWithSignatureBlock(text, userName, mailboxAddress) {
    const lines = getMeaningfulTextLines(text);
    if (!lines.length) {
      return false;
    }

    const normalizedUserName = (userName || "").trim().toLowerCase();
    const normalizedMailbox = (mailboxAddress || "").trim().toLowerCase();
    const first = lines[0].toLowerCase();
    const next = lines[1] ? lines[1].toLowerCase() : "";
    const nextNext = lines[2] ? lines[2].toLowerCase() : "";
    const firstLooksLikeSignature = isSignatureLine(first, normalizedUserName, normalizedMailbox);
    const hasSignatureEvidenceBelow =
      isSignatureLine(next, normalizedUserName, normalizedMailbox) ||
      isSignatureLine(nextNext, normalizedUserName, normalizedMailbox);

    return firstLooksLikeSignature && hasSignatureEvidenceBelow;
  }

  function findSignatureIndexInText(text, userName, mailboxAddress) {
    const normalized = (text || "").replace(/\r\n/g, "\n");
    const lines = normalized.split("\n");
    const normalizedUserName = (userName || "").trim().toLowerCase();
    const normalizedMailbox = (mailboxAddress || "").trim().toLowerCase();

    while (lines.length && !lines[lines.length - 1].trim()) {
      lines.pop();
    }

    for (let i = lines.length - 1; i >= 0; i--) {
      const line = lines[i].trim().toLowerCase();
      if (!line) {
        continue;
      }

      if (isSignatureLine(line, normalizedUserName, normalizedMailbox)) {
        let start = i;
        while (start > 0) {
          const prior = lines[start - 1].trim().toLowerCase();
          if (!prior) {
            start -= 1;
            continue;
          }

          if (isClosingLine(prior) || isSignatureLine(prior, normalizedUserName, normalizedMailbox)) {
            start -= 1;
            continue;
          }

          break;
        }

        return lines.slice(0, start).join("\n").length;
      }
    }

    return -1;
  }

  function isClosingLine(line) {
    const normalized = (line || "").trim().toLowerCase().replace(/[,:;.!-]+$/g, "").trim();
    return SIGNATURE_CLOSINGS.includes(normalized);
  }

  function isSignatureLine(line, normalizedUserName, normalizedMailbox) {
    const normalized = (line || "").trim().toLowerCase();
    if (!normalized) {
      return false;
    }

    if (normalized === "--" || isClosingLine(normalized)) {
      return true;
    }

    if (normalizedMailbox && normalized.includes(normalizedMailbox)) {
      return true;
    }

    if (normalizedUserName && normalized.includes(normalizedUserName)) {
      return true;
    }

    if (/@/.test(normalized) || /\+?\d[\d\s\-()]{6,}/.test(normalized)) {
      return true;
    }

    if (normalized.includes("|")) {
      return true;
    }

    return false;
  }

  function htmlToText(html) {
    const parser = new DOMParser();
    const doc = parser.parseFromString(html || "", "text/html");
    if (!doc.body) {
      return "";
    }

    const blockTags = new Set([
      "address", "article", "aside", "blockquote", "div", "dl", "fieldset", "figcaption", "figure",
      "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main", "nav",
      "ol", "p", "pre", "section", "table", "td", "th", "tr", "ul"
    ]);

    function walk(node) {
      if (!node) {
        return "";
      }

      if (node.nodeType === Node.TEXT_NODE) {
        return node.textContent || "";
      }

      if (node.nodeType !== Node.ELEMENT_NODE) {
        return "";
      }

      const tag = (node.tagName || "").toLowerCase();
      if (tag === "br") {
        return "\n";
      }

      let text = "";
      node.childNodes.forEach((child) => {
        text += walk(child);
      });

      if (blockTags.has(tag) && text.trim()) {
        return `${text}\n`;
      }

      return text;
    }

    return collapseInlineWhitespace(walk(doc.body));
  }

  function splitTextComposeSegments(body, defaults, mailboxAddress) {
    const normalized = (body || "").replace(/\r\n/g, "\n");
    const quoteIndex = findQuotedThreadIndex(normalized);
    const leading = quoteIndex >= 0 ? normalized.slice(0, quoteIndex) : normalized;
    const quotedSuffix = quoteIndex >= 0 ? normalized.slice(quoteIndex) : "";
    const signatureIndex = findSignatureIndexInText(leading, defaults && defaults.userName ? defaults.userName : "", mailboxAddress || "");
    const userDraftSegment = signatureIndex >= 0 ? leading.slice(0, signatureIndex).trimEnd() : leading.trimEnd();
    const signatureSuffix = signatureIndex >= 0 ? leading.slice(signatureIndex) : "";
    const collapsedUserDraftText = collapseWhitespace(userDraftSegment);
    const shouldTreatAsEmptyDraft =
      !collapsedUserDraftText ||
      (startsWithSignatureBlock(leading, defaults && defaults.userName ? defaults.userName : "", mailboxAddress || "") &&
        !stripTrailingSignature(leading, defaults && defaults.userName ? defaults.userName : "", mailboxAddress || ""));

    if (shouldTreatAsEmptyDraft && normalized.trim()) {
      return {
        visibleTopText: "",
        userDraftText: "",
        userDraftSegment: "",
        preservedSuffix: normalized,
        canRewriteUserDraft: false,
        preservedSuffixExact: true
      };
    }

    return {
      visibleTopText: collapseWhitespace(leading),
      userDraftText: collapsedUserDraftText,
      userDraftSegment,
      preservedSuffix: `${signatureSuffix}${quotedSuffix}`,
      canRewriteUserDraft: true,
      preservedSuffixExact: true
    };
  }

  function isSignatureThreadOnlyText(text, userName, mailboxAddress) {
    const normalized = (text || "").replace(/\r\n/g, "\n");
    if (!normalized.trim()) {
      return false;
    }

    const quoteIndex = findQuotedThreadIndex(normalized);
    if (quoteIndex < 0) {
      return false;
    }

    const leading = normalized.slice(0, quoteIndex);
    if (!leading.trim()) {
      return false;
    }

    if (!startsWithSignatureBlock(leading, userName, mailboxAddress)) {
      return false;
    }

    return !stripTrailingSignature(leading, userName, mailboxAddress);
  }

  function findHtmlQuotedThreadIndex(html) {
    const markers = [
      /<div[^>]+id=["']divrplyfwdmsg["'][^>]*>/i,
      /<blockquote\b/i,
      /<div[^>]+class=["'][^"']*gmail_quote[^"']*["'][^>]*>/i
    ];

    return findFirstMatchIndex(html || "", markers);
  }

  function findHtmlBlockStart(html, index) {
    const value = html || "";
    const safeIndex = typeof index === "number" && index >= 0 ? index : value.length;
    const candidates = [
      value.lastIndexOf("<div", safeIndex),
      value.lastIndexOf("<p", safeIndex),
      value.lastIndexOf("<table", safeIndex),
      value.lastIndexOf("<tr", safeIndex),
      value.lastIndexOf("<td", safeIndex),
      value.lastIndexOf("<section", safeIndex),
      value.lastIndexOf("<br", safeIndex)
    ].filter((candidate) => candidate >= 0);

    return candidates.length ? Math.max.apply(null, candidates) : safeIndex;
  }

  function findHtmlSignatureClosingIndex(html, boundaryIndex) {
    const value = html || "";
    const endIndex = typeof boundaryIndex === "number" && boundaryIndex >= 0 ? boundaryIndex : value.length;
    const windowStart = Math.max(0, endIndex - 1024);
    const windowHtml = value.slice(windowStart, endIndex);
    const closingPattern = /<(?:div|p|span|td|tr)[^>]*>\s*(?:best regards|kind regards|warm regards|regards|thanks|thank you|cheers|sincerely)\s*[,:;.!-]?\s*<\/(?:div|p|span|td|tr)>/ig;
    let match = closingPattern.exec(windowHtml);
    let lastIndex = -1;

    while (match) {
      lastIndex = windowStart + match.index;
      match = closingPattern.exec(windowHtml);
    }

    return lastIndex;
  }

  function findHtmlSignatureIndex(html, userName, mailboxAddress) {
    const value = html || "";
    const explicitMarkers = [
      /<!--\s*signature[\s\S]*?-->/i,
      /<div[^>]+id=["'][^"']*(?:signature|x_signature)[^"']*["'][^>]*>/i,
      /<div[^>]+class=["'][^"']*(?:signature|gmail_signature|mso-signature)[^"']*["'][^>]*>/i,
      /<table[^>]+class=["'][^"']*signature[^"']*["'][^>]*>/i,
      /<div[^>]+data-signatureblock[^>]*>/i
    ];
    const explicitIndex = findFirstMatchIndex(value, explicitMarkers);
    if (explicitIndex >= 0) {
      return explicitIndex;
    }

    const lowerValue = value.toLowerCase();
    const nearEndStart = Math.max(0, lowerValue.length - 4096);
    const searchTerms = [
      (userName || "").trim().toLowerCase(),
      (mailboxAddress || "").trim().toLowerCase(),
      `mailto:${(mailboxAddress || "").trim().toLowerCase()}`
    ].filter(Boolean);

    let boundaryIndex = -1;
    searchTerms.forEach((term) => {
      const termIndex = lowerValue.lastIndexOf(term);
      if (termIndex >= nearEndStart) {
        const blockStart = findHtmlBlockStart(value, termIndex);
        boundaryIndex = boundaryIndex < 0 ? blockStart : Math.min(boundaryIndex, blockStart);
      }
    });

    const phonePattern = /(?:tel:|\+?\d[\d\s\-()]{6,})/ig;
    let phoneMatch = phonePattern.exec(value);
    while (phoneMatch) {
      if (phoneMatch.index >= nearEndStart) {
        const blockStart = findHtmlBlockStart(value, phoneMatch.index);
        boundaryIndex = boundaryIndex < 0 ? blockStart : Math.min(boundaryIndex, blockStart);
      }

      phoneMatch = phonePattern.exec(value);
    }

    if (boundaryIndex < 0) {
      return -1;
    }

    const closingIndex = findHtmlSignatureClosingIndex(value, boundaryIndex);
    return closingIndex >= 0 ? closingIndex : boundaryIndex;
  }

  const BLOCK_TAGS = new Set([
    "address", "article", "aside", "blockquote", "div", "dl", "fieldset", "figcaption", "figure",
    "footer", "form", "h1", "h2", "h3", "h4", "h5", "h6", "header", "hr", "li", "main", "nav",
    "ol", "p", "pre", "section", "table", "tbody", "td", "th", "tr", "ul"
  ]);

  function serializeNode(node) {
    if (!node) {
      return "";
    }

    const container = document.createElement("div");
    container.appendChild(node.cloneNode(true));
    return container.innerHTML;
  }

  function serializeNodes(nodes) {
    const container = document.createElement("div");
    (nodes || []).forEach((node) => {
      container.appendChild(node.cloneNode(true));
    });
    return container.innerHTML;
  }

  function getTopLevelBoundaryNode(root, boundaryNode) {
    if (!root || !boundaryNode) {
      return null;
    }

    let current = boundaryNode;
    while (current && current.parentNode && current.parentNode !== root) {
      current = current.parentNode;
    }

    return current && current.parentNode === root ? current : null;
  }

  function serializeElementWithInner(node, innerHtml) {
    if (!innerHtml || node.nodeType !== Node.ELEMENT_NODE) {
      return "";
    }

    const clone = node.cloneNode(false);
    clone.innerHTML = innerHtml;
    return serializeNode(clone);
  }

  function normalizeSignatureText(value) {
    return collapseInlineWhitespace(value || "")
      .trim()
      .toLowerCase()
      .replace(/[,:;.!-]+$/g, "")
      .trim();
  }

  function looksLikePhoneNumber(value) {
    return /\+?\d[\d\s\-()]{6,}/.test(value || "");
  }

  function looksLikeQuoteText(value) {
    const normalized = (value || "").replace(/\u00a0/g, " ");
    return [
      /^\s*From:\s/im,
      /^\s*Date:\s/im,
      /^\s*Sent:\s/im,
      /^\s*Subject:\s/im,
      /^\s*To:\s/im,
      /^\s*Cc:\s/im,
      /^\s*On .+wrote:\s/im,
      /^\s*-{2,}\s*Original Message\s*-{2,}\s*$/im
    ].some((pattern) => pattern.test(normalized));
  }

  function isMeaningfulBlockElement(node) {
    return !!node && node.nodeType === Node.ELEMENT_NODE && BLOCK_TAGS.has((node.tagName || "").toLowerCase());
  }

  function getMeaningfulNodeText(node) {
    return collapseInlineWhitespace(getElementText(node));
  }

  function hasMailtoLink(node) {
    return !!querySelectorSafe(node, "a[href^='mailto:']");
  }

  function hasHttpLink(node) {
    return !!querySelectorSafe(node, "a[href^='http'], a[href^='https']");
  }

  function getDocumentOrderFirst(nodes) {
    return (nodes || []).reduce((first, current) => {
      if (!current) {
        return first;
      }

      if (!first) {
        return current;
      }

      const relation = first.compareDocumentPosition(current);
      return relation & Node.DOCUMENT_POSITION_PRECEDING ? current : first;
    }, null);
  }

  function isBoundaryInside(node, boundaryNode) {
    return !!node &&
      !!boundaryNode &&
      node.nodeType === Node.ELEMENT_NODE &&
      (node === boundaryNode || node.contains(boundaryNode));
  }

  function isNodeStrictlyBefore(node, boundaryNode) {
    if (!node || !boundaryNode || node === boundaryNode) {
      return false;
    }

    if (node.nodeType === Node.ELEMENT_NODE && node.contains(boundaryNode)) {
      return false;
    }

    return !!(node.compareDocumentPosition(boundaryNode) & Node.DOCUMENT_POSITION_FOLLOWING);
  }

  function isNodeAtOrAfterBoundary(node, boundaryNode) {
    if (!boundaryNode) {
      return false;
    }

    return node === boundaryNode ||
      isBoundaryInside(node, boundaryNode) ||
      !!(node.compareDocumentPosition(boundaryNode) & Node.DOCUMENT_POSITION_PRECEDING);
  }

  function splitNodeHtml(node, boundaryNode) {
    if (!node) {
      return { before: "", after: "", found: false };
    }

    if (node === boundaryNode) {
      return { before: "", after: serializeNode(node), found: true };
    }

    if (node.nodeType !== Node.ELEMENT_NODE || !node.contains(boundaryNode)) {
      return { before: serializeNode(node), after: "", found: false };
    }

    let beforeInner = "";
    let afterInner = "";
    let found = false;

    node.childNodes.forEach((child) => {
      if (found) {
        afterInner += serializeNode(child);
        return;
      }

      const segment = splitNodeHtml(child, boundaryNode);
      beforeInner += segment.before;
      if (segment.found) {
        afterInner += segment.after;
        found = true;
      }
    });

    if (!found) {
      return { before: serializeNode(node), after: "", found: false };
    }

    return {
      before: serializeElementWithInner(node, beforeInner),
      after: serializeElementWithInner(node, afterInner),
      found: true
    };
  }

  function splitBodyHtmlAtBoundary(bodyNode, boundaryNode) {
    if (!bodyNode || !boundaryNode) {
      return {
        before: bodyNode ? bodyNode.innerHTML : "",
        after: "",
        exact: false
      };
    }

    let before = "";
    let after = "";
    let found = false;

    bodyNode.childNodes.forEach((child) => {
      if (found) {
        after += serializeNode(child);
        return;
      }

      const segment = splitNodeHtml(child, boundaryNode);
      before += segment.before;
      if (segment.found) {
        after += segment.after;
        found = true;
      }
    });

    return { before, after, exact: false };
  }

  function splitOriginalHtmlAtTopLevelBoundary(originalHtml, root, boundaryNode) {
    const topLevelBoundary = getTopLevelBoundaryNode(root, boundaryNode);
    if (!topLevelBoundary || topLevelBoundary.nodeType !== Node.ELEMENT_NODE) {
      return null;
    }

    const boundaryHtml = topLevelBoundary.outerHTML || serializeNode(topLevelBoundary);
    if (!boundaryHtml) {
      return null;
    }

    const boundaryIndex = (originalHtml || "").indexOf(boundaryHtml);
    if (boundaryIndex < 0) {
      return null;
    }

    return {
      before: originalHtml.slice(0, boundaryIndex),
      after: originalHtml.slice(boundaryIndex),
      exact: true
    };
  }

  function containsStructuredReplyMarkers(html) {
    return !!(html || "").match(/id=["']divrplyfwdmsg["']|<blockquote\b|gmail_quote|signature|mailto:|mso-signature/i);
  }

  function findQuotedThreadBoundaryElement(root) {
    if (!root) {
      return null;
    }

    const container = getComposeBlockContainer(root) || root;
    const explicit = getDocumentOrderFirst([
      root.querySelector("#divRplyFwdMsg"),
      root.querySelector("#divrplyfwdmsg"),
      root.querySelector("blockquote"),
      root.querySelector(".gmail_quote")
    ].filter(Boolean).map((node) => getTopLevelBoundaryNode(container, node) || node));

    if (explicit) {
      return explicit;
    }

    const headerWalker = document.createTreeWalker(container, NodeFilter.SHOW_ELEMENT);
    let headerNode = headerWalker.nextNode();
    while (headerNode) {
      if (looksLikeQuoteText(getElementText(headerNode))) {
        return getTopLevelBoundaryNode(container, headerNode) || headerNode;
      }

      headerNode = headerWalker.nextNode();
    }

    const hr = querySelectorSafe(container, "hr");
    if (hr) {
      return getTopLevelBoundaryNode(container, hr) || hr;
    }

    const walker = document.createTreeWalker(container, NodeFilter.SHOW_ELEMENT);
    let current = walker.nextNode();
    while (current) {
      if (isMeaningfulBlockElement(current) && looksLikeQuoteText(getMeaningfulNodeText(current))) {
        return getTopLevelBoundaryNode(container, current) || current;
      }

      current = walker.nextNode();
    }

    return null;
  }

  function getComposeBlockContainer(root) {
    if (!root) {
      return null;
    }

    const elementChildren = Array.from(root.children || []);
    const hasOuterText = Array.from(root.childNodes || []).some((node) =>
      node && node.nodeType === Node.TEXT_NODE && (node.textContent || "").trim());

    if (!hasOuterText && elementChildren.length === 1) {
      return elementChildren[0];
    }

    return root;
  }

  function getTopLevelComposeBlocks(root, limitNode) {
    if (!root) {
      return [];
    }

    const container = getComposeBlockContainer(root);
    const limitTopLevel = getTopLevelBoundaryNode(container, limitNode);
    const blocks = [];

    container.childNodes.forEach((child) => {
      if (!child || child.nodeType !== Node.ELEMENT_NODE) {
        return;
      }

      if (limitTopLevel && child === limitTopLevel) {
        return;
      }

      if (limitTopLevel && !!(child.compareDocumentPosition(limitTopLevel) & Node.DOCUMENT_POSITION_PRECEDING)) {
        return;
      }

      if (isMeaningfulBlockElement(child) ||
        matchesSelector(child, "[id*='signature' i], [class*='signature' i], [data-signatureblock]") ||
        getMeaningfulNodeText(child) ||
        hasMailtoLink(child)) {
        blocks.push(child);
      }
    });

    return blocks;
  }

  function getEligibleSignatureBlocks(root, limitNode) {
    const blocks = [];
    const walker = document.createTreeWalker(root, NodeFilter.SHOW_ELEMENT);
    let current = walker.nextNode();

    while (current) {
      if (limitNode && isNodeAtOrAfterBoundary(current, limitNode)) {
        current = walker.nextNode();
        continue;
      }

      if (isMeaningfulBlockElement(current) || matchesSelector(current, "[id*='signature' i], [class*='signature' i], [data-signatureblock]")) {
        const text = getMeaningfulNodeText(current);
        if (text || querySelectorSafe(current, "a[href^='mailto:']")) {
          blocks.push(current);
        }
      }

      current = walker.nextNode();
    }

    return blocks;
  }

  function looksLikeSignatureDetails(text, node, userName, mailboxAddress) {
    const normalized = normalizeSignatureText(text);
    const normalizedUserName = normalizeSignatureText(userName || "");
    const normalizedMailbox = normalizeSignatureText(mailboxAddress || "");
    const hasMailboxLink = Array.from(querySelectorAllSafe(node, "a[href]")).some((link) =>
      (link.getAttribute("href") || "").toLowerCase().includes((mailboxAddress || "").toLowerCase()));

    if (!normalized) {
      return false;
    }

    if (matchesSelector(node, "[id*='signature' i], [class*='signature' i], [data-signatureblock]")) {
      return true;
    }

    if (normalizedMailbox && (normalized.includes(normalizedMailbox) || hasMailboxLink)) {
      return true;
    }

    if (normalizedUserName && normalized.includes(normalizedUserName)) {
      return true;
    }

    if (looksLikePhoneNumber(normalized)) {
      return true;
    }

    if (hasHttpLink(node) || /\bwww\./i.test(text || "")) {
      return true;
    }

    if (/\b(?:cto|ceo|cfo|coo|vp|svp|director|manager|lead|architect|engineer|consultant|analyst|specialist|assistant|coordinator|president|founder|partner|office|mobile|phone|tel|fax)\b/i.test(text || "")) {
      return true;
    }

    if (normalized.includes("|")) {
      return true;
    }

    return false;
  }

  function looksLikeSupplementalSignatureLine(text) {
    const raw = (text || "").trim();
    const normalized = normalizeSignatureText(raw);
    if (!normalized) {
      return false;
    }

    if (/[.?!]$/.test(raw) && !/@|https?:\/\/|www\./i.test(raw)) {
      return false;
    }

    const wordCount = normalized.split(/\s+/).filter(Boolean).length;
    if (wordCount > 10) {
      return false;
    }

    return /\b(?:team|department|dept|group|company|inc|llc|ltd|corp|organization|sales|success|support|operations|marketing|finance|legal|hr)\b/i.test(raw) ||
      /\b(?:cto|ceo|cfo|coo|vp|svp|director|manager|lead|architect|engineer|consultant|analyst|specialist|assistant|coordinator|president|founder|partner)\b/i.test(raw);
  }

  function findSignatureBoundaryInBlocks(blocks, userName, mailboxAddress) {
    let signatureStart = null;
    let sawSignatureEvidence = false;

    for (let index = blocks.length - 1; index >= 0; index -= 1) {
      const current = blocks[index];
      const currentText = getMeaningfulNodeText(current);
      const normalizedCurrent = normalizeSignatureText(currentText);
      const isDetail = looksLikeSignatureDetails(currentText, current, userName, mailboxAddress);
      const isClosing = isClosingLine(normalizedCurrent);
      const isSupplemental = looksLikeSupplementalSignatureLine(currentText);

      if (isDetail) {
        sawSignatureEvidence = true;
        signatureStart = current;
        continue;
      }

      if (sawSignatureEvidence && (isClosing || isSupplemental)) {
        signatureStart = current;
        continue;
      }

      if (!normalizedCurrent) {
        continue;
      }

      break;
    }

    return sawSignatureEvidence ? signatureStart : null;
  }

  function findSignatureBoundaryElement(root, defaults, mailboxAddress, quoteBoundary) {
    if (!root) {
      return null;
    }

    const explicit = getDocumentOrderFirst(
      Array.from(root.querySelectorAll("[id*='signature' i], [class*='signature' i], [data-signatureblock]"))
        .filter((node) => !quoteBoundary || isNodeStrictlyBefore(node, quoteBoundary))
    );
    if (explicit) {
      return explicit;
    }

    const normalizedUserName = defaults && defaults.userName ? defaults.userName : "";
    const topLevelBoundary = findSignatureBoundaryInBlocks(
      getTopLevelComposeBlocks(root, quoteBoundary).slice(-12),
      normalizedUserName,
      mailboxAddress
    );
    if (topLevelBoundary) {
      return topLevelBoundary;
    }

    const blocks = getEligibleSignatureBlocks(root, quoteBoundary).slice(-12);

    for (let index = 0; index < blocks.length; index += 1) {
      const current = blocks[index];
      const currentText = getMeaningfulNodeText(current);
      const normalizedCurrent = normalizeSignatureText(currentText);
      if (!normalizedCurrent) {
        continue;
      }

      const next = blocks.slice(index + 1).find((node) => getMeaningfulNodeText(node));
      const nextText = getMeaningfulNodeText(next);

      if (isClosingLine(normalizedCurrent) && looksLikeSignatureDetails(nextText, next, normalizedUserName, mailboxAddress)) {
        return current;
      }

      if (looksLikeSignatureDetails(currentText, current, normalizedUserName, mailboxAddress)) {
        const previous = reverseCopy(blocks.slice(0, index)).find((node) => getMeaningfulNodeText(node));
        const previousText = normalizeSignatureText(getMeaningfulNodeText(previous));
        return previous && isClosingLine(previousText) ? previous : current;
      }
    }

    return null;
  }

  function chooseComposeBoundary(signatureBoundary, quoteBoundary) {
    if (!signatureBoundary) {
      return quoteBoundary;
    }

    if (!quoteBoundary) {
      return signatureBoundary;
    }

    return signatureBoundary.compareDocumentPosition(quoteBoundary) & Node.DOCUMENT_POSITION_FOLLOWING
      ? signatureBoundary
      : quoteBoundary;
  }

  function splitHtmlComposeSegments(body, defaults, mailboxAddress) {
    if (!body) {
      return {
        visibleTopText: "",
        userDraftText: "",
        userDraftSegment: "",
        preservedSuffix: "",
        canRewriteUserDraft: true,
        preservedSuffixExact: true
      };
    }

    const parser = new DOMParser();
    const doc = parser.parseFromString(body || "", "text/html");
    const root = doc.body;
    if (root) {
      const quoteBoundary = findQuotedThreadBoundaryElement(root);
      const signatureBoundary = findSignatureBoundaryElement(root, defaults || {}, mailboxAddress || "", quoteBoundary);
      const composeBoundary = chooseComposeBoundary(signatureBoundary, quoteBoundary);

      if (composeBoundary) {
        const topLevelBoundary = getTopLevelBoundaryNode(root, composeBoundary);
        let split = null;

        if (topLevelBoundary && topLevelBoundary === composeBoundary) {
          split = splitOriginalHtmlAtTopLevelBoundary(body, root, composeBoundary);
        }

        if (!split) {
          split = splitBodyHtmlAtBoundary(root, composeBoundary);
        }

        if ((!split.before || !split.before.trim()) && (body || "").trim()) {
          split = splitBodyHtmlAtBoundary(root, composeBoundary);
        }

        const userDraftSegment = (split.before || "").trim();
        const preservedSuffix = split.after || "";
        const userDraftText = htmlToText(userDraftSegment);
        const collapsedUserDraftText = collapseWhitespace(userDraftText);
        const shouldTreatAsEmptyDraft = !collapsedUserDraftText;

        if (shouldTreatAsEmptyDraft) {
          return {
            visibleTopText: "",
            userDraftText: "",
            userDraftSegment: "",
            preservedSuffix: body || "",
            canRewriteUserDraft: false,
            preservedSuffixExact: split.exact === true
          };
        }

        return {
          visibleTopText: collapsedUserDraftText,
          userDraftText: collapsedUserDraftText,
          userDraftSegment,
          preservedSuffix,
          canRewriteUserDraft: true,
          preservedSuffixExact: split.exact === true
        };
      }
    }

    const quoteIndex = findHtmlQuotedThreadIndex(body);
    const leading = quoteIndex >= 0 ? body.slice(0, quoteIndex) : body;
    const quotedSuffix = quoteIndex >= 0 ? body.slice(quoteIndex) : "";
    const signatureIndex = findHtmlSignatureIndex(leading, defaults && defaults.userName ? defaults.userName : "", mailboxAddress || "");
    const userDraftSegment = signatureIndex >= 0 ? leading.slice(0, signatureIndex).trim() : leading.trim();
    const signatureSuffix = signatureIndex >= 0 ? leading.slice(signatureIndex) : "";
    const preservedSuffix = `${signatureSuffix}${quotedSuffix}`;
    const visibleTopText = htmlToText(leading);
    const userDraftText = signatureIndex >= 0
      ? htmlToText(userDraftSegment)
      : stripTrailingSignature(htmlToText(leading), defaults && defaults.userName ? defaults.userName : "", mailboxAddress || "");
    const likelyHasSignature = signatureIndex >= 0 || userDraftText !== visibleTopText;
    const canRewriteUserDraft = signatureIndex >= 0 || (!likelyHasSignature && !!quotedSuffix) || (!likelyHasSignature && !preservedSuffix);
    if (!collapseWhitespace(userDraftText) && (body || "").trim()) {
      return {
        visibleTopText: "",
        userDraftText: "",
        userDraftSegment: "",
        preservedSuffix: body || "",
        canRewriteUserDraft: false,
        preservedSuffixExact: true
      };
    }

    return {
      visibleTopText,
      userDraftText,
      userDraftSegment,
      preservedSuffix,
      canRewriteUserDraft,
      preservedSuffixExact: true
    };
  }

  function splitComposeSegmentsFromBody(body, bodyType, defaults, mailboxAddress) {
    const isHtml = bodyType === Office.CoercionType.Html;
    const strippedBody = body || "";
    return isHtml
      ? splitHtmlComposeSegments(strippedBody, defaults || {}, mailboxAddress || "")
      : splitTextComposeSegments(strippedBody, defaults || {}, mailboxAddress || "");
  }

  function chooseResolvedComposeSegments(composeSegments, textComposeSegments, strippedBody, forceEmptyComposeDraft, preferEmptyComposeDraft) {
    if (preferEmptyComposeDraft) {
      return {
        visibleTopText: "",
        userDraftText: "",
        userDraftSegment: "",
        preservedSuffix: strippedBody || "",
        canRewriteUserDraft: false
      };
    }

    const htmlDraftText = collapseWhitespace(composeSegments && composeSegments.userDraftText ? composeSegments.userDraftText : "");
    const textDraftText = collapseWhitespace(textComposeSegments && textComposeSegments.userDraftText ? textComposeSegments.userDraftText : "");
    const htmlCanRewrite = !composeSegments || composeSegments.canRewriteUserDraft !== false;

    if (!htmlDraftText && textDraftText && textDraftText.length <= 2 && !htmlCanRewrite) {
      return {
        visibleTopText: "",
        userDraftText: "",
        userDraftSegment: "",
        preservedSuffix: strippedBody || "",
        canRewriteUserDraft: false
      };
    }

    if (htmlDraftText && textDraftText && htmlDraftText !== textDraftText) {
      const htmlIndexInText = textDraftText.indexOf(htmlDraftText);
      const leadingNoiseLength = htmlIndexInText >= 0 ? htmlIndexInText : -1;
      const textOnlyPrefix = leadingNoiseLength >= 0 ? textDraftText.slice(0, leadingNoiseLength) : "";
      const hasShortLeadingNoise =
        leadingNoiseLength >= 0 &&
        leadingNoiseLength <= 16 &&
        /[\u0590-\u05FF.,;:\-()\s]/.test(textOnlyPrefix);

      if (hasShortLeadingNoise) {
        return {
          visibleTopText: htmlDraftText,
          userDraftText: htmlDraftText,
          userDraftSegment: composeSegments.userDraftSegment || "",
          preservedSuffix: composeSegments.preservedSuffix || "",
          canRewriteUserDraft: htmlCanRewrite
        };
      }
    }

    const resolvedText = textDraftText || htmlDraftText;
    const useHtmlSegment = htmlDraftText && (!textDraftText || htmlDraftText === textDraftText);
    return {
      visibleTopText: textComposeSegments.visibleTopText || composeSegments.visibleTopText || "",
      userDraftText: resolvedText,
      userDraftSegment: useHtmlSegment ? (composeSegments.userDraftSegment || "") : (composeSegments.userDraftSegment || ""),
      preservedSuffix: resolvedText ? (composeSegments.preservedSuffix || "") : (strippedBody || ""),
      canRewriteUserDraft: !!resolvedText && htmlCanRewrite
    };
  }

  async function getComposeSnapshot(defaults, mode) {
    const composeContext = captureComposeContext();
    const item = ensureComposeItem();
    const bodyType = await getBodyType();
    const composeType = await getComposeType();
    const body = await getBodyByType(bodyType);
    let textBody = body;
    const subject = await getAsync(item.subject, "getAsync");
    const to = await getRecipients("to");
    const cc = await getRecipients("cc");
    const userIdentity = getUserIdentity();
    const senderIdentity = await getComposeSenderIdentity();
    const mailboxAddress = userIdentity.emailAddress || "";
    const isHtml = bodyType === Office.CoercionType.Html;
    if (isHtml) {
      try {
        textBody = await getBodyByType(Office.CoercionType.Text);
      } catch (_) {
        textBody = htmlToText(body || "");
      }
    }
    const strippedBody = body || "";
    const strippedTextBody = textBody || "";
    const composeDefaults = mergeObjects(defaults || {}, {
      userName: senderIdentity.displayName || (defaults && defaults.userName ? defaults.userName : "") || userIdentity.displayName || ""
    });
    const composeSegments = isHtml
      ? splitHtmlComposeSegments(strippedBody, composeDefaults, mailboxAddress)
      : splitTextComposeSegments(strippedBody, composeDefaults, mailboxAddress);
    const textComposeSegments = splitTextComposeSegments(strippedTextBody, composeDefaults, mailboxAddress);
    // Generated reply markers identify a replaceable area, not disposable content.
    // Outlook's current text (including edits to a previous result) is authoritative.
    for (const segments of [composeSegments, textComposeSegments]) {
      segments.userDraftText = cleanDraftMarkers(segments.userDraftText);
      segments.visibleTopText = cleanDraftMarkers(segments.visibleTopText);
    }
    const forceEmptyComposeDraft = isSignatureThreadOnlyText(
      cleanDraftMarkers(strippedTextBody),
      composeDefaults && composeDefaults.userName ? composeDefaults.userName : "",
      mailboxAddress
    );
    const preferEmptyComposeDraft =
      forceEmptyComposeDraft ||
      (composeSegments &&
        composeSegments.canRewriteUserDraft === false &&
        !composeSegments.userDraftText) ||
      (textComposeSegments &&
        textComposeSegments.canRewriteUserDraft === false &&
        !textComposeSegments.userDraftText);
    const resolvedSegments = chooseResolvedComposeSegments(
      composeSegments,
      textComposeSegments,
      strippedBody,
      forceEmptyComposeDraft,
      preferEmptyComposeDraft
    );
    // Summary language belongs to the incoming email, independently of the typed reply.
    // Use the normalized text snapshot so quote-boundary indices remain exact.
    const summaryText = (strippedTextBody || "").replace(/\r\n/g, "\n");
    const summaryQuoteIndex = findQuotedThreadIndex(summaryText);
    const summaryBodyText = summaryQuoteIndex >= 0
      ? summaryText.slice(summaryQuoteIndex).replace(/^\s*(?:On .+wrote:|-{2,}\s*Original Message\s*-{2,})\s*\n?/i, "")
      : summaryText;
    const resolvedUserDraftText = resolvedSegments.userDraftText || "";
    const resolvedVisibleTopText = resolvedSegments.visibleTopText || "";
    const resolvedPreservedSuffix = resolvedSegments.preservedSuffix || "";
    const hasGeneratedBlock = (body || "").includes(isHtml ? HTML_START : TEXT_START);
    const currentDraft = resolvedUserDraftText || null;
    const userInstruction = currentDraft
      ? "Rewrite the current draft into a polished email reply in the user's style. Keep the same meaning and concrete details from the draft, but do not keep the same wording and do not add any signature. Do not ignore the draft and do not replace it with a generic reply to the thread."
      : null;

    assertComposeContext(composeContext);
    return {
      composeContext,
      request: {
        client: {
          platform: toHostPlatform(),
          hostVersion: getDiagnostics().hostVersion || "",
          mailboxAddress,
          mailboxDisplayName: userIdentity.displayName || senderIdentity.displayName || ""
        },
        compose: {
          subject: subject || "",
          fromAddress: senderIdentity.emailAddress || mailboxAddress,
          fromDisplayName: senderIdentity.displayName || userIdentity.displayName || "",
          to,
          cc,
          bodyType: isHtml ? "html" : "text",
          bodyHtml: isHtml ? body : null,
          bodyText: textBody || null
        },
        userPreferences: {
          profileId: defaults.profileId || "default",
          tone: defaults.tone || "professional",
          length: defaults.length || "concise",
          language: defaults.language || "English"
        },
        operation: {
          mode,
          currentDraft,
          userInstruction
        }
      },
      composeState: {
        composeType,
        bodyType,
        body,
        hasGeneratedBlock,
        hostPlatform: toHostPlatform(),
        runtimeHostDetails: getRuntimeHostDetails(),
        canRewriteUserDraft: resolvedSegments.canRewriteUserDraft !== false,
        visibleTopText: resolvedVisibleTopText || "",
        userDraftText: resolvedUserDraftText,
        summaryBodyText,
        userDraftSegment: resolvedSegments.userDraftSegment || "",
        preservedSuffix: resolvedPreservedSuffix,
        preservedSuffixExact: composeSegments.preservedSuffixExact === true
      }
    };
  }

  async function getReadSnapshot(defaults) {
    const item = ensureMessageItem();
    const mailboxAddress = getMailboxAddress();
    let bodyText = "";

    if (item.body && typeof item.body.getAsync === "function") {
      try {
        bodyText = await getAsync(item.body, "getAsync", Office.CoercionType.Text);
      } catch (_) {
        try {
          bodyText = await getAsync(item.body, "getAsync");
        } catch (_) {
          bodyText = "";
        }
      }
    } else if (typeof item.body === "string") {
      bodyText = item.body;
    }

    const fromValue = await readValue(item.from);
    const from = fromValue || item.from || {};
    return {
      mailboxAddress,
      profileId: defaults && defaults.profileId ? defaults.profileId : "default",
      subject: typeof item.subject === "string"
        ? item.subject
        : ((await readValue(item.subject)) || ""),
      fromDisplayName: (from.displayName || from.name || "").trim(),
      fromAddress: (from.emailAddress || from.address || "").trim(),
      to: await readRecipientsField(item.to),
      cc: await readRecipientsField(item.cc),
      bodyText: bodyText || ""
    };
  }

  function toHostPlatform() {
    return getRuntimeHostDetails().resolvedPlatform;
  }

  function getMailboxAddress() {
    return getUserIdentity().emailAddress;
  }

  function getMailboxDisplayName() {
    return getUserIdentity().displayName;
  }

  let backendAccessKey = "";
  let backendKeyPrompt = null;

  function requestBackendAccessKey() {
    if (backendKeyPrompt) return backendKeyPrompt;
    backendKeyPrompt = new Promise((resolve) => {
      const previousFocus = document.activeElement;
      const dialog = document.createElement("dialog");
      dialog.setAttribute("aria-label", "Backend access key");
      const form = document.createElement("form");
      const label = document.createElement("label");
      label.textContent = "Backend access key";
      const input = document.createElement("input");
      input.type = "password";
      input.required = true;
      input.autocomplete = "off";
      label.append(input);
      const help = document.createElement("p");
      help.textContent = "This Maily backend requires an access key. Enter the key supplied by its owner. It is kept only while this pane is open.";
      const submit = document.createElement("button");
      submit.type = "submit";
      submit.textContent = "Continue";
      const cancel = document.createElement("button");
      cancel.type = "button";
      cancel.textContent = "Cancel";
      let completed = false;
      const finish = (key) => {
        if (completed) return;
        completed = true;
        input.value = "";
        dialog.close();
        dialog.remove();
        if (previousFocus && previousFocus.focus) previousFocus.focus();
        resolve(key);
      };
      form.addEventListener("submit", (event) => {
        event.preventDefault();
        if (input.value.trim()) finish(input.value);
      });
      cancel.addEventListener("click", () => finish(null));
      dialog.addEventListener("cancel", (event) => { event.preventDefault(); finish(null); });
      form.append(help, label, submit, cancel);
      dialog.append(form);
      document.body.append(dialog);
      dialog.showModal();
      input.focus();
    }).finally(() => { backendKeyPrompt = null; });
    return backendKeyPrompt;
  }

  async function fetchBackend(path, options, interactiveAuth = true) {
    // Never attach the backend credential to external or non-API destinations.
    const isApi = typeof path === "string" && /^\/api\//i.test(path) && !path.includes("\\");
    const withKey = (key) => mergeObjects(options, {
      redirect: "error",
      headers: mergeObjects(options.headers || {}, key ? { "X-Maily-Api-Key": key } : {})
    });
    let attemptedKey = backendAccessKey;
    let response = await fetch(path, isApi ? withKey(attemptedKey) : options);
    if (response.status !== 401 || !isApi || !interactiveAuth) return response;
    // A delayed rejection belongs to the credential used by that attempt, not
    // a newer key accepted while the request was in flight. Retry it once.
    if (backendAccessKey && backendAccessKey !== attemptedKey) {
      attemptedKey = backendAccessKey;
      if (options.signal && options.signal.aborted) throw new Error("Request cancelled.");
      response = await fetch(path, withKey(attemptedKey));
      if (response.status !== 401) return response;
      if (backendAccessKey !== attemptedKey) return response;
    }
    if (backendAccessKey === attemptedKey) backendAccessKey = "";
    const key = await requestBackendAccessKey();
    if (!key) throw new Error("Backend access cancelled.");
    if (options.signal && options.signal.aborted) throw new Error("Request cancelled.");
    // Publish the submitted session credential before retrying so concurrent
    // callers can share it while verification is in flight. Rejection clears
    // only this credential and cannot invalidate a newer submission.
    backendAccessKey = key;
    response = await fetch(path, withKey(key));
    if (response.status === 401 && backendAccessKey === key) backendAccessKey = "";
    return response;
  }

  async function callJson(path, options) {
    let response;
    const requestOptions = mergeObjects(options || {}, {
      credentials: "include",
      cache: "no-store"
    });
    requestOptions.headers = mergeObjects({ "Content-Type": "application/json" }, options && options.headers ? options.headers : {});
    try {
      response = await fetchBackend(path, requestOptions);
    } catch (error) {
      throw new Error(describeFetchError(error));
    }

    if (!response.ok) {
      const text = await response.text();
      let message = text || `Request failed: ${response.status}`;
      try {
        const payload = JSON.parse(text || "{}");
        if (payload && typeof payload.error === "string" && payload.error.trim()) {
          message = payload.error.trim();
          if (payload.requestId) {
            message = `${message} (Request: ${payload.requestId})`;
          }
        }
      } catch (_) {
      }
      throw new Error(message);
    }

    return await response.json();
  }

  async function callVoid(path, options, interactiveAuth = true) {
    let response;
    const requestOptions = mergeObjects(options || {}, {
      credentials: "include",
      cache: "no-store"
    });
    requestOptions.headers = mergeObjects({ "Content-Type": "application/json" }, options && options.headers ? options.headers : {});
    try {
      response = await fetchBackend(path, requestOptions, interactiveAuth);
    } catch (error) {
      throw new Error(describeFetchError(error));
    }

    if (!response.ok) {
      const text = await response.text();
      let message = text || `Request failed: ${response.status}`;
      try {
        const payload = JSON.parse(text || "{}");
        if (payload && typeof payload.error === "string" && payload.error.trim()) {
          message = payload.error.trim();
          if (payload.requestId) {
            message = `${message} (Request: ${payload.requestId})`;
          }
        }
      } catch (_) {
      }
      throw new Error(message);
    }
  }

  function formatReadinessMessage(readiness) {
    const userMessage = readiness && typeof readiness.userMessage === "string" ? readiness.userMessage.trim() : "";
    const recommendedAction = readiness && typeof readiness.recommendedAction === "string" ? readiness.recommendedAction.trim() : "";
    if (userMessage && recommendedAction) {
      return `${userMessage} ${recommendedAction}`;
    }

    if (userMessage) {
      return userMessage;
    }

    const codex = readiness && readiness.checks && readiness.checks.codex ? readiness.checks.codex : null;
    const details = codex && codex.details ? codex.details : {};
    const detailMessage = typeof details.message === "string" ? details.message.trim() : "";
    const detailAction = typeof details.recommendedAction === "string" ? details.recommendedAction.trim() : "";
    if (detailMessage && detailAction) {
      return `${detailMessage} ${detailAction}`;
    }

    if (detailMessage) {
      return detailMessage;
    }

    return "Maily setup is not ready yet. Reopen this pane and check the Maily setup or service configuration if the problem continues.";
  }

  async function getBackendReadiness() {
    await callJson("/api/health");
    return await callJson("/api/outlook/readiness");
  }

  async function ensureBackendAvailable(options) {
    const requireCodex = !options || options.requireCodex !== false;
    const readiness = await getBackendReadiness();
    if (requireCodex && readiness && readiness.codexReady === false) {
      throw new Error(formatReadinessMessage(readiness));
    }

    if (options && options.requireFullReady && readiness && readiness.ready === false) {
      throw new Error(formatReadinessMessage(readiness));
    }

    return readiness;
  }

  function applyComposeTypography(html, editableHtml, currentHtml) {
    html = (html || "").replaceAll(HTML_START, "").replaceAll(HTML_END, "");
    const font = { family: "Aptos, Calibri, sans-serif", size: "11pt" };
    if (typeof DOMParser === "undefined") {
      return `<div style="font-family:${font.family};font-size:${font.size}">${html}</div>`;
    }
    let foundFamily = false, foundSize = false;
    for (const source of [editableHtml, currentHtml]) {
      if (!source || (foundFamily && foundSize)) continue;
      const doc = new DOMParser().parseFromString(source, "text/html");
      const rules = [];
      for (const sheet of doc.querySelectorAll("style")) {
        for (const match of sheet.textContent.matchAll(/([^{}]+)\{([^{}]*)\}/g)) {
          for (const selector of match[1].split(",")) {
            // Outlook commonly declares p.MsoNormal in a document stylesheet.
            // Only simple tag/class selectors are needed; never attach styles.
            const simple = selector.trim();
            if (!/^(?:[a-z][a-z0-9]*)?(?:\.[a-z_][a-z0-9_-]*)?$/i.test(simple) || !simple) continue;
            const declaration = doc.createElement("span").style;
            declaration.cssText = match[2];
            rules.push({ selector: simple, family: declaration.fontFamily, size: declaration.fontSize,
              specificity: simple.includes(".") ? 10 + (simple[0] === "." ? 0 : 1) : 1 });
          }
        }
      }
      rules.sort((a, b) => a.specificity - b.specificity);
      // Never import source markup, CSS rules, images or signature decoration.
      doc.querySelectorAll('script,style,link,meta,[id="Signature"],[id="signature"]').forEach(e => e.remove());
      const elements = Array.from(doc.body.querySelectorAll("*"));
      const textElements = elements.filter(e => !e.textContent.trim() || Array.from(e.childNodes).some(n => n.nodeType === 3 && n.textContent.trim()));
      for (const element of [...textElements, ...elements, doc.body]) {
        for (let node = element; node && node.style; node = node.parentElement) {
          let ruleFamily = "", ruleSize = "";
          for (const rule of rules) {
            if (node.matches(rule.selector)) {
              if (rule.family) ruleFamily = rule.family;
              if (rule.size) ruleSize = rule.size;
            }
          }
          const family = node.style.fontFamily || ruleFamily || (node.tagName === "FONT" ? node.getAttribute("face") || "" : "");
          const size = node.style.fontSize || ruleSize;
          // Copy only bounded literal font values. No URLs, expressions, CSS
          // variables, selectors or markup can cross from a message into output.
          if (!foundFamily && family.length <= 160 && /^[\p{L}\p{N}\s,'"-]+$/u.test(family) && !/^(inherit|initial|unset|revert)$/i.test(family)) {
            font.family = family; foundFamily = true;
          }
          if (!foundSize && /^(?:\d{1,2}(?:\.\d{1,2})?)(?:pt|px)$/.test(size) && parseFloat(size) >= 6 && parseFloat(size) <= 72) {
            font.size = size; foundSize = true;
          }
        }
        if (foundFamily && foundSize) break;
      }
    }
    const generated = new DOMParser().parseFromString(html || "", "text/html");
    const wrapper = generated.createElement("div");
    while (generated.body.firstChild) wrapper.appendChild(generated.body.firstChild);
    generated.body.appendChild(wrapper);
    // Set every generated element so provider/default inline fonts cannot win
    // against the compose font through CSS inheritance.
    for (const element of [wrapper, ...wrapper.querySelectorAll("*")]) {
      element.style.fontFamily = font.family;
      element.style.fontSize = font.size;
      if (element.tagName === "FONT") { element.removeAttribute("face"); element.removeAttribute("size"); }
    }
    return generated.body.innerHTML;
  }

  function joinComposeSegments(payload, suffix, isHtml) {
    if (!suffix) {
      return payload;
    }

    if (isHtml || payload.endsWith("\n") || suffix.startsWith("\n")) {
      return `${payload}${suffix}`;
    }

    return `${payload}\n${suffix}`;
  }

  function canUndoGeneratedDraft() {
    if (!undoRecord || composeMutationBusy) return false;
    try { assertComposeContext(undoRecord.context); return true; }
    catch (_) { undoRecord = null; return false; }
  }

  async function rememberVerifiedWrite(context, bodyType, before, intendedAfter) {
    // Mail content lives only in this pane's closure. If Outlook normalizes the
    // write or the user edits before confirmation, do not guess an Undo target.
    try {
      if (await getBodyType() !== bodyType) return;
      const after = await getBodyByType(bodyType);
      assertComposeContext(context);
      if (after === intendedAfter) undoRecord = { context, bodyType, before, after };
    } catch (_) { undoRecord = null; }
  }

  async function undoGeneratedDraft() {
    if (composeMutationBusy) throw new Error("Wait for the current draft change to finish before using Undo.");
    const record = undoRecord;
    undoRecord = null;
    if (!record) throw new Error("There is no verified Maily change to undo in this pane.");
    composeMutationBusy = true;
    try {
      assertComposeContext(record.context);
      if (await getBodyType() !== record.bodyType || await getBodyByType(record.bodyType) !== record.after ||
          await getBodyByType(record.bodyType) !== record.after || await getBodyType() !== record.bodyType) {
        throw new Error("The draft changed after Maily wrote it. Undo was stopped to preserve your newer edits.");
      }
      assertComposeContext(record.context);
      // Consume before mutation: a failed callback may conceal a partial write.
      await setBody(record.before, record.bodyType, record.context.item);
      return { mode: "undo" };
    } finally { composeMutationBusy = false; }
  }

  async function insertGeneratedDraft(draft, options) {
    if (composeMutationBusy) throw new Error("Wait for the current draft change to finish.");
    composeMutationBusy = true;
    undoRecord = null;
    try { return await insertGeneratedDraftCore(draft, options); }
    finally { composeMutationBusy = false; }
  }

  async function insertGeneratedDraftCore(draft, options) {
    const expectedContext = options && options.expectedContext ? options.expectedContext : captureComposeContext();
    assertComposeContext(expectedContext);
    const replaceExisting = !!(options && options.replaceExisting);
    const rewriteComposeBody = !!(options && options.rewriteComposeBody);
    const preservedSuffix = options && options.preservedSuffix ? options.preservedSuffix : "";
    const originalUserDraftSegment = options && options.originalUserDraftSegment ? options.originalUserDraftSegment : "";
    const composeDefaults = options && options.composeDefaults ? options.composeDefaults : {};
    const bodyType = await getBodyType();
    const isHtml = bodyType === Office.CoercionType.Html;
    const tokenStart = isHtml ? HTML_START : TEXT_START;
    const tokenEnd = isHtml ? HTML_END : TEXT_END;
    const currentBody = await getBodyByType(bodyType);
    assertComposeContext(expectedContext);
    const payload = isHtml
      ? wrapForHtml(applyComposeTypography(applyHtmlDirection(draft.html || "", draft.text || ""), originalUserDraftSegment, currentBody))
      : wrapForText(draft.text || "");
    const changedMessage = "The draft changed while Maily was working. Nothing was replaced; generate again from the current draft.";
    if (options && typeof options.expectedBody === "string" &&
        (currentBody !== options.expectedBody || (options.expectedBodyType && options.expectedBodyType !== bodyType))) {
      throw new Error(changedMessage);
    }
    if (rewriteComposeBody && (toHostPlatform() !== "outlook-classic" || (options && options.expectedContext)) &&
        (!options || typeof options.expectedBody !== "string")) {
      throw new Error("The original draft snapshot is unavailable. Generate again before rewriting.");
    }
    async function setUnchangedBody(content) {
      if (await getBodyByType(bodyType) !== currentBody) throw new Error(changedMessage);
      assertComposeContext(expectedContext);
      await setBody(content, bodyType, expectedContext.item);
      await rememberVerifiedWrite(expectedContext, bodyType, currentBody, content);
    }

    if (replaceExisting) {
      const replaced = replaceBetween(currentBody || "", tokenStart, tokenEnd, payload);
      if (replaced.replaced) {
        await setUnchangedBody(replaced.output);
        return { mode: "replace" };
      }
    }

    if (rewriteComposeBody) {
      // Generated taskpane writes use the captured Office item: an unsaved draft
      // has no stable EntryID for safely selecting a native compose window.
      if (toHostPlatform() === "outlook-classic" && !(options && options.expectedContext)) {
        try {
          assertComposeContext(expectedContext);
          const bridgeResult = await rewriteActiveComposeWithBridge(draft, {
            subject: options && options.subject ? options.subject : "",
            userName: composeDefaults && composeDefaults.userName ? composeDefaults.userName : getMailboxDisplayName(),
            mailboxAddress: getMailboxAddress(),
            seedText: options && options.seedText ? options.seedText : "",
            composeType: options && options.composeType ? options.composeType : "unknown"
          });
          return { mode: "replace-user-draft", strategy: "bridge", bridgeResult };
        } catch (bridgeError) {
          // Failure may mean the draft changed or native mutation has an unknown
          // outcome. A second writer must never overwrite it speculatively.
          throw new Error("Outlook could not safely rewrite this draft. It may have changed; generate again from the current draft.");
        }
      }

      const sourceBody = currentBody || "";
      if (originalUserDraftSegment) {
        const prefixIndex = sourceBody.indexOf(originalUserDraftSegment);
        if (prefixIndex >= 0 && !sourceBody.slice(0, prefixIndex).trim()) {
          await setUnchangedBody(`${sourceBody.slice(0, prefixIndex)}${payload}${sourceBody.slice(prefixIndex + originalUserDraftSegment.length)}`);
          return { mode: "replace-user-draft", strategy: "office-js" };
        }
      }

      const fallbackSegments = splitComposeSegmentsFromBody(
        sourceBody,
        bodyType,
        composeDefaults,
        getMailboxAddress()
      );
      const fallbackSuffix = fallbackSegments && fallbackSegments.canRewriteUserDraft !== false
        ? (fallbackSegments.preservedSuffix || "")
        : preservedSuffix;

      if (!fallbackSuffix && preservedSuffix) {
        await setUnchangedBody(joinComposeSegments(payload, preservedSuffix, isHtml));
        return { mode: "replace-user-draft", strategy: "office-js" };
      }

      await setUnchangedBody(joinComposeSegments(payload, fallbackSuffix, isHtml));
      return { mode: "replace-user-draft", strategy: "office-js" };
    }

    assertComposeContext(expectedContext);
    await prependBody(payload, bodyType, expectedContext.item);
    await rememberVerifiedWrite(expectedContext, bodyType, currentBody, payload + currentBody);
    return { mode: "prepend" };
  }

  function copyText(text) {
    if (!text) {
      return;
    }

    if (navigator.clipboard && navigator.clipboard.writeText) {
      navigator.clipboard.writeText(text).catch(() => {});
    }
  }

  async function logClientEvent(eventName, status, detail, pane) {
    try {
      await callVoid("/api/outlook/client-log", {
        method: "POST",
        body: JSON.stringify({
          eventName: ["bridge-rewrite-fallback", "daily-briefing-complete", "daily-briefing-failed", "daily-briefing-open", "daily-briefing-start-failed", "generate-complete", "generate-failed", "generate-start", "insert-again", "search-open", "search-refine", "search-start", "search-start-failed", "settings-load", "settings-pane-load", "settings-pane-open", "settings-pane-start-failed", "settings-save", "style-refresh", "summary-complete", "summary-failed", "summary-open", "summary-start-failed", "taskpane-auto-hide", "taskpane-open", "taskpane-start-failed"].includes(eventName) ? eventName : "unknown",
          pane: ["taskpane", "summary", "search", "daily-briefing", "settings"].includes(pane) ? pane : "unknown",
          status: ["ok", "error", "warning", "fallback", "info"].includes(status) ? status : "info"
        })
      }, false);
    } catch (_) {
      // Avoid breaking compose flows if telemetry fails.
    }
  }

  window.MailyOffice = {
    officeReady,
    callJson,
    getComposeSnapshot,
    captureComposeContext,
    assertComposeContext,
    getReadSnapshot,
    insertGeneratedDraft,
    canUndoGeneratedDraft,
    undoGeneratedDraft,
    copyText,
    getMailboxAddress,
    getMailboxDisplayName,
    logClientEvent,
    getBackendReadiness,
    ensureBackendAvailable,
    containsHebrew,
    detectPreferredLanguage,
    applyDirectionToElement,
    getDirectionalText,
    getDirectionAttributes,
    applyHtmlDirection,
    applyComposeTypography
  };

  window.MailyOffice.pruneLegacyPaneButtons = function pruneLegacyPaneButtons() {
    if (typeof document === "undefined" || !document.getElementById) {
      return;
    }

    ["regenerateBtn", "replaceBtn", "copyBtn", "closeSettingsBtn"].forEach((id) => {
      const element = document.getElementById(id);
      if (element) {
        element.remove();
      }
    });
  };

  window.MailyOffice.__test = {
    isSignatureThreadOnlyText,
    splitHtmlComposeSegments,
    splitTextComposeSegments,
    htmlToText,
    containsHebrew,
    getDirectionalText
  };
})();
