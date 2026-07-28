// ==UserScript==
// @name         BHVP – Volet Outlook vers ChatGPT
// @namespace    bhvp-outlook-chatgpt
// @version      1.2.3
// @description  Envoie le courrier visible vers ChatGPT, récupère automatiquement sa réponse et peut l’insérer dans un brouillon Outlook.
// @homepageURL  https://github.com/GLAD1981/Tampermonkey
// @updateURL    https://raw.githubusercontent.com/GLAD1981/Tampermonkey/main/BHVP-Volet-Outlook.user.js
// @downloadURL  https://raw.githubusercontent.com/GLAD1981/Tampermonkey/main/BHVP-Volet-Outlook.user.js
// @match        https://outlook.office.com/*
// @match        https://outlook.cloud.microsoft/*
// @match        https://outlook.office365.com/*
// @match        https://outlook.live.com/*
// @match        https://chatgpt.com/*
// @grant        GM_setValue
// @grant        GM_getValue
// @grant        GM_deleteValue
// @grant        GM_addValueChangeListener
// @grant        GM_setClipboard
// @run-at       document-idle
// @noframes
// ==/UserScript==

(() => {
  'use strict';

  const VERSION = '1.2.3';
  const STORAGE_PENDING = 'bhvp_pending_mail_v2';
  const STORAGE_SETTINGS = 'bhvp_panel_settings_v2';
  const STORAGE_ACTIVE_REQUEST = 'bhvp_active_request_v1';
  const STORAGE_RESPONSE = 'bhvp_chatgpt_response_v1';
  const HOST_ID = 'bhvp-outlook-panel-host';
  const NOTICE_ID = 'bhvp-outlook-panel-notice';
  const CHATGPT_WINDOW_NAME = 'BHVP_CHATGPT';

  const DEFAULT_SETTINGS = {
    profile: 'generic',
    collapsed: false,
    includeThread: true,
    autoSend: false,
    popupWindow: true,
    reserveSpace: true,
    width: 350
  };

  let lastPointerTarget = null;
  let currentMail = null;
  let currentResponse = null;
  let refreshTimer = null;
  const layoutState = { root: null, body: null, rootStyles: null, bodyStyles: null };

  function gmGet(key, fallback = null) {
    try {
      return Promise.resolve(GM_getValue(key, fallback));
    } catch (_) {
      return Promise.resolve(fallback);
    }
  }

  function gmSet(key, value) {
    try {
      return Promise.resolve(GM_setValue(key, value));
    } catch (error) {
      return Promise.reject(error);
    }
  }

  function gmDelete(key) {
    try {
      return Promise.resolve(GM_deleteValue(key));
    } catch (_) {
      return Promise.resolve();
    }
  }

  function cleanText(value) {
    return String(value || '')
      .replace(/\u00a0/g, ' ')
      .replace(/\r\n?/g, '\n')
      .replace(/[ \t]+\n/g, '\n')
      .replace(/\n[ \t]+/g, '\n')
      .replace(/[ \t]{3,}/g, '  ')
      .replace(/\n{4,}/g, '\n\n\n')
      .trim();
  }

  function isVisible(element) {
    if (!(element instanceof HTMLElement)) return false;
    if (element.closest(`#${HOST_ID}`)) return false;
    const style = getComputedStyle(element);
    const rect = element.getBoundingClientRect();
    return style.display !== 'none' &&
      style.visibility !== 'hidden' &&
      Number(style.opacity || 1) !== 0 &&
      rect.width > 180 &&
      rect.height > 35 &&
      rect.bottom > 0 &&
      rect.right > 0 &&
      rect.top < innerHeight &&
      rect.left < innerWidth;
  }

  function uniqueElements(elements) {
    return [...new Set(elements)];
  }

  function visibleElements(selector, root = document) {
    try {
      return [...root.querySelectorAll(selector)].filter(isVisible);
    } catch (_) {
      return [];
    }
  }

  function candidateScore(element, text) {
    const rect = element.getBoundingClientRect();
    const viewportCenterX = innerWidth / 2;
    const viewportCenterY = innerHeight / 2;
    const centerX = rect.left + rect.width / 2;
    const centerY = rect.top + rect.height / 2;
    let score = Math.min(text.length, 12000) / 80;

    const role = (element.getAttribute('role') || '').toLowerCase();
    const aria = (element.getAttribute('aria-label') || '').toLowerCase();
    const testId = (element.getAttribute('data-testid') || '').toLowerCase();

    if (role === 'document') score += 80;
    if (/message body|corps du message|contenu du message/.test(aria)) score += 100;
    if (/message.?body|mail.?body/.test(testId)) score += 80;
    if (centerX > viewportCenterX) score += 30;
    if (Math.abs(centerY - viewportCenterY) < innerHeight * 0.35) score += 15;
    if (rect.width > innerWidth * 0.35) score += 10;
    if (rect.width > innerWidth * 0.92 || rect.height > innerHeight * 0.95) score -= 140;
    if (text.length > 80000) score -= 100;
    if (element.querySelectorAll('button, input, [role="button"]').length > 30) score -= 50;

    if (lastPointerTarget instanceof Element &&
        (element.contains(lastPointerTarget) || lastPointerTarget.contains(element))) {
      score += 180;
    }

    return score;
  }

  function locateMessageBody() {
    const selectors = [
      '[role="document"]',
      '[aria-label*="Message body" i]',
      '[aria-label*="Corps du message" i]',
      '[aria-label*="Contenu du message" i]',
      '[data-testid*="message-body" i]',
      '[data-testid*="messageBody" i]',
      '[data-app-section="MailReadCompose"] [contenteditable="false"]',
      '[data-app-section="MailReadCompose"]'
    ];

    const elements = uniqueElements(selectors.flatMap((selector) => visibleElements(selector)));
    const candidates = elements
      .map((element) => ({
        element,
        text: cleanText(element.innerText || element.textContent || '')
      }))
      .filter(({ text }) => text.length >= 20 && text.length <= 180000)
      .map((candidate) => ({
        ...candidate,
        score: candidateScore(candidate.element, candidate.text)
      }))
      .sort((a, b) => b.score - a.score);

    return candidates[0] || null;
  }

  function headingIsPlausible(text) {
    if (!text || text.length < 2 || text.length > 500) return false;
    return !/^(boîte de réception|inbox|courrier|mail|nouveau courrier|new mail|répondre|reply|transférer|forward|supprimer|delete|archive)$/i.test(text.trim());
  }

  function locateSubject(bodyElement) {
    const roots = [];
    let node = bodyElement;
    for (let i = 0; node && i < 10; i += 1, node = node.parentElement) roots.push(node);
    roots.push(document);

    const selectors = [
      '[data-testid*="subject" i]',
      '[aria-label*="Objet" i]',
      '[aria-label*="Subject" i]',
      'h1', 'h2', 'h3',
      '[role="heading"]'
    ];

    for (const root of roots) {
      for (const selector of selectors) {
        const found = visibleElements(selector, root)
          .map((element) => cleanText(element.innerText || element.textContent || element.getAttribute('aria-label') || ''))
          .filter(headingIsPlausible)
          .sort((a, b) => a.length - b.length);
        if (found.length) return found[0];
      }
    }
    return '(objet non détecté)';
  }

  function locateSender(bodyElement) {
    const roots = [];
    let node = bodyElement;
    for (let i = 0; node && i < 8; i += 1, node = node.parentElement) roots.push(node);

    const selectors = [
      '[data-testid*="sender" i]',
      '[aria-label^="De :" i]',
      '[aria-label^="From:" i]',
      '[title*="@"]',
      'button[aria-label*="@"]'
    ];

    for (const root of roots) {
      for (const selector of selectors) {
        const texts = visibleElements(selector, root)
          .map((element) => cleanText(
            element.getAttribute('aria-label') ||
            element.getAttribute('title') ||
            element.innerText ||
            element.textContent || ''
          ))
          .filter((text) => text.length >= 3 && text.length <= 300);
        if (texts.length) return texts[0];
      }
    }
    return '(expéditeur non détecté)';
  }

  function stripPreviousThread(text) {
    const markers = [
      /^-{2,}\s*Message d'origine\s*-{2,}$/im,
      /^-{2,}\s*Original Message\s*-{2,}$/im,
      /^De\s*:\s.+$/im,
      /^From\s*:\s.+$/im,
      /^Le .+ a écrit\s*:\s*$/im,
      /^On .+ wrote\s*:\s*$/im
    ];

    let cutAt = text.length;
    for (const marker of markers) {
      const match = marker.exec(text);
      if (match && match.index > 40) cutAt = Math.min(cutAt, match.index);
    }
    return cleanText(text.slice(0, cutAt));
  }

  function extractCurrentMail(latestOnly = false) {
    const found = locateMessageBody();
    if (!found) return null;

    let body = found.text;
    if (latestOnly) body = stripPreviousThread(body);

    return {
      subject: locateSubject(found.element),
      sender: locateSender(found.element),
      body,
      includesThread: !latestOnly,
      capturedAt: new Date().toISOString()
    };
  }

  function profileLabel(profile) {
    return profile === 'personal'
      ? 'boîte professionnelle personnelle'
      : 'boîte générique BHVP';
  }

  function actionInstruction(action) {
    switch (action) {
      case 'reply':
        return 'Rédige une réponse au courriel ci-dessous.';
      case 'summary':
        return 'Résume le courriel ou le fil ci-dessous.';
      case 'learn':
        return `Observe les réponses rédigées par l’utilisateur dans ce fil. Relève uniquement les règles de rédaction réellement visibles et récurrentes. Distingue les règles certaines des simples hypothèses. Mémorise les règles suffisamment claires pour les prochaines réponses professionnelles.`;
      case 'raw':
      default:
        return '';
    }
  }

  function buildPrompt(mail, action, settings, extraInstruction = '') {
    const extra = cleanText(extraInstruction);
    const baseInstruction = actionInstruction(action);
    const instruction = [baseInstruction, extra].filter(Boolean).join('\n');
    const parts = [];

    if (instruction) parts.push(instruction);
    parts.push(`SOURCE OUTLOOK : ${profileLabel(settings.profile)}`);
    parts.push(`OBJET : ${mail.subject}
EXPÉDITEUR : ${mail.sender}

COURRIEL OU FIL VISIBLE :
${mail.body}`);

    return parts.join('\n\n');
  }

  function showPageNotice(message, error = false) {
    let notice = document.getElementById(NOTICE_ID);
    if (!notice) {
      notice = document.createElement('div');
      notice.id = NOTICE_ID;
      Object.assign(notice.style, {
        position: 'fixed',
        left: '18px',
        bottom: '18px',
        zIndex: '2147483647',
        maxWidth: '420px',
        padding: '10px 13px',
        borderRadius: '8px',
        color: '#fff',
        font: '13px/1.4 system-ui, sans-serif',
        boxShadow: '0 5px 22px rgba(0,0,0,.28)'
      });
      document.documentElement.appendChild(notice);
    }
    notice.style.background = error ? '#9f1d20' : '#245b3a';
    notice.textContent = message;
    notice.hidden = false;
    clearTimeout(notice._timer);
    notice._timer = setTimeout(() => { notice.hidden = true; }, 4200);
  }

  function openChatGPTWindow(settings) {
    if (!settings.popupWindow) {
      return window.open('https://chatgpt.com/', CHATGPT_WINDOW_NAME);
    }

    const width = Math.min(650, Math.max(480, Math.round(screen.availWidth * 0.34)));
    const height = Math.max(650, screen.availHeight - 40);
    const left = Math.max(0, screen.availLeft + screen.availWidth - width);
    const top = Math.max(0, screen.availTop + 10);
    const features = [
      'popup=yes',
      `width=${width}`,
      `height=${height}`,
      `left=${left}`,
      `top=${top}`,
      'resizable=yes',
      'scrollbars=yes'
    ].join(',');

    return window.open('https://chatgpt.com/', CHATGPT_WINDOW_NAME, features);
  }

  async function sendMail(action, ui) {
    const settings = ui.readSettings();
    const includeThread = action === 'learn' ? true : settings.includeThread;

    // Utilise la capture déjà affichée dans le volet. Ne rescane pas Outlook au
    // moment du clic : l'interface peut alors ne laisser dans le DOM que le
    // message actuellement visible et amputer un fil qui avait été correctement
    // capturé quelques secondes auparavant.
    let mail = currentMail;
    if (!mail || Boolean(mail.includesThread) !== Boolean(includeThread)) {
      mail = extractCurrentMail(!includeThread);
      if (mail) {
        currentMail = mail;
        ui.renderMail(mail);
      }
    }

    if (!mail) {
      ui.setStatus('Courrier non détecté. Ouvre le message, survole son texte, puis clique sur Actualiser.', true);
      return;
    }

    const payload = {
      id: `${Date.now()}-${Math.random().toString(36).slice(2)}`,
      createdAt: Date.now(),
      prompt: buildPrompt(mail, action, settings, ui.extra.value),
      autoSend: settings.autoSend,
      action,
      subject: mail.subject,
      sender: mail.sender,
      profile: settings.profile
    };

    currentResponse = null;
    ui.renderResponse(null);
    await gmDelete(STORAGE_RESPONSE);
    await gmSet(STORAGE_ACTIVE_REQUEST, payload);
    await gmSet(STORAGE_PENDING, payload);
    const opened = openChatGPTWindow(settings);
    if (!opened) {
      ui.setStatus('Chrome a bloqué l’ouverture de ChatGPT. Autorise les fenêtres surgissantes pour Outlook.', true);
      return;
    }

    ui.setStatus(settings.autoSend
      ? 'Courrier envoyé vers ChatGPT.'
      : 'Courrier placé dans ChatGPT pour relecture avant envoi.');
  }

  function normalizedIdentity(value) {
    return cleanText(value)
      .toLocaleLowerCase('fr-FR')
      .replace(/[\s\u25a1\u25a0]+/g, ' ')
      .trim();
  }

  function sameDetectedMessage(a, b) {
    if (!a || !b) return false;
    const subjectA = normalizedIdentity(a.subject);
    const subjectB = normalizedIdentity(b.subject);
    if (subjectA && subjectB && subjectA !== '(objet non détecté)' && subjectB !== '(objet non détecté)') {
      return subjectA === subjectB;
    }
    const senderA = normalizedIdentity(a.sender);
    const senderB = normalizedIdentity(b.sender);
    return Boolean(senderA && senderB && senderA === senderB);
  }

  function shouldAcceptAutomaticRefresh(candidate) {
    if (!currentMail) return true;
    if (!sameDetectedMessage(currentMail, candidate)) return true;
    if (Boolean(currentMail.includesThread) !== Boolean(candidate.includesThread)) return true;

    // Outlook virtualise les conversations : après un clic, une perte de focus,
    // une capture d'écran ou un redimensionnement, le DOM peut ne contenir que
    // la carte visible. On refuse donc qu'un rafraîchissement automatique
    // remplace une capture par une version nettement plus courte du même fil.
    const currentLength = currentMail.body?.length || 0;
    const candidateLength = candidate.body?.length || 0;
    return candidateLength >= currentLength * 0.85;
  }

  function scheduleRefresh(ui, force = false) {
    clearTimeout(refreshTimer);
    refreshTimer = setTimeout(() => {
      const settings = ui.readSettings();
      const mail = extractCurrentMail(!settings.includeThread);
      if (!mail) {
        if (force) ui.setStatus('Aucun courrier ouvert détecté.', true);
        return;
      }

      if (!force && !shouldAcceptAutomaticRefresh(mail)) {
        return;
      }

      currentMail = mail;
      ui.renderMail(mail);
      if (force) {
        ui.setStatus(`Courrier actualisé et figé (${mail.body.length.toLocaleString('fr-FR')} caractères).`);
      }
    }, force ? 20 : 700);
  }



  function rememberLayoutStyles(root) {
    if (layoutState.root === root && layoutState.rootStyles) return;
    restoreOutlookLayout();
    layoutState.root = root || null;
    layoutState.body = document.body || null;
    if (root) {
      layoutState.rootStyles = {
        width: root.style.width,
        maxWidth: root.style.maxWidth,
        right: root.style.right,
        transition: root.style.transition
      };
    }
    if (layoutState.body) {
      layoutState.bodyStyles = {
        paddingRight: layoutState.body.style.paddingRight,
        boxSizing: layoutState.body.style.boxSizing,
        overflowX: layoutState.body.style.overflowX
      };
    }
  }

  function restoreOutlookLayout() {
    if (layoutState.root && layoutState.rootStyles) {
      Object.assign(layoutState.root.style, layoutState.rootStyles);
    }
    if (layoutState.body && layoutState.bodyStyles) {
      Object.assign(layoutState.body.style, layoutState.bodyStyles);
    }
    layoutState.root = null;
    layoutState.body = null;
    layoutState.rootStyles = null;
    layoutState.bodyStyles = null;
  }

  function findOutlookAppRoot() {
    if (!document.body) return null;
    const children = [...document.body.children].filter((element) =>
      element instanceof HTMLElement &&
      element.id !== HOST_ID &&
      isVisible(element)
    );
    return children
      .map((element) => {
        const rect = element.getBoundingClientRect();
        return { element, rect, area: rect.width * rect.height };
      })
      .filter(({ rect }) => rect.width >= innerWidth * 0.50 && rect.height >= innerHeight * 0.72)
      .sort((a, b) => b.area - a.area)[0]?.element || null;
  }

  function applyOutlookLayout(settings, panel) {
    const width = Math.max(300, Math.min(Number(settings.width) || 350, 560));
    panel.style.setProperty('--bhvp-width', `${width}px`);

    const shouldReserve = Boolean(settings.reserveSpace) && !panel.classList.contains('collapsed');
    if (!shouldReserve) {
      restoreOutlookLayout();
      return;
    }

    const root = layoutState.root?.isConnected ? layoutState.root : findOutlookAppRoot();
    if (layoutState.root !== root) rememberLayoutStyles(root);

    if (root) {
      const computed = getComputedStyle(root);
      root.style.transition = 'width .12s ease, right .12s ease';
      if ((computed.position === 'fixed' || computed.position === 'absolute') &&
          computed.left !== 'auto' && computed.right !== 'auto') {
        root.style.right = `${width}px`;
        root.style.width = 'auto';
        root.style.maxWidth = 'none';
      } else {
        root.style.width = `calc(100% - ${width}px)`;
        root.style.maxWidth = `calc(100% - ${width}px)`;
      }
      return;
    }

    if (document.body) {
      if (!layoutState.bodyStyles) rememberLayoutStyles(null);
      document.body.style.boxSizing = 'border-box';
      document.body.style.paddingRight = `${width}px`;
      document.body.style.overflowX = 'hidden';
    }
  }

  function locateOutlookComposeEditor() {
    const selectors = [
      '[data-app-section*="MailCompose" i] [contenteditable="true"]',
      '[contenteditable="true"][aria-label*="Corps du message" i]',
      '[contenteditable="true"][aria-label*="Message body" i]',
      '[contenteditable="true"][aria-label*="Contenu du message" i]',
      '[contenteditable="true"][role="textbox"]'
    ];

    const candidates = uniqueElements(selectors.flatMap((selector) => visibleElements(selector)))
      .filter((element) => !element.closest(`#${HOST_ID}`))
      .map((element) => {
        const rect = element.getBoundingClientRect();
        const aria = (element.getAttribute('aria-label') || '').toLowerCase();
        let score = 0;
        if (/corps du message|message body|contenu du message/.test(aria)) score += 200;
        if (element.closest('[data-app-section*="MailCompose" i]')) score += 150;
        if (element.getAttribute('role') === 'textbox') score += 30;
        if (rect.width > 350) score += 20;
        if (rect.height > 80) score += 20;
        if (/rechercher|search/.test(aria)) score -= 500;
        return { element, score };
      })
      .sort((a, b) => b.score - a.score);

    return candidates[0]?.element || null;
  }

  function buildPlainTextFragment(text) {
    const fragment = document.createDocumentFragment();
    const lines = String(text).replace(/\r\n?/g, '\n').split('\n');
    lines.forEach((line, index) => {
      if (index) fragment.appendChild(document.createElement('br'));
      if (line) fragment.appendChild(document.createTextNode(line));
    });
    fragment.appendChild(document.createElement('br'));
    fragment.appendChild(document.createElement('br'));
    return fragment;
  }

  function insertIntoOutlookDraft(text) {
    const editor = locateOutlookComposeEditor();
    if (!editor) {
      return {
        ok: false,
        message: 'Aucun brouillon Outlook n’est ouvert. Clique sur Répondre dans le courrier, puis réessaie.'
      };
    }

    editor.focus();
    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(editor);
    range.collapse(true);
    selection.removeAllRanges();
    selection.addRange(range);

    let inserted = false;
    try {
      inserted = document.execCommand('insertText', false, `${text}\n\n`);
    } catch (_) {
      inserted = false;
    }

    if (!inserted) {
      try {
        editor.insertBefore(buildPlainTextFragment(text), editor.firstChild);
        inserted = true;
      } catch (_) {
        inserted = false;
      }
    }

    if (!inserted) {
      return { ok: false, message: 'Outlook a refusé l’insertion. Utilise le bouton Copier comme solution de secours.' };
    }

    try {
      editor.dispatchEvent(new InputEvent('input', {
        bubbles: true,
        inputType: 'insertText',
        data: text
      }));
    } catch (_) {
      editor.dispatchEvent(new Event('input', { bubbles: true }));
    }
    editor.dispatchEvent(new Event('change', { bubbles: true }));
    return { ok: true };
  }

  function createPanelUi(settings) {
    const host = document.createElement('div');
    host.id = HOST_ID;
    host.style.position = 'fixed';
    host.style.top = '0';
    host.style.right = '0';
    host.style.height = '100vh';
    host.style.zIndex = '2147483646';
    host.style.pointerEvents = 'none';

    const shadow = host.attachShadow({ mode: 'open' });

    // Outlook impose Trusted Types et bloque les affectations à innerHTML.
    // Toute l'interface est donc construite avec createElement/textContent.
    const style = document.createElement('style');
    style.textContent = `
      :host { all: initial; }
      * { box-sizing: border-box; }
      #panel {
        pointer-events: auto;
        width: var(--bhvp-width, 350px);
        height: 100vh;
        background: #f7f7f8;
        color: #202124;
        border-left: 1px solid #c9cdd2;
        box-shadow: -4px 0 18px rgba(0,0,0,.13);
        font: 13px/1.4 system-ui, -apple-system, BlinkMacSystemFont, "Segoe UI", sans-serif;
        display: flex;
        flex-direction: column;
        overflow: hidden;
      }
      #panel.collapsed { width: 46px; }
      #panel.collapsed .hide-collapsed { display: none !important; }
      header {
        min-height: 48px;
        display: flex;
        align-items: center;
        gap: 8px;
        padding: 8px 9px;
        background: #202123;
        color: white;
      }
      header strong { flex: 1; font-size: 14px; }
      .icon-button {
        width: 30px;
        height: 30px;
        border: 0;
        border-radius: 6px;
        background: rgba(255,255,255,.12);
        color: white;
        cursor: pointer;
        font-size: 16px;
      }
      main { padding: 12px; overflow: auto; flex: 1; }
      label { display: block; font-weight: 600; margin: 0 0 5px; }
      select, textarea {
        width: 100%;
        border: 1px solid #b9bec5;
        border-radius: 7px;
        background: white;
        color: #202124;
        padding: 8px;
        font: inherit;
      }
      textarea { min-height: 66px; resize: vertical; }
      .section { margin-bottom: 14px; }
      .mail-card {
        background: white;
        border: 1px solid #d6d9dd;
        border-radius: 8px;
        padding: 9px;
      }
      .mail-title { font-weight: 700; margin-bottom: 4px; overflow-wrap: anywhere; }
      .mail-sender { color: #555; font-size: 12px; margin-bottom: 7px; overflow-wrap: anywhere; }
      .mail-preview {
        white-space: pre-wrap;
        max-height: 180px;
        overflow: auto;
        color: #333;
        font-size: 12px;
        border-top: 1px solid #eceef0;
        padding-top: 7px;
      }
      #response-section[hidden] { display: none; }
      #response-text { min-height: 180px; white-space: pre-wrap; }
      .response-source { margin: 0 0 7px; color: #555; font-size: 11px; overflow-wrap: anywhere; }
      .response-buttons { display: grid; grid-template-columns: 1fr 1fr 1fr; gap: 6px; margin-top: 7px; }
      .response-button {
        border: 1px solid #aeb4bb;
        border-radius: 7px;
        padding: 7px 5px;
        background: white;
        cursor: pointer;
        font: 600 11px system-ui, sans-serif;
      }
      .response-button.primary { border-color: #2f5f96; background: #2f5f96; color: white; }
      .meta { color: #666; font-size: 11px; margin-top: 6px; }
      .buttons { display: grid; grid-template-columns: 1fr 1fr; gap: 7px; }
      .action {
        border: 0;
        border-radius: 7px;
        padding: 9px 7px;
        cursor: pointer;
        background: #2f5f96;
        color: white;
        font: 600 12px system-ui, sans-serif;
      }
      .action.secondary { background: #4f5964; }
      .refresh {
        width: 100%;
        margin-top: 7px;
        border: 1px solid #aeb4bb;
        border-radius: 7px;
        padding: 7px;
        background: white;
        cursor: pointer;
        font: 600 12px system-ui, sans-serif;
      }
      .check {
        display: flex;
        align-items: flex-start;
        gap: 7px;
        font-weight: 400;
        margin: 7px 0;
      }
      .check input { margin-top: 2px; }
      .range-row {
        display: grid;
        grid-template-columns: 1fr auto;
        gap: 8px;
        align-items: center;
        margin: 9px 0 3px;
      }
      .range-row input { width: 100%; }
      .range-value { min-width: 48px; text-align: right; color: #555; font-size: 12px; }
      #status {
        min-height: 36px;
        padding: 9px 10px;
        background: #e8f2eb;
        border-top: 1px solid #ccd8cf;
        color: #25462e;
        font-size: 12px;
      }
      #status.error { background: #f7e9e9; border-color: #e2c2c2; color: #7a2023; }
      .privacy { font-size: 11px; color: #666; }
      .version { opacity: .7; font-weight: 400; font-size: 11px; }
    `;
    shadow.appendChild(style);

    const panel = document.createElement('section');
    panel.id = 'panel';
    if (settings.collapsed) panel.classList.add('collapsed');

    const header = document.createElement('header');
    const toggle = document.createElement('button');
    toggle.id = 'toggle';
    toggle.className = 'icon-button';
    toggle.type = 'button';
    toggle.title = 'Réduire ou ouvrir le volet';
    toggle.textContent = settings.collapsed ? '‹' : '›';

    const heading = document.createElement('strong');
    heading.className = 'hide-collapsed';
    heading.appendChild(document.createTextNode('Courriers BHVP '));
    const version = document.createElement('span');
    version.className = 'version';
    version.textContent = `v${VERSION}`;
    heading.appendChild(version);
    header.append(toggle, heading);

    const main = document.createElement('main');
    main.className = 'hide-collapsed';

    const profileSection = document.createElement('div');
    profileSection.className = 'section';
    const profileLabelElement = document.createElement('label');
    profileLabelElement.htmlFor = 'profile';
    profileLabelElement.textContent = 'Boîte utilisée';
    const profile = document.createElement('select');
    profile.id = 'profile';
    const genericOption = document.createElement('option');
    genericOption.value = 'generic';
    genericOption.textContent = 'Boîte générique BHVP';
    const personalOption = document.createElement('option');
    personalOption.value = 'personal';
    personalOption.textContent = 'Boîte professionnelle personnelle';
    profile.append(genericOption, personalOption);
    profileSection.append(profileLabelElement, profile);

    const mailSection = document.createElement('div');
    mailSection.className = 'section';
    const mailLabel = document.createElement('label');
    mailLabel.textContent = 'Courrier détecté';
    const mailCard = document.createElement('div');
    mailCard.className = 'mail-card';
    const mailTitle = document.createElement('div');
    mailTitle.id = 'mail-title';
    mailTitle.className = 'mail-title';
    mailTitle.textContent = 'Aucun courrier détecté';
    const mailSender = document.createElement('div');
    mailSender.id = 'mail-sender';
    mailSender.className = 'mail-sender';
    const mailPreview = document.createElement('div');
    mailPreview.id = 'mail-preview';
    mailPreview.className = 'mail-preview';
    mailPreview.textContent = 'Ouvre un courrier dans Outlook.';
    const mailMeta = document.createElement('div');
    mailMeta.id = 'mail-meta';
    mailMeta.className = 'meta';
    mailCard.append(mailTitle, mailSender, mailPreview, mailMeta);
    const refresh = document.createElement('button');
    refresh.id = 'refresh';
    refresh.className = 'refresh';
    refresh.type = 'button';
    refresh.textContent = 'Actualiser la détection';
    mailSection.append(mailLabel, mailCard, refresh);

    const extraSection = document.createElement('div');
    extraSection.className = 'section';
    const extraLabel = document.createElement('label');
    extraLabel.htmlFor = 'extra';
    extraLabel.textContent = 'Instruction à transmettre (facultative)';
    const extra = document.createElement('textarea');
    extra.id = 'extra';
    extra.placeholder = 'Ex. : répondre en deux phrases ; observer seulement ma dernière réponse…';
    extraSection.append(extraLabel, extra);

    const buttonSection = document.createElement('div');
    buttonSection.className = 'section buttons';
    const makeActionButton = (text, action, secondary = false) => {
      const button = document.createElement('button');
      button.type = 'button';
      button.className = secondary ? 'action secondary' : 'action';
      button.dataset.action = action;
      button.textContent = text;
      return button;
    };
    buttonSection.append(
      makeActionButton('Préparer une réponse', 'reply'),
      makeActionButton('Envoyer sans consigne', 'raw', true),
      makeActionButton('Résumer', 'summary', true),
      makeActionButton('Observer mes réponses', 'learn', true)
    );

    const responseSection = document.createElement('div');
    responseSection.id = 'response-section';
    responseSection.className = 'section';
    responseSection.hidden = true;
    const responseLabel = document.createElement('label');
    responseLabel.htmlFor = 'response-text';
    responseLabel.textContent = 'Texte prêt à insérer';
    const responseSource = document.createElement('div');
    responseSource.id = 'response-source';
    responseSource.className = 'response-source';
    const responseText = document.createElement('textarea');
    responseText.id = 'response-text';
    responseText.spellcheck = true;
    const responseButtons = document.createElement('div');
    responseButtons.className = 'response-buttons';
    const insertResponse = document.createElement('button');
    insertResponse.type = 'button';
    insertResponse.className = 'response-button primary';
    insertResponse.textContent = 'Insérer';
    insertResponse.title = 'Insérer au début du brouillon Outlook actuellement ouvert';
    const copyResponse = document.createElement('button');
    copyResponse.type = 'button';
    copyResponse.className = 'response-button';
    copyResponse.textContent = 'Copier';
    const clearResponse = document.createElement('button');
    clearResponse.type = 'button';
    clearResponse.className = 'response-button';
    clearResponse.textContent = 'Effacer';
    responseButtons.append(insertResponse, copyResponse, clearResponse);
    responseSection.append(responseLabel, responseSource, responseText, responseButtons);

    const optionsSection = document.createElement('div');
    optionsSection.className = 'section';
    const makeCheck = (id, text) => {
      const label = document.createElement('label');
      label.className = 'check';
      const input = document.createElement('input');
      input.id = id;
      input.type = 'checkbox';
      const span = document.createElement('span');
      span.textContent = text;
      label.append(input, span);
      return { label, input };
    };
    const includeThreadCheck = makeCheck('includeThread', 'Inclure tout le fil visible et déplié');
    const popupWindowCheck = makeCheck('popupWindow', 'Ouvrir ChatGPT dans une fenêtre séparée à droite');
    const autoSendCheck = makeCheck('autoSend', 'Envoyer automatiquement dans ChatGPT');
    const reserveSpaceCheck = makeCheck('reserveSpace', 'Réserver la place dans Outlook (évite de masquer le courrier)');

    const rangeRow = document.createElement('div');
    rangeRow.className = 'range-row';
    const widthRange = document.createElement('input');
    widthRange.id = 'panelWidth';
    widthRange.type = 'range';
    widthRange.min = '300';
    widthRange.max = '520';
    widthRange.step = '10';
    const widthValue = document.createElement('span');
    widthValue.className = 'range-value';
    rangeRow.append(widthRange, widthValue);

    const privacy = document.createElement('div');
    privacy.className = 'privacy';
    privacy.textContent = 'Rien n’est transmis avant le clic sur l’un des boutons. Après génération, la dernière réponse de ChatGPT revient ici automatiquement. Ouvre d’abord un brouillon de réponse Outlook, puis clique sur « Insérer ». Aucun courriel n’est envoyé automatiquement.';
    optionsSection.append(
      includeThreadCheck.label,
      popupWindowCheck.label,
      autoSendCheck.label,
      reserveSpaceCheck.label,
      rangeRow,
      privacy
    );

    main.append(profileSection, mailSection, extraSection, buttonSection, responseSection, optionsSection);

    const status = document.createElement('div');
    status.id = 'status';
    status.className = 'hide-collapsed';
    status.textContent = 'Prêt.';

    panel.append(header, main, status);
    shadow.appendChild(panel);
    document.documentElement.appendChild(host);

    const includeThread = includeThreadCheck.input;
    const popupWindow = popupWindowCheck.input;
    const autoSend = autoSendCheck.input;
    const reserveSpace = reserveSpaceCheck.input;

    profile.value = settings.profile;
    includeThread.checked = settings.includeThread;
    popupWindow.checked = settings.popupWindow;
    autoSend.checked = settings.autoSend;
    reserveSpace.checked = settings.reserveSpace !== false;
    widthRange.value = String(Math.max(300, Math.min(Number(settings.width) || 350, 520)));
    widthValue.textContent = `${widthRange.value} px`;

    const ui = {
      host,
      shadow,
      panel,
      extra,
      readSettings() {
        return {
          profile: profile.value,
          collapsed: panel.classList.contains('collapsed'),
          includeThread: includeThread.checked,
          popupWindow: popupWindow.checked,
          autoSend: autoSend.checked,
          reserveSpace: reserveSpace.checked,
          width: Number(widthRange.value) || 350
        };
      },
      async saveSettings() {
        await gmSet(STORAGE_SETTINGS, this.readSettings());
      },
      renderMail(mail) {
        mailTitle.textContent = mail.subject || '(objet non détecté)';
        mailSender.textContent = mail.sender || '(expéditeur non détecté)';
        const preview = mail.body.length > 2200 ? `${mail.body.slice(0, 2200)}\n[…]` : mail.body;
        mailPreview.textContent = preview;
        mailMeta.textContent = `${mail.body.length.toLocaleString('fr-FR')} caractères détectés`;
      },
      renderResponse(response) {
        if (!response?.text) {
          currentResponse = null;
          responseSection.hidden = true;
          responseSource.textContent = '';
          responseText.value = '';
          return;
        }
        const displayedText = response.action === 'reply'
          ? extractPasteReadyReply(response.text)
          : response.text;
        currentResponse = { ...response, text: displayedText };
        responseSection.hidden = false;
        const sourceBits = [];
        if (response.subject) sourceBits.push(`Objet : ${response.subject}`);
        if (response.action) sourceBits.push(`Action : ${response.action}`);
        responseSource.textContent = sourceBits.join(' · ');
        responseText.value = displayedText;
      },
      setStatus(message, error = false) {
        status.textContent = message;
        status.classList.toggle('error', error);
      }
    };

    toggle.addEventListener('click', async () => {
      const collapsed = panel.classList.toggle('collapsed');
      toggle.textContent = collapsed ? '‹' : '›';
      await ui.saveSettings();
      applyOutlookLayout(ui.readSettings(), panel);
      if (!collapsed) scheduleRefresh(ui, true);
    });

    [profile, includeThread, popupWindow, autoSend, reserveSpace].forEach((element) => {
      element.addEventListener('change', async () => {
        await ui.saveSettings();
        if (element === includeThread) scheduleRefresh(ui, true);
        if (element === reserveSpace) applyOutlookLayout(ui.readSettings(), panel);
      });
    });

    widthRange.addEventListener('input', () => {
      widthValue.textContent = `${widthRange.value} px`;
      applyOutlookLayout(ui.readSettings(), panel);
    });
    widthRange.addEventListener('change', () => ui.saveSettings());

    refresh.addEventListener('click', () => scheduleRefresh(ui, true));
    shadow.querySelectorAll('[data-action]').forEach((button) => {
      button.addEventListener('click', () => sendMail(button.dataset.action, ui));
    });

    insertResponse.addEventListener('click', () => {
      const text = cleanText(responseText.value);
      if (!text) {
        ui.setStatus('Aucun texte à insérer.', true);
        return;
      }
      const detected = extractCurrentMail(false);
      if (currentResponse?.subject && detected?.subject &&
          cleanText(currentResponse.subject) !== cleanText(detected.subject)) {
        const proceed = window.confirm(
          `La réponse a été préparée pour « ${currentResponse.subject} », mais le courrier actuellement détecté est « ${detected.subject} ».\n\nInsérer quand même ?`
        );
        if (!proceed) return;
      }
      const result = insertIntoOutlookDraft(text);
      if (result.ok) ui.setStatus('Texte inséré au début du brouillon Outlook. Relis-le avant d’envoyer.');
      else ui.setStatus(result.message, true);
    });

    copyResponse.addEventListener('click', async () => {
      const text = cleanText(responseText.value);
      if (!text) {
        ui.setStatus('Aucun texte à copier.', true);
        return;
      }
      try {
        if (typeof GM_setClipboard === 'function') GM_setClipboard(text, 'text');
        else await navigator.clipboard.writeText(text);
        ui.setStatus('Texte copié.');
      } catch (error) {
        ui.setStatus(`Copie impossible : ${error.message}`, true);
      }
    });

    clearResponse.addEventListener('click', async () => {
      currentResponse = null;
      ui.renderResponse(null);
      await gmDelete(STORAGE_RESPONSE);
      ui.setStatus('Texte récupéré effacé.');
    });

    applyOutlookLayout(ui.readSettings(), panel);
    return ui;
  }

  async function installOutlookPanel() {
    if (document.getElementById(HOST_ID)) return;

    document.addEventListener('pointermove', (event) => {
      if (!(event.target instanceof Element) || event.target.closest(`#${HOST_ID}`)) return;
      lastPointerTarget = event.target;
    }, true);

    const stored = await gmGet(STORAGE_SETTINGS, DEFAULT_SETTINGS);
    const settings = { ...DEFAULT_SETTINGS, ...(stored || {}) };
    if (stored && typeof stored.includeThread !== 'boolean' && typeof stored.latestOnly === 'boolean') {
      settings.includeThread = !stored.latestOnly;
    }
    const ui = createPanelUi(settings);

    GM_addValueChangeListener(STORAGE_RESPONSE, (_key, _oldValue, newValue, remote) => {
      if (!remote || !newValue?.text) return;
      currentResponse = newValue;
      ui.renderResponse(newValue);
      ui.setStatus('Réponse de ChatGPT récupérée. Ouvre un brouillon Outlook puis clique sur « Insérer ».');
      showPageNotice('Réponse ChatGPT récupérée dans le volet BHVP.');
    });

    const storedResponse = await gmGet(STORAGE_RESPONSE, null);
    if (storedResponse?.text) ui.renderResponse(storedResponse);

    scheduleRefresh(ui, true);

    const observer = new MutationObserver(() => {
      applyOutlookLayout(ui.readSettings(), ui.panel);
      scheduleRefresh(ui, false);
    });
    observer.observe(document.body || document.documentElement, { childList: true, subtree: true });

    window.addEventListener('focus', () => {
      applyOutlookLayout(ui.readSettings(), ui.panel);
      scheduleRefresh(ui, false);
    });
    window.addEventListener('resize', () => applyOutlookLayout(ui.readSettings(), ui.panel));
    window.addEventListener('beforeunload', restoreOutlookLayout);
    document.addEventListener('click', (event) => {
      if (!(event.target instanceof Element) || event.target.closest(`#${HOST_ID}`)) return;
      scheduleRefresh(ui, false);
    }, true);
  }

  function findChatEditor() {
    return document.querySelector(
      '#prompt-textarea[contenteditable="true"], ' +
      'textarea#prompt-textarea, ' +
      '[contenteditable="true"][data-virtualkeyboard="true"], ' +
      'div.ProseMirror[contenteditable="true"]'
    );
  }

  function editorText(editor) {
    if (!editor) return '';
    if (editor instanceof HTMLTextAreaElement) return editor.value;
    return editor.innerText || editor.textContent || '';
  }

  function fillEditor(editor, text) {
    editor.focus();

    if (editor instanceof HTMLTextAreaElement) {
      const setter = Object.getOwnPropertyDescriptor(HTMLTextAreaElement.prototype, 'value')?.set;
      if (setter) setter.call(editor, text);
      else editor.value = text;
      editor.dispatchEvent(new Event('input', { bubbles: true }));
      editor.dispatchEvent(new Event('change', { bubbles: true }));
      return;
    }

    const selection = window.getSelection();
    const range = document.createRange();
    range.selectNodeContents(editor);
    selection.removeAllRanges();
    selection.addRange(range);

    let inserted = false;
    try {
      inserted = document.execCommand('insertText', false, text);
    } catch (_) {
      inserted = false;
    }

    if (!inserted || cleanText(editorText(editor)) !== cleanText(text)) {
      editor.replaceChildren();
      const paragraph = document.createElement('p');
      paragraph.textContent = text;
      editor.appendChild(paragraph);
    }

    editor.dispatchEvent(new InputEvent('input', {
      bubbles: true,
      inputType: 'insertText',
      data: text
    }));
    editor.dispatchEvent(new Event('change', { bubbles: true }));
  }

  function findSendButton() {
    const candidates = [
      '[data-testid="send-button"]',
      'button[aria-label*="Envoyer" i]',
      'button[aria-label*="Send" i]'
    ];
    for (const selector of candidates) {
      const button = document.querySelector(selector);
      if (button instanceof HTMLButtonElement && !button.disabled) return button;
    }
    return null;
  }

  async function insertIntoChatGPT(payload) {
    if (!payload?.prompt) return false;
    if (payload.createdAt && Date.now() - payload.createdAt > 10 * 60 * 1000) {
      await gmDelete(STORAGE_PENDING);
      return false;
    }

    const deadline = Date.now() + 25000;
    let editor = null;
    while (Date.now() < deadline && !editor) {
      editor = findChatEditor();
      if (!editor) await new Promise((resolve) => setTimeout(resolve, 300));
    }
    if (!editor) return false;

    fillEditor(editor, payload.prompt);

    const verifyDeadline = Date.now() + 5000;
    while (Date.now() < verifyDeadline && cleanText(editorText(editor)).length < Math.min(30, payload.prompt.length)) {
      await new Promise((resolve) => setTimeout(resolve, 180));
      fillEditor(editor, payload.prompt);
    }

    await gmDelete(STORAGE_PENDING);
    await gmSet(STORAGE_ACTIVE_REQUEST, payload);
    editor.scrollIntoView({ block: 'center' });
    installChatResponseWatcher(payload);

    if (payload.autoSend) {
      const sendDeadline = Date.now() + 7000;
      let sendButton = null;
      while (Date.now() < sendDeadline && !sendButton) {
        sendButton = findSendButton();
        if (!sendButton) await new Promise((resolve) => setTimeout(resolve, 200));
      }
      sendButton?.click();
    }

    return true;
  }

  function chatMessageText(element) {
    if (!(element instanceof Element)) return '';
    const preferred = element.querySelector('.markdown, [class*="prose"], [data-message-content]');
    return cleanText((preferred || element).innerText || (preferred || element).textContent || '');
  }

  function normalizeForCompare(value) {
    return cleanText(value).replace(/\s+/g, ' ').trim();
  }

  function isReplySalutationLine(line) {
    const value = cleanText(line);
    if (!value || value.length > 140) return false;
    return /^(?:bonjour(?:\s+(?:madame|monsieur|mesdames|messieurs|à toutes et à tous|[^,;:!?]{1,90}))?|madame(?:\s*,\s*monsieur)?|monsieur|chère?\s+[^,;:!?]{1,90}|cher\s+[^,;:!?]{1,90}|dear\s+(?:sir\s+or\s+madam|madam\s+or\s+sir|sir|madam|mr\.?\s+.+|mrs\.?\s+.+|ms\.?\s+.+|miss\s+.+|dr\.?\s+.+|professor\s+.+|.+)|hello(?:\s+.+)?|good\s+(?:morning|afternoon|evening)(?:\s+.+)?)\s*[,!:]?$/i.test(value);
  }

  function isReplyClosingLine(line) {
    const value = cleanText(line);
    if (!value || value.length > 100) return false;
    return /^(?:bien\s+cordialement|très\s+cordialement|cordialement|bien\s+à\s+vous|respectueusement|sincères?\s+salutations|salutations\s+distinguées|avec\s+mes\s+salutations|avec\s+nos\s+salutations|kind\s+regards|best\s+regards|warm\s+regards|regards|yours\s+sincerely|yours\s+faithfully|sincerely|many\s+thanks|thank\s+you)\s*[,!.]?$/i.test(value);
  }

  function extractPasteReadyReply(value) {
    const raw = cleanText(value);
    if (!raw) return '';

    const lines = raw.split('\n').map((line) => line.replace(/[ \t]+$/g, ''));
    let start = -1;
    for (let i = 0; i < lines.length; i += 1) {
      if (isReplySalutationLine(lines[i])) {
        start = i;
        break;
      }
    }

    // Ne rien retrancher si aucune vraie formule d'appel n'est détectée :
    // mieux vaut conserver un texte inhabituel que supprimer une partie utile.
    if (start < 0) return raw;

    let end = -1;
    for (let i = start + 1; i < lines.length; i += 1) {
      if (isReplyClosingLine(lines[i])) end = i;
    }

    // Lorsqu'une formule de clôture est trouvée, tout ce qui suit est considéré
    // comme une signature. Sinon, on retire seulement le préambule avant l'appel.
    const selected = lines.slice(start, end >= start ? end + 1 : lines.length);
    return cleanText(selected.join('\n'));
  }

  function promptLooksLikeMessage(prompt, messageText) {
    const expected = normalizeForCompare(prompt);
    const actual = normalizeForCompare(messageText);
    if (!expected || !actual) return false;
    if (expected === actual) return true;
    const head = expected.slice(0, Math.min(220, expected.length));
    const tail = expected.slice(Math.max(0, expected.length - 180));
    return actual.includes(head) && actual.includes(tail) && actual.length >= expected.length * 0.72;
  }

  function findAssistantResponseForRequest(request) {
    const messages = [...document.querySelectorAll('[data-message-author-role="user"], [data-message-author-role="assistant"]')];
    let matchingUserIndex = -1;
    for (let i = messages.length - 1; i >= 0; i -= 1) {
      if (messages[i].getAttribute('data-message-author-role') !== 'user') continue;
      if (promptLooksLikeMessage(request.prompt, chatMessageText(messages[i]))) {
        matchingUserIndex = i;
        break;
      }
    }
    if (matchingUserIndex < 0) return null;
    for (let i = matchingUserIndex + 1; i < messages.length; i += 1) {
      if (messages[i].getAttribute('data-message-author-role') === 'assistant') return messages[i];
    }
    return null;
  }

  function chatIsGenerating() {
    const selectors = [
      '[data-testid="stop-button"]',
      'button[aria-label*="Arrêter" i]',
      'button[aria-label*="Stop generating" i]',
      'button[aria-label*="Stop" i]'
    ];
    return selectors.some((selector) => {
      const element = document.querySelector(selector);
      return element instanceof HTMLElement && isVisible(element);
    });
  }

  async function publishChatResponse(request, text) {
    const rawText = cleanText(text);
    const readyText = request.action === 'reply'
      ? extractPasteReadyReply(rawText)
      : rawText;
    const response = {
      requestId: request.id,
      text: readyText,
      rawText: readyText !== rawText ? rawText : '',
      trimmedForOutlook: readyText !== rawText,
      capturedAt: Date.now(),
      subject: request.subject || '',
      sender: request.sender || '',
      action: request.action || '',
      profile: request.profile || ''
    };
    if (!response.text) return false;
    await gmSet(STORAGE_RESPONSE, response);
    await gmDelete(STORAGE_ACTIVE_REQUEST);
    return true;
  }

  function installChatResponseWatcher(request) {
    if (!request?.id || !request?.prompt) return;
    let lastText = '';
    let stableSince = 0;
    let finished = false;

    const check = async () => {
      if (finished) return;
      if (request.createdAt && Date.now() - request.createdAt > 60 * 60 * 1000) {
        finished = true;
        return;
      }
      const assistant = findAssistantResponseForRequest(request);
      const text = assistant ? chatMessageText(assistant) : '';
      if (!text) return;
      if (text !== lastText) {
        lastText = text;
        stableSince = Date.now();
        return;
      }
      if (!stableSince) stableSince = Date.now();
      if (!chatIsGenerating() && Date.now() - stableSince >= 1800) {
        finished = await publishChatResponse(request, text);
      }
    };

    const observer = new MutationObserver(() => { check(); });
    observer.observe(document.body || document.documentElement, { childList: true, subtree: true, characterData: true });
    const interval = setInterval(async () => {
      await check();
      if (finished) {
        clearInterval(interval);
        observer.disconnect();
      }
    }, 600);

    installManualReturnButton(request, async () => {
      const assistant = findAssistantResponseForRequest(request) ||
        [...document.querySelectorAll('[data-message-author-role="assistant"]')].at(-1);
      const text = assistant ? chatMessageText(assistant) : '';
      if (!text) return false;
      finished = await publishChatResponse(request, text);
      return finished;
    });
  }

  function installManualReturnButton(request, onClick) {
    const id = 'bhvp-return-to-outlook';
    document.getElementById(id)?.remove();
    const button = document.createElement('button');
    button.id = id;
    button.type = 'button';
    button.textContent = '↩ Envoyer vers Outlook';
    Object.assign(button.style, {
      position: 'fixed',
      right: '18px',
      bottom: '82px',
      zIndex: '2147483647',
      border: '1px solid #8b8f94',
      borderRadius: '8px',
      padding: '8px 11px',
      background: '#202123',
      color: '#fff',
      font: '600 12px system-ui, sans-serif',
      cursor: 'pointer',
      boxShadow: '0 4px 14px rgba(0,0,0,.2)'
    });
    button.title = 'Solution de secours si la récupération automatique ne se déclenche pas';
    button.addEventListener('click', async () => {
      const ok = await onClick();
      button.textContent = ok ? '✓ Envoyé vers Outlook' : 'Aucune réponse détectée';
      setTimeout(() => { if (button.isConnected) button.textContent = '↩ Envoyer vers Outlook'; }, 2400);
    });
    document.documentElement.appendChild(button);
  }

  function installChatGPTReceiver() {
    GM_addValueChangeListener(STORAGE_PENDING, (_key, _oldValue, newValue, remote) => {
      if (remote && newValue) insertIntoChatGPT(newValue);
    });

    Promise.all([
      gmGet(STORAGE_PENDING, null),
      gmGet(STORAGE_ACTIVE_REQUEST, null)
    ]).then(([payload, request]) => {
      if (payload) insertIntoChatGPT(payload);
      else if (request?.id && request?.prompt) installChatResponseWatcher(request);
    });
  }

  if (/^outlook\./i.test(location.hostname)) {
    installOutlookPanel().catch((error) => {
      console.error('[BHVP panel]', error);
      showPageNotice(`Le volet BHVP n’a pas pu démarrer : ${error.message}`, true);
    });
  } else if (location.hostname === 'chatgpt.com') {
    installChatGPTReceiver();
  }
})();
