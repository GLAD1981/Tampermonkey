// ==UserScript==
// @name         BHVP – Volet Outlook vers ChatGPT
// @namespace    bhvp-outlook-chatgpt
// @version      1.3.1
// @description  Envoie le courrier visible vers ChatGPT, récupère automatiquement sa réponse et peut l’insérer dans un brouillon Outlook.
// @homepageURL  https://github.com/GLAD1981/Tampermonkey
// @updateURL    https://raw.githubusercontent.com/GLAD1981/Tampermonkey/main/BHVP-Volet-Outlook.user.js
// @downloadURL  https://raw.githubusercontent.com/GLAD1981/Tampermonkey/main/BHVP-Volet-Outlook.user.js
// @require      https://raw.githubusercontent.com/GLAD1981/Tampermonkey/fc5d720564a7f2250fd88518f17b330cd3f2acd4/BHVP-Volet-Outlook.user.js
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

  const HOST_ID = 'bhvp-outlook-panel-host';
  const RETURN_BUTTON_ID = 'bhvp-return-to-outlook';
  const PATCH_MARKER = 'bhvp-simple-ui-v131';

  function textOf(element) {
    return String(element?.textContent || '').replace(/\s+/g, ' ').trim();
  }

  function removeManualReturnButton() {
    document.getElementById(RETURN_BUTTON_ID)?.remove();
  }

  function patchOutlookPanel() {
    const host = document.getElementById(HOST_ID);
    const shadow = host?.shadowRoot;
    if (!shadow || shadow.getElementById(PATCH_MARKER)) return false;

    const marker = document.createElement('span');
    marker.id = PATCH_MARKER;
    marker.hidden = true;
    shadow.appendChild(marker);

    const panel = shadow.getElementById('panel');
    const main = panel?.querySelector('main');
    if (!panel || !main) return false;

    const version = shadow.querySelector('.version');
    if (version) version.textContent = 'v1.3.1';

    // Supprime le pavé « Boîte utilisée », tout en conservant sa valeur
    // en mémoire dans le code historique.
    const profile = shadow.getElementById('profile');
    profile?.closest('.section')?.remove();

    // Ne conserve que le compteur de caractères du courrier détecté.
    const mailMeta = shadow.getElementById('mail-meta');
    const mailSection = mailMeta?.closest('.section');
    if (mailMeta && mailSection) {
      mailSection.replaceChildren(mailMeta);
      mailMeta.className = 'detected-count';
      Object.assign(mailMeta.style, {
        margin: '0',
        padding: '5px 0 2px',
        color: '#555',
        fontSize: '12px'
      });
    }
    shadow.getElementById('refresh')?.remove();

    // Un seul bouton de préparation : sans instruction, il déclenche
    // l'ancien mode « Envoyer sans consigne » ; avec instruction, il
    // déclenche la préparation d'une réponse.
    const replyOriginal = shadow.querySelector('button[data-action="reply"]');
    const rawOriginal = shadow.querySelector('button[data-action="raw"]');
    const summaryOriginal = shadow.querySelector('button[data-action="summary"]');
    const learnButton = shadow.querySelector('button[data-action="learn"]');
    const extra = shadow.getElementById('extra');
    const buttons = replyOriginal?.parentElement;

    if (replyOriginal && rawOriginal && buttons) {
      const replyButton = replyOriginal.cloneNode(true);
      replyButton.removeAttribute('data-action');
      replyButton.textContent = 'Préparer une réponse';
      replyButton.addEventListener('click', () => {
        const hasInstruction = Boolean(String(extra?.value || '').trim());
        (hasInstruction ? replyOriginal : rawOriginal).click();
      });

      replyOriginal.hidden = true;
      rawOriginal.hidden = true;
      summaryOriginal?.remove();
      buttons.prepend(replyButton);
      buttons.style.gridTemplateColumns = learnButton ? '1fr 1fr' : '1fr';
    } else {
      summaryOriginal?.remove();
      rawOriginal?.remove();
    }

    // Inclut toujours tout le fil visible, sans afficher l'option.
    const includeThread = shadow.getElementById('includeThread');
    if (includeThread) {
      if (!includeThread.checked) {
        includeThread.checked = true;
        includeThread.dispatchEvent(new Event('change', { bubbles: true }));
      }
      includeThread.closest('label')?.remove();
    }

    // Ouvre toujours ChatGPT dans une fenêtre séparée, sans afficher l'option.
    const popupWindow = shadow.getElementById('popupWindow');
    if (popupWindow) {
      if (!popupWindow.checked) {
        popupWindow.checked = true;
        popupWindow.dispatchEvent(new Event('change', { bubbles: true }));
      }
      popupWindow.closest('label')?.remove();
    }

    // Fixe la largeur à 350 px et supprime le curseur de réglage.
    const panelWidth = shadow.getElementById('panelWidth');
    if (panelWidth) {
      panelWidth.value = '350';
      panelWidth.dispatchEvent(new Event('input', { bubbles: true }));
      panelWidth.dispatchEvent(new Event('change', { bubbles: true }));
      panelWidth.closest('.range-row')?.remove();
    }

    // Réserve toujours la place dans Outlook, sans afficher l'option.
    const reserveSpace = shadow.getElementById('reserveSpace');
    if (reserveSpace) {
      if (!reserveSpace.checked) {
        reserveSpace.checked = true;
        reserveSpace.dispatchEvent(new Event('change', { bubbles: true }));
      }
      reserveSpace.closest('label')?.remove();
    }

    // Retire l'avertissement d'information demandé.
    shadow.querySelector('.privacy')?.remove();

    return true;
  }

  if (/^outlook\./i.test(location.hostname)) {
    const observer = new MutationObserver(() => patchOutlookPanel());
    observer.observe(document.documentElement, { childList: true, subtree: true });
    patchOutlookPanel();
  } else if (location.hostname === 'chatgpt.com') {
    const observer = new MutationObserver(removeManualReturnButton);
    observer.observe(document.documentElement, { childList: true, subtree: true });
    removeManualReturnButton();
  }
})();
