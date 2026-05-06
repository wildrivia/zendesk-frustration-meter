// ==UserScript==
// @name         Zendesk Frustration Meter
// @namespace    https://github.com/wildrivia/zendesk-frustration-meter
// @version      0.9.4
// @description  Analyzes customer frustration levels in Zendesk tickets using rule-based scoring. Shows progression timeline, categories, and matched phrases.
// @author       OJ
// @match        https://*.zendesk.com/agent/tickets/*
// @grant        none
// @run-at       document-idle
// @downloadURL  https://raw.githubusercontent.com/wildrivia/zendesk-frustration-meter/main/zendesk-frustration-meter.user.js
// @updateURL    https://raw.githubusercontent.com/wildrivia/zendesk-frustration-meter/main/zendesk-frustration-meter.user.js
// ==/UserScript==

(function () {
  'use strict';

  // ─── Constants ────────────────────────────────────────────────────────────

  const PANEL_ID = 'fm-panel-v05';
  const STYLE_ID = 'fm-styles-v05';
  const POSITION_KEY = 'fm-panel-position-v05';
  const INITIAL_DELAY = 1500;
  const NAV_DELAY = 2000;
  const LOAD_DELAY = 2000;

  // Support tier SLAs — determines how wait-time boosts are scored.
  // Each entry lists all tag variants Zendesk may use for that tier.
  const SLA_TIERS = [
    { tags: ['target_customer', 'support_target_customer'], label: 'Target Customer', hours: 4  },
    { tags: ['tier_1', 'support_tier_1'],                   label: 'Tier 1',          hours: 6  },
    { tags: ['tier_2', 'support_tier_2'],                   label: 'Tier 2',          hours: 12 },
    { tags: ['tier_3', 'support_tier_3'],                   label: 'Tier 3',          hours: 24 },
  ];

  // Tickets from these requester email domains are Stripe/partner bridge tickets — scored N/A.
  const BYPASS_REQUESTER_DOMAINS = ['stripe.com'];
  // Tag-based bypass — tickets with any of these tags are Stripe bridge tickets, scored N/A.
  const BYPASS_TAGS = ['stripe_general_inquiry', 'stripe_notifications', 'wcpay_stripe_notifications'];

  // ─── Frustration Model ────────────────────────────────────────────────────

  const CATEGORIES = {
    repetition: {
      weight: 3,
      label: 'Repeated Effort',
      color: '#8b5cf6',
      phrases: [
        'still not working', 'already tried', 'again', 'as i said',
        'for the second time', 'for the third time', 'multiple times',
        'once again', 'i already explained', 'i already sent',
        'tried that already', 'still having the same issue', 'keep having',
        'repeatedly', 'every time i', 'nothing has changed', 'same problem',
        'still the same', 'still happening', 'still broken', 'still getting',
        'same issue', 'same error', 'back again', 'happened again',
        'issue persists', 'problem persists', 'continues to',
      ],
    },
    delay: {
      weight: 2,
      label: 'Delay / Waiting',
      color: '#f59e0b',
      phrases: [
        'still waiting', 'taking too long', 'waiting for days', 'no response',
        'urgent', 'days later', "haven't heard back", 'this is taking forever',
        'waiting so long', 'it has been days', "i've been waiting", 'still no update',
      ],
    },
    blocked: {
      weight: 2,
      label: 'Blocked / Not Working',
      color: '#3b82f6',
      phrases: [
        'not working', "can't", 'cannot', 'unable to', "doesn't work", 'broken',
        'failed', 'error', 'locked out', 'stuck', "won't let me", 'impossible',
        'not loading', "doesn't load", "can't access", 'keeps crashing', "won't work",
        'not possible to', 'not able to', 'stopped working', 'no longer works',
        'no longer working', "doesn't let me", "won't add", "won't remove",
        "won't open", 'not functioning', 'stopped functioning',
        'showing empty', 'showing as empty', 'not showing', "isn't working",
        "doesn't work anymore", 'not working anymore',
        'keeps failing', 'keeps giving', 'not going through',
        'plugin conflict', 'css conflict', 'jquery conflict', 'javascript conflict',
        'compatibility issue', 'not compatible', 'incompatible',
        'interfering', 'breaking my', 'broke my', 'breaks my', 'messing up', 'messed up',
      ],
    },
    support_failure: {
      weight: 4,
      label: 'Support Failure',
      color: '#ef4444',
      phrases: [
        "this didn't help", 'not understanding', 'same answer', 'nobody is helping',
        'already explained this', 'not addressing my issue', 'asking the same questions',
        'no one has resolved', 'nobody has resolved', 'told the same thing',
        'not helpful', "that didn't work", "doesn't help",
      ],
    },
    emotion: {
      weight: 3,
      label: 'Direct Frustration',
      color: '#f97316',
      phrases: [
        'frustrated', 'frustrating', 'annoyed', 'upset', 'angry', 'disappointed',
        'ridiculous', 'unacceptable', 'fed up', 'irritated', 'exhausting',
        'tired of this', 'absurd', 'terrible', 'horrible', 'awful',
        'this is a joke', 'pathetic', 'outrageous',
      ],
    },
    escalation_request: {
      weight: 2,
      label: 'Refund / Dispute / Cancel',
      color: '#f59e0b',
      phrases: [
        'refund', 'money back', 'want a refund', 'want to cancel',
        'cancellation', 'cancel my subscription', 'cancel my order',
        'cancel the subscription', 'cancel the order', 'cancel it',
        'request a refund', 'requesting a refund', 'asking for a refund',
        'asking for a cancellation', 'process a refund', 'issue a refund',
        'dispute', 'disputed', 'disputing', 'payment dispute', 'chargeback dispute',
        'open a dispute', 'opened a dispute', 'file a dispute', 'raise a dispute',
        'filing a dispute', 'i disputed',
      ],
    },
    escalation: {
      weight: 5,
      label: 'Escalation',
      color: '#d92d20',
      phrases: [
        'cancel my account', 'close my account', 'escalate', 'manager', 'supervisor',
        'complaint', 'report this', 'chargeback', 'switching', 'legal',
        'social media', 'go elsewhere', 'switch provider', 'leaving',
      ],
    },
  };

  const CAPS_ACRONYM_EXCLUSIONS = new Set([
    // Technical / product acronyms
    'CSAT', 'HTTP', 'HTML', 'API', 'URL', 'FAQ', 'USPS', 'UPS', 'SMTP',
    'FTP', 'DNS', 'SSL', 'CSS', 'SQL', 'PHP', 'SKU', 'CSV',
    // Common short caps words that are not intensity signals
    'YOUR', 'THIS', 'WHAT', 'HAVE', 'THAT', 'WITH',
    'FROM', 'WILL', 'HELP', 'JUST', 'NEED', 'CANT',
    'DONT', 'WONT',
  ]);

  const AGENT_PHRASES = [
    'thanks for reaching out', 'thank you for contacting', 'let me help',
    "i'd be happy to help", "i'll look into", 'please try the following',
    'could you please', 'can you please', 'please share', 'best regards',
    'kind regards', 'happiness engineer', 'our team', "we'll get back to you",
    'please note that', 'i can see that', "i've checked",
  ];

  // Phrases that unambiguously identify an agent reply (never written by customers)
  const AGENT_PHRASES_DEFINITIVE = [
    // Opening / greeting
    'thanks for reaching out',
    'thank you for reaching out',
    'thanks for contacting',
    'thank you for contacting',
    'thank you for getting in touch',
    "i'd be happy to help",
    "i'd be happy to assist",
    "i'll look into this",
    "i'll take a look",
    "let me look into",
    "let me check on",
    "let me take a look",
    // Agent identity / role
    'happiness engineer',
    'happiness team',
    'our support team',
    'our team will',
    '| automattic',
    '@ automattic',
    'automattic, inc',
    'customer success manager',
    'growth engineer',
    // Zendesk promotional footers — only appear on HE outbound emails, never on customer replies
    'check out pressable',
    'with woo, you',
    'high-volume stores',
    'view boost oxygen',
    'google for woocommerce',
    // Verification / investigation language
    "i've checked your",
    "i've verified",
    "i've confirmed",
    "i've taken a look",
    "i've had a look",
    // Instructions
    'please try the following',
    'please follow these steps',
    'could you please try',
    'could you please share',
    'could you please send',
    // Closing / sign-off — NOTE: generic closings like "kind regards", "best regards",
    // "many thanks", "hope this helps" are intentionally excluded here because business
    // customers use them constantly. Only keep phrases that are unambiguously agent-only.
    // Phrases like "looking forward to hearing from you", "let me know if you have any
    // other questions", "if you have any other questions", and "i can confirm that" are
    // also excluded — professional business customers routinely write these.
    "we'll get back to you",
    "don't hesitate to reach out",
    "please don't hesitate",
    "feel free to reach out",
    // CSAT / feedback follow-up (HE closing on resolved tickets)
    'thanks for sharing your feedback',
    'thank you for sharing your feedback',
    'thanks for taking the time to share',
    'thank you for taking the time to share',
    // Internal quality review notes (frustration trigger analysis, CSAT review templates)
    'csat review',
    'frustration trigger',
  ];

  const LEVEL_THRESHOLDS = [
    { max: 3,  label: 'Low',      color: '#2da44e' },
    { max: 7,  label: 'Moderate', color: '#f5c542' },
    { max: 12, label: 'High',     color: '#f79009' },
    { max: Infinity, label: 'Severe', color: '#d92d20' },
  ];

  // ─── Utility helpers ──────────────────────────────────────────────────────

  function escapeRegex(str) {
    return str.replace(/[.*+?^${}()|[\]\\]/g, '\\$&');
  }

  function escapeHtml(str) {
    return String(str)
      .replace(/&/g, '&amp;')
      .replace(/</g, '&lt;')
      .replace(/>/g, '&gt;')
      .replace(/"/g, '&quot;')
      .replace(/'/g, '&#39;');
  }

  function getLevel(score) {
    return LEVEL_THRESHOLDS.find(t => score <= t.max);
  }

  function cleanText(raw) {
    if (!raw) return '';
    let text = raw;
    // Remove quoted reply blocks
    text = text.replace(/On .+?wrote:[\s\S]*/i, '');
    // Strip email signature sign-off and everything after it.
    // Business customers routinely close with "Kind Regards," followed by their name,
    // company, contact details, and legal disclaimers \u2014 none of which should be analysed.
    // Must run before whitespace collapse so the newline anchor works.
    text = text.replace(/\n[ \t]*(kind regards|best regards|warm regards|many thanks|sincerely|regards)[,.]?[ \t]*([\r\n][\s\S]*)?$/i, '');
    // Strip "---- Original Message ----" / "---- Forwarded Message ----" separators
    // and everything after — common in Outlook and webmail quoted replies.
    text = text.replace(/\n[ \t]*[-─]{3,}[^\n\r]*(original|forwarded|reply)[^\n\r]*[-─]{3,}[\s\S]*/i, '');
    // Strip Outlook-style quoted headers (From: / Sent: / To: / Subject:) and everything after.
    text = text.replace(/\n[ \t]*From:[ \t]+\S[\s\S]*/i, '');
    // Strip "> " quoted lines (RFC 2822 standard quote style used by many email clients).
    text = text.replace(/^>.*$/gm, '');
    // Remove remaining common email footer lines
    text = text.replace(/^From:.*$/gm, '');
    text = text.replace(/^Sent from my.*$/gim, '');
    text = text.replace(/^Get Outlook for.*$/gim, '');
    // Remove horizontal rules
    text = text.replace(/^-{3,}.*$/gm, '');
    // Normalize curly/smart apostrophes and quotes to straight versions
    // Fixes matching failures when customers type in Word, iOS, etc.
    text = text.replace(/[\u2018\u2019\u201A\u201B\u02BC\u02BB]/g, "'");
    text = text.replace(/[\u201C\u201D\u201E\u201F]/g, '"');
    // Collapse whitespace
    text = text.replace(/\s+/g, ' ').trim();
    return text;
  }

  function isLikelyAgentReply(text) {
    const lower = text.toLowerCase();
    let matches = 0;
    for (const phrase of AGENT_PHRASES) {
      if (lower.includes(phrase)) {
        matches++;
        if (matches >= 2) return true;
      }
    }
    return false;
  }

  // Stricter single-phrase check for contexts where we're fairly confident
  // it's inbound (e.g. WebInteraction) but want to exclude clear HE replies.
  function isDefinitelyAgentReply(text) {
    const lower = text.toLowerCase();
    for (const phrase of AGENT_PHRASES_DEFINITIVE) {
      if (!lower.includes(phrase)) continue;
      // Customers reference their support contact by role: "my Happiness Engineer",
      // "their HE replied". Not an agent self-identification signal in those cases.
      if ((phrase === 'happiness engineer' || phrase === 'happiness team') &&
          /\b(my|your|their|our|the)\s+happiness/i.test(lower)) continue;
      return true;
    }
    return false;
  }

  // ─── Frustration Analyzer ─────────────────────────────────────────────────

  function analyzeMessage(text) {
    const lower = text.toLowerCase();
    const matchedCategories = {};
    const allPhrases = [];

    for (const [catKey, cat] of Object.entries(CATEGORIES)) {
      const matched = [];
      for (const phrase of cat.phrases) {
        const re = new RegExp(`\\b${escapeRegex(phrase)}\\b`, 'i');
        if (re.test(lower)) {
          matched.push(phrase);
        }
      }
      if (matched.length > 0) {
        matchedCategories[catKey] = matched;
        allPhrases.push(...matched.map(p => ({ phrase: p, category: catKey })));
      }
    }

    // Base score from categories
    let score = 0;
    for (const [catKey, phrases] of Object.entries(matchedCategories)) {
      score += CATEGORIES[catKey].weight;
    }

    // Intensity signals — operate on original case
    let capsBonus = 0;
    const capsWords = text.match(/\b[A-Z]{4,}\b/g) || [];
    for (const word of capsWords) {
      if (!CAPS_ACRONYM_EXCLUSIONS.has(word)) {
        capsBonus++;
        if (capsBonus >= 2) break;
      }
    }
    score += capsBonus;

    let exclBonus = 0;
    const exclMatches = text.match(/!!+/g) || [];
    exclBonus = Math.min(exclMatches.length, 2);
    score += exclBonus;

    let questBonus = 0;
    const questMatches = text.match(/\?\?+/g) || [];
    questBonus = Math.min(questMatches.length, 2);
    score += questBonus;

    // Combo boosts
    const boosts = [];
    const cats = Object.keys(matchedCategories);

    if (matchedCategories.repetition && matchedCategories.blocked) {
      score += 2;
      boosts.push({ reason: 'Repeated effort on a blocked issue', value: 2 });
    }
    if (matchedCategories.delay && matchedCategories.support_failure) {
      score += 2;
      boosts.push({ reason: 'Waiting + feeling unsupported', value: 2 });
    }
    if (matchedCategories.emotion && matchedCategories.support_failure) {
      score += 2;
      boosts.push({ reason: 'Explicit frustration + support friction', value: 2 });
    }
    if (matchedCategories.escalation) {
      score += 3;
      boosts.push({ reason: 'Escalation language detected', value: 3 });
    }
    if (cats.length >= 3) {
      score += 2;
      boosts.push({ reason: 'Multiple trigger categories matched', value: 2 });
    }

    const level = getLevel(score);

    return {
      score,
      level: level.label,
      color: level.color,
      categories: matchedCategories,
      phrases: allPhrases,
      boosts,
      intensity: { capsBonus, exclBonus, questBonus },
    };
  }

  function computeTrend(messages, thread, isSolved) {
    if (messages.length < 2) return { label: 'New ticket', color: '#6b7280' };

    // Count consecutive customer messages at the tail of the thread with no HE reply.
    // 2+ unanswered follow-ups means the customer is escalating regardless of word choice.
    // Skipped on solved/closed tickets — trailing messages are expected closing exchanges.
    let trailingUnanswered = 0;
    for (let i = thread.length - 1; i >= 0; i--) {
      if (thread[i].type === 'customer') trailingUnanswered++;
      else break;
    }
    if (trailingUnanswered >= 2 && !isSolved) return { label: 'Increasing', color: '#f79009' };

    const mid = Math.floor(messages.length / 2);
    const firstHalf = messages.slice(0, mid);
    const secondHalf = messages.slice(mid);
    const avg = arr => arr.reduce((s, m) => s + m.analysis.score, 0) / arr.length;
    const gap = avg(secondHalf) - avg(firstHalf);
    if (gap > 2) return { label: 'Increasing', color: '#f79009' };
    if (gap < -2) return { label: 'Decreasing', color: '#2da44e' };
    return { label: 'Stable', color: '#6b7280' };
  }

  // Extract a Date from a Zendesk comment article element.
  function getTimestamp(article) {
    // Most reliable: <time datetime="..."> element (email/web messages)
    const timeEl = article.querySelector('time[datetime]');
    if (timeEl) {
      const dt = new Date(timeEl.getAttribute('datetime'));
      if (!isNaN(dt.getTime())) return dt;
    }
    // Fallback: any element whose title looks like a datetime
    const candidates = article.querySelectorAll('[title]');
    for (const el of candidates) {
      const title = el.getAttribute('title') || '';
      if (/\d{4}-\d{2}-\d{2}/.test(title) || /\w+ \d+,? \d{4}/.test(title)) {
        const dt = new Date(title);
        if (!isNaN(dt.getTime())) return dt;
      }
    }
    // Chat fallback: parse the visible timestamp-relative text e.g. "Feb 26 23:08"
    const relEl = article.querySelector('[data-test-id="timestamp-relative"]');
    if (relEl) {
      const text = (relEl.textContent || '').trim();
      const m = text.match(/^(\w+)\s+(\d+)\s+(\d+):(\d+)$/);
      if (m) {
        const now = new Date();
        const year = now.getFullYear();
        const candidate = new Date(`${m[1]} ${m[2]}, ${year} ${m[3]}:${m[4]}`);
        if (!isNaN(candidate.getTime())) {
          // If parsing gives a future date, it must be from last year
          return candidate > now
            ? new Date(`${m[1]} ${m[2]}, ${year - 1} ${m[3]}:${m[4]}`)
            : candidate;
        }
      }
    }
    return null;
  }

  // Format a timestamp for display in the progression timeline.
  function formatTimestamp(ts) {
    if (!ts) return '';
    const now = new Date();
    const diffDays = Math.floor((now - ts) / 86400000);
    const hh = String(ts.getHours()).padStart(2, '0');
    const mm = String(ts.getMinutes()).padStart(2, '0');
    const time = `${hh}:${mm}`;
    if (diffDays === 0) return time;
    const day = ts.toLocaleDateString('en-GB', { day: 'numeric', month: 'short' });
    return `${day} ${time}`;
  }

  // Compute timing-based frustration boosts from the message thread.
  function computeTimingBoosts(messages) {
    const boosts = [];
    const withTime = messages.filter(m => m.timestamp instanceof Date);
    if (withTime.length < 2) return boosts;

    // Measure time between each pair of consecutive messages
    for (let i = 1; i < withTime.length; i++) {
      const gapMs  = withTime[i].timestamp - withTime[i - 1].timestamp;
      const gapMin = Math.round(gapMs / 60000);
      const gapHrs = gapMs / 3600000;

      if (gapMin < 60) {
        boosts.push({ reason: `Follow-up sent ${gapMin} min after previous message`, value: 2 });
      } else if (gapHrs < 6) {
        boosts.push({ reason: `Follow-up sent ${Math.round(gapHrs * 10) / 10} hrs after previous message`, value: 1 });
      } else if (gapHrs < 24) {
        boosts.push({ reason: `Follow-up sent same day (${Math.round(gapHrs)} hrs later)`, value: 1 });
      }
      // Gaps of days or more: issue persistence — not scored here to avoid double-counting
    }

    return boosts;
  }

  // Format a wait duration in hours into a readable label e.g. "1d 6h" or "3h 20m".
  function formatWaitDuration(hrs) {
    if (hrs >= 24) {
      const d = Math.floor(hrs / 24);
      const h = Math.round(hrs % 24);
      return h > 0 ? `${d}d ${h}h` : `${d}d`;
    }
    const h = Math.floor(hrs);
    const m = Math.round((hrs - h) * 60);
    return m > 0 ? `${h}h ${m}m` : `${h}h`;
  }

  function getTicketStatus() {
    const STATUSES = ['solved', 'closed', 'open', 'pending', 'on-hold'];
    const spans = document.querySelectorAll('span');
    for (const span of spans) {
      if (span.children.length > 0) continue;
      const t = span.textContent.trim().toLowerCase();
      if (STATUSES.includes(t)) return t;
    }
    return null;
  }

  // Read the ticket's support tier from its tags and return the matching SLA entry.
  function getTicketSla() {
    const tags = new Set();
    document.querySelectorAll('[data-test-id="ticket-system-field-tags-item-selected"]').forEach(el => {
      const t = (el.textContent || '').trim().toLowerCase();
      if (t) tags.add(t);
    });
    for (const tier of SLA_TIERS) {
      if (tier.tags.some(tag => tags.has(tag))) return tier;
    }
    return null;
  }

  // Compute wait-time boosts: how long the customer had to wait for an HE to respond.
  // Scored relative to the ticket's SLA tier when a tier tag is found; falls back to
  // absolute thresholds when no tier tag is present.
  function applyWaitBoost(boosts, waitHrs, sla, label) {
    const waitLabel = formatWaitDuration(waitHrs);
    if (sla) {
      const ratio = waitHrs / sla.hours;
      const slaCtx = `${sla.label} SLA: ${sla.hours}h`;
      if (ratio > 2) {
        boosts.push({ reason: `${label} ${waitLabel} — ${slaCtx} (severely breached)`, value: 3 });
      } else if (ratio > 1) {
        boosts.push({ reason: `${label} ${waitLabel} — ${slaCtx} (breached)`, value: 2 });
      } else if (ratio > 0.75) {
        boosts.push({ reason: `${label} ${waitLabel} — ${slaCtx} (approaching limit)`, value: 1 });
      }
    } else {
      if (waitHrs >= 24) {
        boosts.push({ reason: `${label} ${waitLabel} for HE response`, value: 3 });
      } else if (waitHrs >= 6) {
        boosts.push({ reason: `${label} ${waitLabel} for HE response`, value: 2 });
      } else if (waitHrs >= 1) {
        boosts.push({ reason: `${label} ${waitLabel} for HE response`, value: 1 });
      }
    }
  }

  function computeWaitTimeBoosts(thread, isSolved) {
    const boosts = [];
    const sla = getTicketSla();
    let lastCustomerTs = null;
    let firstUnansweredTs = null; // start of the current unanswered customer streak

    for (const event of thread) {
      if (event.type === 'customer') {
        lastCustomerTs = event.timestamp instanceof Date ? event.timestamp : null;
        // Mark the start of an unanswered streak on the first message in it
        if (!firstUnansweredTs && lastCustomerTs) firstUnansweredTs = lastCustomerTs;
      } else if (event.type === 'agent' && lastCustomerTs) {
        if (!(event.timestamp instanceof Date)) continue;
        const waitMs = event.timestamp - lastCustomerTs;
        if (waitMs <= 0) continue;
        applyWaitBoost(boosts, waitMs / 3600000, sla, 'Waited');
        lastCustomerTs = null;
        firstUnansweredTs = null; // agent replied — streak is over
      }
    }

    // Customer's last message(s) have no HE reply yet — score the open wait.
    // Skipped on solved/closed tickets: the final customer message is often a social
    // closing that doesn't warrant a response, so the open wait is not meaningful.
    if (firstUnansweredTs && !isSolved) {
      const waitMs = Date.now() - firstUnansweredTs.getTime();
      if (waitMs > 0) applyWaitBoost(boosts, waitMs / 3600000, sla, 'No HE reply — waiting');
    }

    return boosts;
  }

  // Identify likely frustration themes across the full conversation.
  // Based on OJ's P2 research: "Mapping Frustration Triggers in Negative CSATs".
  // Identifies which of the 7 frustration triggers from the Team Libra learnup apply.
  // Each trigger requires a specific primary signal — msgCount alone never fires a theme.
  function detectFrustrationThemes(messages, thread, linearLinks = []) {
    if (!messages.length) return { themes: [], nextSteps: [] };

    const allCategories = {};
    let allText = '';
    for (const msg of messages) {
      allText += ' ' + msg.text.toLowerCase();
      for (const cat of Object.keys(msg.analysis.categories)) {
        allCategories[cat] = true;
      }
    }

    const msgCount  = messages.length;
    const allBoosts = messages.flatMap(m => m.analysis.boosts);
    const hasAny    = words => words.some(w => allText.includes(w));

    // Recent window — last 3 customer messages only.
    // Phrase-based triggers use this so that old frustrated messages on a now-calm
    // ticket don't bleed into the context. Timing-based signals still use full history.
    const recentMsgs = messages.slice(-3);
    const recentCategories = {};
    let recentText = '';
    for (const msg of recentMsgs) {
      recentText += ' ' + msg.text.toLowerCase();
      for (const cat of Object.keys(msg.analysis.categories)) {
        recentCategories[cat] = true;
      }
    }
    const hasAnyRecent = words => words.some(w => recentText.includes(w));

    // HE reply text — used for trigger 06 so we can detect bug/workaround confirmations
    // from the agent side even when customer messages don't say "known issue".
    let heText = '';
    for (const event of thread) {
      if (event.type === 'agent' && event.text) heText += ' ' + event.text.toLowerCase();
    }
    const heConfirms = words => words.some(w => heText.includes(w));

    // Wait-time signals derived from score boosts
    const hasOpenWait   = allBoosts.some(b => b.reason && b.reason.startsWith('No HE reply'));
    // Significant wait = any wait boost with value ≥ 2, covering both SLA-relative and absolute formats
    const hasWaitBreach = allBoosts.some(b => b.value >= 2 &&
      (b.reason.startsWith('Waited') || b.reason.startsWith('No HE reply')));
    const hasLongWait   = allBoosts.some(b => b.reason && /\b[2-9]\d*d\b/.test(b.reason));

    // Count consecutive unanswered customer messages at the tail of the thread
    let trailingUnanswered = 0;
    for (let i = thread.length - 1; i >= 0; i--) {
      if (thread[i].type === 'customer') trailingUnanswered++;
      else break;
    }

    // True only when an agent event follows at least one customer event.
    // Automated pre-chat messages (Woo, queue notice) appear before customers
    // and are typed 'agent', but should not count as a human reply.
    let seenCustomer = false;
    let hadHEReply   = false;
    for (const event of thread) {
      if (event.type === 'customer') seenCustomer = true;
      else if (event.type === 'agent' && seenCustomer) { hadHEReply = true; break; }
    }

    // Detect whether the customer has been actively completing requests made of them.
    // Used by trigger 03 to distinguish a cooperative customer waiting for an outcome
    // from one who is passively frustrated about slow replies.
    const cooperativeCustomer = hasAny([
      'i have provided', "i've provided", 'i already sent', 'i already provided',
      'i submitted', 'i uploaded', 'i shared the', 'i responded to',
      'already provided', 'already sent', 'already submitted',
      'should be good to go', 'should be all set',
      'is this what you', 'is this the',
    ]) || messages.some(m => m.text === '(no new text — customer forwarded previous reply)');

    const hasDispute = hasAny(['dispute', 'disputed', 'disputing', 'payment dispute',
      'chargeback dispute', 'open a dispute', 'opened a dispute', 'file a dispute',
      'raise a dispute', 'filing a dispute', 'i disputed']);

    const themes = [];
    const steps  = new Set();

    // 01 · AI-driven Frustration
    // Fires when: no HE has ever replied (open wait with no prior human contact),
    // OR customer explicitly signals AI friction in their text.
    // Does NOT fire just because the latest follow-up is unanswered — that's trigger 03.
    // mentionsAI uses hasAnyRecent (last 3 messages only) to avoid false positives from
    // early automated queue/bot messages ("A real person will be with you shortly") that
    // appear before the HE connects and are sometimes included in the thread.
    const mentionsAI = hasAnyRecent(['bot', 'odie', 'automated reply', 'automated response', 'robot',
      'real person', 'real human', 'speak to someone', 'talk to someone', 'actual person',
      'human support', 'human agent', 'not a human', 'chatbot', 'ai response']);
    if ((hasOpenWait && !hadHEReply) || mentionsAI) {
      themes.push({
        id: '01',
        name: 'AI-driven Frustration',
        detail: mentionsAI
          ? 'Customer signalled frustration with automated responses or difficulty reaching a human.'
          : 'No human has responded yet — this ticket needs immediate attention.',
      });
      if (mentionsAI) {
        steps.add('Acknowledge this is now a human response and take clear ownership of the issue.');
      } else {
        steps.add('Prioritise — no human has replied yet. Open by taking clear ownership of the issue.');
      }
    }

    // 02 · Misunderstood Issue or Intent
    // Fires when: a recent reply appeared to miss the point or felt generic.
    const missedPoint = recentCategories.support_failure ||
      hasAnyRecent(['not understanding', 'not addressing', "didn't help", "that didn't help",
               'wrong problem', 'not what i asked', 'not what i meant', 'misunderstood',
               'different issue', 'same answer', 'generic response', 'copy paste',
               'copy-paste', 'canned response', 'template response']);
    if (missedPoint) {
      themes.push({
        id: '02',
        name: 'Misunderstood Issue or Intent',
        detail: 'A previous reply may have addressed the wrong problem, felt generic, or missed the customer\'s specific context.',
      });
      steps.add("Mirror the customer's specific concern back before offering a solution.");
      steps.add('Reference specifics: product name, URL, order number, or the exact error described.');
    }

    // 03 · Perceived Inaction
    // Fires when: SLA was breached AND the customer shows a behavioral/emotional signal
    // (long multi-day wait, trailing unanswered messages, delay or emotion language),
    // OR when trailing unanswered >= 2 / delay / repetition independently.
    // A single wait breach on a cooperative customer without other signals is not enough —
    // that avoids false positives on professional customers who don't complain about wait times.
    const hasEmotionalOrBehavioral = recentCategories.emotion || recentCategories.support_failure ||
      recentCategories.delay || trailingUnanswered >= 2;
    const waitBreachWithSignal = hasWaitBreach && (hasLongWait || trailingUnanswered >= 2 || hasEmotionalOrBehavioral);
    if (waitBreachWithSignal || trailingUnanswered >= 2 || recentCategories.delay ||
        (recentCategories.repetition && msgCount >= 4 && (trailingUnanswered >= 1 || !hadHEReply))) {
      themes.push({
        id: '03',
        name: 'Perceived Inaction',
        detail: trailingUnanswered >= 2
          ? 'Customer has sent multiple follow-ups into silence — no human response received.'
          : 'Slow or absent replies have left this customer feeling ignored or stalled.',
      });
      if (cooperativeCustomer) {
        steps.add("This customer has completed what was asked of them. Before replying, confirm what you did with what they provided. If you still need more from them, acknowledge their effort first and explain specifically why.");
      } else if (hasWaitBreach || hasLongWait || trailingUnanswered >= 2) {
        steps.add('Acknowledge the wait time explicitly before anything else.');
      }
      if (msgCount >= 4) {
        steps.add('Break the loop — summarise everything tried so far and propose one specific next action.');
      }
      steps.add('Signal clear ownership: confirm you\'re handling this and commit to a follow-up time.');
    }

    // 04 · Process Gaps
    // Fires when: customer was bounced between teams, hit setup friction, or expected a different channel.
    // Also fires when the HE's replies show a pattern of internal escalation — polite/professional
    // customers rarely complain about being bounced, but the HE side reveals it.
    const processGapFromCustomer = hasAnyRecent(['transferred', 'bounced between', 'different team',
      'another team', 'wrong team', 'wrong department', 'live chat', 'phone support', 'phone call',
      'i expected', 'expected to', 'misrouted', 'redirected to']);
    const processGapFromHE = heConfirms(['reach out to our', 'reached out to our', 'working with our',
      'in touch with our', 'internal team', 'escalated internally', 'passed this on', 'passed this along',
      'marketplace team', 'partner team', 'vendor team', 'extensions team', 'our other team']);
    if (processGapFromCustomer || processGapFromHE) {
      themes.push({
        id: '04',
        name: 'Process Gaps',
        detail: processGapFromHE && !processGapFromCustomer
          ? 'This issue involves another internal team — the customer is waiting on a handoff they cannot see or control.'
          : 'Customer may have been bounced between teams, hit unexpected setup friction, or experienced a support channel mismatch.',
      });
      steps.add(processGapFromHE && !processGapFromCustomer
        ? 'Get a concrete answer from the other team before replying — another "I\'ve escalated internally" without an outcome will land badly.'
        : 'Clarify ownership — confirm who is handling this and what happens next.');
    }

    // 05 · Policy Frustration
    // Fires when: frustration is directed at the rule itself, not the support quality.
    // Also fires when the HE's own reply confirms a WooPayments/payment account issue —
    // customers often don't name the product, but the HE's context reveals it.
    // 'woopayments'/'woo payments' intentionally excluded from policyFromCustomer — customers name
    // the product in technical contexts constantly, which causes false positives. Those terms are
    // kept in policyFromHE where the HE's own reply provides meaningful context.
    const policyFromCustomer = hasAnyRecent(['policy', 'unfair', 'without notice', 'without warning',
      'no reason explained', 'not explained', 'account hold', 'funds held', 'suspended',
      'chargeback', 'refund denied', 'no refund', 'licensing', 'subscription terms',
      'dispute', 'disputed', 'disputing', 'payment dispute',
      'verification', 'verify my account', 'verify my identity', 'payout', 'payout delay',
      'payout held', 'stripe verification', 'stripe requirements']);
    const policyFromHE = heConfirms(['woopayments', 'woo payments', 'account hold', 'funds held',
      'account suspended', 'stripe account', 'payment account', 'payout', 'verification required',
      'stripe verification', 'stripe requirements']);
    if (policyFromCustomer || policyFromHE || (recentCategories.escalation && !recentCategories.support_failure)) {
      const hasVerification = hasAny(['verification', 'verify my account', 'verify my identity',
        'stripe verification', 'stripe requirements', 'identity check']);
      const hasPayoutIssue  = hasAny(['payout', 'payout delay', 'payout held', 'funds held']);
      const detail = hasDispute
        ? 'Customer has raised a payment dispute — a formal process involving their bank or payment provider that WooPayments cannot fully control.'
        : hasVerification
          ? 'Customer is frustrated with account verification or Stripe identity requirements — a process outside their control and often lacking clear timelines.'
          : hasPayoutIssue
            ? 'Customer is experiencing a payout delay or hold — funds feel inaccessible for reasons they cannot see or control.'
            : (policyFromHE && !policyFromCustomer)
              ? 'This is a WooPayments or payment account issue — a process the customer cannot see into. Opacity drives frustration even when support is responsive.'
              : 'Customer frustration is directed at a payment or account process they cannot control — this may relate to verification, payout timelines, dispute handling, or account policies.';
      themes.push({ id: '05', name: 'Payment / Policy Friction', detail });
      if (hasDispute) {
        steps.add('Acknowledge the dispute directly — name it, don\'t avoid it. The customer needs to know you understand the severity.');
        steps.add('Explain the dispute process clearly: what happens next, what evidence or action is needed, and a realistic timeline.');
        steps.add('Be honest about what WooPayments can and cannot influence — the final outcome may rest with the customer\'s bank.');
      } else {
        steps.add('Lead with empathy first. Offer a concrete next step or a clear timeline for resolution.');
        steps.add("Don't over-explain the process — focus on what you can do and by when.");
      }
    }

    // 06 · Outside Support's Control
    // Fires when: a bug, third-party product, or out-of-scope issue is the root cause.
    // Also checks HE reply text — bug confirmations come from the agent side, not the customer.
    if (hasAny(['known bug', 'known issue', 'feature request', 'third party', 'third-party',
                 'not supported', 'your plugin', 'your theme', 'hosting issue', 'server issue',
                 'misrouted', 'out of scope',
                 'plugin conflict', 'conflict with', 'conflicting with', 'compatibility issue',
                 'not compatible', 'incompatible']) ||
        heConfirms(['known issue', 'known bug', 'development team', 'our engineers',
                    'working on a fix', 'working toward a fix', 'working to fix',
                    'planned fix', 'release', 'workaround',
                    'third-party', 'third party', 'outside our control', 'not in scope',
                    'conflict', 'conflicting', 'plugin conflict', 'css conflict',
                    'jquery conflict', 'compatibility issue', 'not compatible',
                    'incompatible', 'interfering with', 'i tested', 'after testing',
                    'after further testing', 'i found that', 'i was able to reproduce',
                    'i can reproduce', 'confirmed the issue'])) {
      themes.push({
        id: '06',
        name: 'Product or Engineering Issue',
        detail: 'Root cause is a known bug, product gap, or third-party dependency — the fix lives with engineering, not support.',
      });
      steps.add("Be transparent about what's in scope — and offer the clearest available path forward.");
      // Only suggest checking Linear when it looks like an internal WooCommerce/Automattic bug.
      // For third-party plugin conflicts, the fix lives with the plugin developer, not engineering.
      const isThirdParty = heConfirms(['third-party', 'third party', 'outside our control',
        'plugin developer', 'theme developer', 'hosting provider', 'aspiring',
        'another plugin', 'another developer', 'conflict between', 'conflict with the',
        'forward to', 'forward this to', 'reach out to the']) ||
        hasAny(['third party', 'third-party', 'your plugin', 'your theme']);
      if (linearLinks.length > 0) {
        steps.add('A Linear issue is already linked in an internal note — check it for the latest status and share any update or timeline with the customer.');
      } else if (!isThirdParty) {
        steps.add('Check Linear for an open issue on this bug and follow up with the customer once there\'s an update.');
      } else {
        steps.add('Raise this with the relevant plugin or theme developer and share a workaround or timeline with the customer.');
      }
    }

    // 07 · Unresolved Feeling (outcome only — fallback when no clearer trigger applies)
    if (themes.length === 0 && msgCount >= 3 && (recentCategories.emotion || recentCategories.support_failure)) {
      themes.push({
        id: '07',
        name: 'Unresolved Feeling',
        detail: 'Issue may not be fully resolved — or the customer never received clear confirmation that it was.',
      });
      steps.add('Explicitly confirm what was resolved and state the next step or outcome clearly.');
    }

    // Cross-cutting: when a refund or cancellation was explicitly requested, always surface a direct action step.
    if (allCategories.escalation_request) {
      steps.add('Address the refund, dispute, or cancellation request directly — confirm what you can do and give a clear timeline.');
    }

    return { themes, nextSteps: [...steps] };
  }

  // ─── DOM Extraction ───────────────────────────────────────────────────────

  function getRequesterName() {
    const selectors = [
      '[data-test-id="ticket-system-field-requester-select"]',
      '[data-test-id="tooltip-requester-name"]',
      '[data-test-id="tabs-nav-item-users"]',
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      if (el) {
        const name = (el.value || el.textContent || '').trim();
        if (name) return name;
      }
    }
    return null;
  }

  function getRequesterEmail() {
    const wrapper = document.querySelector('[data-test-id="requester-field"]');
    if (wrapper) {
      // Check title attributes on any child element — Zendesk often puts the email in a title
      const withTitle = wrapper.querySelectorAll('[title]');
      for (const el of withTitle) {
        const t = (el.getAttribute('title') || '').trim();
        if (t.includes('@')) return t.toLowerCase();
      }
    }
    // Fallback: common selectors for the email value
    const selectors = [
      '[data-test-id="requester-field"] input[type="email"]',
      '[data-test-id="requester-field"] [data-garden-id="forms.input"]',
      '.requester [title*="@"]',
    ];
    for (const sel of selectors) {
      const el = document.querySelector(sel);
      const text = el && ((el.getAttribute('title') || el.value || el.textContent || '').trim());
      if (text && text.includes('@')) return text.toLowerCase();
    }
    return null;
  }

  function isExemptTicket() {
    const email = getRequesterEmail();
    if (email) {
      const domain = email.split('@')[1];
      if (domain && BYPASS_REQUESTER_DOMAINS.some(function(d) { return domain === d || domain.endsWith('.' + d); })) return true;
    }
    if (BYPASS_TAGS.length > 0) {
      // Primary: individual tag items (may not include all tags if list is long)
      const tags = new Set();
      document.querySelectorAll('[data-test-id="ticket-system-field-tags-item-selected"]').forEach(function(el) {
        var t = el.textContent.trim().toLowerCase();
        if (t) tags.add(t);
      });
      if (BYPASS_TAGS.some(function(tag) { return tags.has(tag); })) return true;

      // Fallback: full tag container text includes all tags even when list is truncated
      const fullTagText = (
        (document.querySelector('[data-test-id="ticket-system-field-tags-multiselect"]') ||
         document.querySelector('[data-test-id="ticket-fields-tags"]') || {}).textContent || ''
      ).toLowerCase();
      if (fullTagText && BYPASS_TAGS.some(function(tag) { return fullTagText.indexOf(tag) !== -1; })) return true;
    }
    return false;
  }

  function isCustomerComment(el, authorName, requesterName) {
    // Internal note check
    if (el.classList.contains('is-internal')) return false;
    if (el.querySelector('.internal-note')) return false;
    if (el.dataset.isInternal === 'true') return false;

    if (requesterName && authorName) {
      return authorName.toLowerCase().includes(requesterName.toLowerCase()) ||
             requesterName.toLowerCase().includes(authorName.toLowerCase());
    }

    if (authorName) {
      // No requester name — exclude if likely agent
      return !isLikelyAgentReply(authorName);
    }

    return true;
  }

  function getTextFromEl(el, selectors) {
    for (const sel of selectors) {
      const found = el.querySelector(sel);
      if (found) return found.innerText || found.textContent || '';
    }
    return el.innerText || el.textContent || '';
  }

  function getAuthorFromEl(el, selectors) {
    for (const sel of selectors) {
      const found = el.querySelector(sel);
      if (found) return (found.innerText || found.textContent || '').trim();
    }
    return '';
  }

  function strategy1(requesterName) {
    const comments = document.querySelectorAll('.comment');
    if (!comments.length) return null;

    const messages = [];
    for (const el of comments) {
      if (el.classList.contains('is-internal')) continue;
      if (el.querySelector('.internal-note')) continue;

      const author = getAuthorFromEl(el, [
        '.comment-author .author-name',
        '.comment-author',
        '.author .name',
      ]);

      // If requester name is known and author doesn't match → definitely skip (it's an HE reply)
      if (requesterName && author && !isCustomerComment(el, author, requesterName)) continue;

      // If requester name is unknown, fall back to text-based agent detection
      if (!requesterName || !author) {
        const text = getTextFromEl(el, ['.zd-comment', '.comment-body', '.rich-text-comment']);
        const cleaned = cleanText(text);
        if (cleaned && isLikelyAgentReply(cleaned)) continue;
      }

      const text = getTextFromEl(el, ['.zd-comment', '.comment-body', '.rich-text-comment']);
      const cleaned = cleanText(text);
      if (cleaned.length >= 10) {
        messages.push({ text: cleaned, author });
      }
    }

    return messages.length ? { messages, strategy: 'strategy1' } : null;
  }

  function strategy2(requesterName) {
    const comments = document.querySelectorAll('[data-comment-id]');
    if (!comments.length) return null;

    const messages = [];
    for (const el of comments) {
      if (el.dataset.isInternal === 'true') continue;
      if (el.closest('[data-is-internal="true"]')) continue;
      if (el.classList.contains('is-internal')) continue;

      const author = getAuthorFromEl(el, [
        '[data-test-id="author-name"]',
        '.author-name',
      ]);

      // If requester name is known and author doesn't match → skip
      if (requesterName && author && !isCustomerComment(el, author, requesterName)) continue;

      // If requester name is unknown, fall back to text-based agent detection
      if (!requesterName || !author) {
        const text = getTextFromEl(el, [
          '[data-test-id="comment-body"]',
          '.zd-comment',
          '.comment-body',
        ]);
        const cleaned = cleanText(text);
        if (cleaned && isLikelyAgentReply(cleaned)) continue;
      }

      const text = getTextFromEl(el, [
        '[data-test-id="comment-body"]',
        '.zd-comment',
        '.comment-body',
      ]);
      const cleaned = cleanText(text);
      if (cleaned.length >= 10) {
        messages.push({ text: cleaned, author });
      }
    }

    return messages.length ? { messages, strategy: 'strategy2' } : null;
  }

  function strategy3() {
    // Messaging / chat-style UIs where requester bubbles are on the left
    const selectors = [
      '.c-message-bubble--requester',
      '[data-side="left"]',
      '.message--incoming',
      '.conversation-message--requester',
    ];

    const messages = [];
    const seen = new Set();

    for (const sel of selectors) {
      const els = document.querySelectorAll(sel);
      for (const el of els) {
        if (seen.has(el)) continue;
        seen.add(el);
        const text = cleanText(el.innerText || el.textContent || '');
        if (text.length >= 10) {
          messages.push({ text, author: 'customer' });
        }
      }
    }

    return messages.length ? { messages, strategy: 'strategy3' } : null;
  }

  function strategy4(requesterName) {
    // Modern Zendesk Agent Workspace — tries a wide range of known selectors
    const containerSelectors = [
      '[data-test-id="ticket-events"]',
      '[data-test-id="conversation-log"]',
      '[data-test-id="omni-log-item-component"]',
      '[data-test-id="ticket-comment-event"]',
      'article[data-event-id]',
      'article[data-comment-id]',
      '[role="article"]',
      '.conversation-event',
      '.event-container',
      '.ticket-event',
      '.ember-view .comment',
    ];

    for (const containerSel of containerSelectors) {
      const containers = document.querySelectorAll(containerSel);
      if (!containers.length) continue;

      const messages = [];
      for (const container of containers) {
        // Skip internal notes
        const isInternal =
          container.classList.contains('is-internal') ||
          container.dataset.isInternal === 'true' ||
          !!container.querySelector('[data-test-id="internal-note-badge"], .internal-note-badge, .is-internal');
        if (isInternal) continue;

        // Try to get author
        const author = getAuthorFromEl(container, [
          '[data-test-id="author-name"]',
          '[data-test-id="event-author"]',
          '.author-name', '.actor-name', '.author .name',
          '[data-garden-id] span',
        ]);

        if (requesterName && author) {
          const match =
            author.toLowerCase().includes(requesterName.toLowerCase()) ||
            requesterName.toLowerCase().includes(author.toLowerCase());
          if (!match) continue;
        }

        // Get body text
        const body = getTextFromEl(container, [
          '[data-test-id="comment-body"]',
          '[data-test-id="rich-text-comment"]',
          '.zd-comment', '.comment-body', '.rich-text-comment',
          '[data-garden-id="notifications.notification"]',
        ]);
        const cleaned = cleanText(body);
        if (cleaned.length >= 10) {
          if (!requesterName && isLikelyAgentReply(cleaned)) continue;
          messages.push({ text: cleaned, author: author || 'customer' });
        }
      }

      if (messages.length > 0) return { messages, strategy: `strategy4:${containerSel}` };
    }

    return null;
  }

  // Primary strategy for modern Zendesk Agent Workspace.
  // Uses the omni-log structure confirmed from DOM inspection.
  // NOTE: dataset API shows camelCase keys (e.g. dataset.testId) but the actual
  // HTML attributes are hyphenated (data-test-id, data-originated-from).
  function strategyOmniLog(requesterName) {
    const articles = document.querySelectorAll('[data-test-id="omni-log-comment-item"]');
    if (!articles.length) return null;

    const messages = [];
    // Full ordered timeline of human messages (customer + agent) used for wait-time calculation
    const thread = [];
    // Linear issue links found in internal notes
    const linearLinks = [];

    for (const article of articles) {
      // Get the origin of this message (ApiInteraction / EmailInteraction / WebInteraction / etc.)
      const contentEl = article.querySelector('[data-test-id="omni-log-message-content"]');
      const originatedFrom = contentEl
        ? (contentEl.getAttribute('data-originated-from') || '')
        : '';

      // Internal notes: scan for Linear links before skipping.
      // Must run before the origin filter — internal notes sometimes have
      // originatedFrom === 'ApiInteraction' or 'Trigger' and would be dropped otherwise.
      const isInternal =
        !!article.querySelector('[data-test-id="omni-log-internal-note-tag"]') ||
        article.classList.contains('is-internal') ||
        article.dataset.isInternal === 'true';
      if (isInternal) {
        article.querySelectorAll('a[href*="linear.app"]').forEach(a => {
          if (a.href) linearLinks.push(a.href);
        });
        const articleText = article.textContent || '';
        const urlMatches = articleText.match(/https?:\/\/linear\.app\/[^\s"'<>]+/g);
        if (urlMatches) linearLinks.push(...urlMatches);
        continue;
      }

      // Skip system/bot/automation non-internal interactions entirely
      if (originatedFrom === 'ApiInteraction' || originatedFrom === 'Trigger') continue;

      // Get the comment body.
      // Email/web messages use .zd-comment; chat messages use omni-log-message-content.
      const bodyEl = article.querySelector('.zd-comment') ||
                     article.querySelector('[data-test-id="omni-log-message-content"]');
      if (!bodyEl) continue;
      const rawBodyText = bodyEl.innerText || bodyEl.textContent || '';
      const cleaned = cleanText(rawBodyText);
      // When quote-stripping empties the message (customer replied by forwarding the
      // previous reply with no new text above it), still count the reply in the thread.
      const text = cleaned.length >= 10 ? cleaned
        : rawBodyText.trim().length >= 10 ? '(no new text — customer forwarded previous reply)' : '';
      if (text.length < 10) continue;

      // Skip Zendesk system notification lines — these appear in the omni-log as
      // non-internal events but are not real customer or agent messages.
      if (/^last \d+ messages? were emailed/i.test(text) ||
          /^this conversation (is|was) (now )?(closed|resolved)/i.test(text) ||
          /^ticket (has been|was) (assigned|updated|closed|solved|reopened)/i.test(text) ||
          /^system status report/i.test(text)) continue;

      // Chat direction: AgentBadge is present on HE messages, absent on customer messages.
      // This is more reliable than phrase matching for chat.
      const hasAgentBadge = !!article.querySelector('[data-test-id="omni-log-avatar-badge-AgentBadge"]');

      // Author name (used when requester name is known)
      const author = getAuthorFromEl(article, [
        '[data-test-id="omni-log-item-sender"]',
        '[data-test-id*="author"]',
        '[data-test-id*="actor"]',
      ]);

      let isCustomer = false;

      if (hasAgentBadge) {
        // Agent badge is the most reliable signal — Zendesk sets it on all HE replies
        // sent through the web interface. Messages without it are inbound.
        isCustomer = false;
      } else if (requesterName && author) {
        // No agent badge — message is inbound (requester or another contact from their org).
        // HE replies always carry an agent badge; no-badge messages are treated as customer
        // regardless of whether the author name matches the requester exactly.
        isCustomer = true;
      } else if (originatedFrom === 'MobileSdkInteraction') {
        // Customer's own mobile/SDK submission — always inbound
        isCustomer = true;
      } else if (originatedFrom === 'EmailInteraction') {
        isCustomer = !isDefinitelyAgentReply(text);
      } else if (originatedFrom === '') {
        // Chat message with no origin marker and no agent badge.
        // Still check for agent language — automated welcome/queue messages also lack an origin.
        isCustomer = !isDefinitelyAgentReply(text);
      } else {
        // WebInteraction and other origins — check for agent language
        isCustomer = !isDefinitelyAgentReply(text);
      }

      const timestamp = getTimestamp(article);

      if (isCustomer) {
        messages.push({ text, author: author || originatedFrom, timestamp });
        thread.push({ type: 'customer', timestamp });
      } else {
        // HE/agent message — don't analyze text, but record in timeline for wait-time calc
        thread.push({ type: 'agent', timestamp, text });
      }
    }

    return messages.length ? { messages, strategy: 'omni-log', thread, linearLinks } : null;
  }

  function debugDom() {
    const report = [];
    report.push('=== Frustration Meter DOM Debug ===');
    report.push('URL: ' + window.location.href);
    report.push('Requester name found: ' + (getRequesterName() || 'NONE'));

    const probes = [
      // Confirmed working selectors (omni-log structure, hyphenated as in HTML)
      '[data-test-id="omni-log-comment-item"]',
      '[data-test-id="omni-log-message-content"]',
      '[data-test-id="omni-log-item-message"]',
      '[data-test-id="omni-log-container"]',
      '.zd-comment',
      // Classic Zendesk
      '.comment', '[data-comment-id]', '.comment-body',
      // Other modern selectors
      '[data-test-id="omni-log-item-component"]', '[data-test-id="ticket-comment-event"]',
      '[data-test-id="comment-body"]', '[data-test-id="rich-text-comment"]',
      '[data-test-id="ticket-events"]', '[data-test-id="conversation-log"]',
      '[data-test-id="author-name"]',
      '[role="article"]', 'article[data-event-id]', 'article[data-comment-id]',
      '.c-message-bubble--requester', '[data-side="left"]',
      '.actor-name', '.author-name',
    ];

    for (const sel of probes) {
      const found = document.querySelectorAll(sel);
      if (found.length > 0) {
        report.push(`FOUND [${sel}]: ${found.length} element(s)`);
        // Show first element's outerHTML snippet
        const snippet = found[0].outerHTML.slice(0, 300).replace(/\s+/g, ' ');
        report.push('  First: ' + snippet);
      } else {
        report.push(`      [${sel}]: 0`);
      }
    }

    const output = report.join('\n');
    console.log(output);
    return output;
  }

  function extractMessages() {
    const requesterName = getRequesterName();
    const result =
      strategyOmniLog(requesterName) ||  // Modern Zendesk Agent Workspace
      strategy1(requesterName) ||
      strategy2(requesterName) ||
      strategy3() ||
      strategy4(requesterName);

    return {
      messages: result ? result.messages : [],
      strategy: result ? result.strategy : 'none',
      requesterName: requesterName || 'Unknown',
      thread: (result && result.thread) ? result.thread : [],
      linearLinks: (result && result.linearLinks) ? result.linearLinks : [],
    };
  }

  function analyzeSelectedText() {
    const sel = window.getSelection();
    if (!sel || !sel.toString().trim()) {
      return null;
    }
    const text = cleanText(sel.toString());
    if (text.length < 10) return null;
    return { text, author: 'selected' };
  }

  // ─── Styles ───────────────────────────────────────────────────────────────

  function injectStyles() {
    if (document.getElementById(STYLE_ID)) return;
    const style = document.createElement('style');
    style.id = STYLE_ID;
    style.textContent = `
      #${PANEL_ID} * {
        box-sizing: border-box;
        font-family: -apple-system, BlinkMacSystemFont, "Segoe UI", Arial, sans-serif;
      }
      #${PANEL_ID} {
        position: fixed;
        top: 60px;
        right: 16px;
        width: 340px;
        max-height: 88vh;
        background: #ffffff;
        border: 1px solid #d1d5db;
        border-radius: 8px;
        box-shadow: 0 4px 24px rgba(0,0,0,0.13);
        z-index: 999999;
        display: flex;
        flex-direction: column;
        overflow: hidden;
        font-size: 13px;
        color: #111827;
      }
      #${PANEL_ID} .fm-header {
        display: flex;
        align-items: center;
        justify-content: space-between;
        padding: 10px 12px;
        background: #f9fafb;
        border-bottom: 1px solid #e5e7eb;
        cursor: grab;
        user-select: none;
        flex-shrink: 0;
      }
      #${PANEL_ID} .fm-header:active {
        cursor: grabbing;
      }
      #${PANEL_ID} .fm-header-title {
        font-weight: 600;
        font-size: 13px;
        color: #374151;
        letter-spacing: 0.01em;
      }
      #${PANEL_ID} .fm-header-right {
        display: flex;
        align-items: center;
        gap: 6px;
      }
      #${PANEL_ID} .fm-level-badge {
        font-size: 11px;
        font-weight: 700;
        padding: 2px 8px;
        border-radius: 12px;
        color: #fff;
        letter-spacing: 0.03em;
      }
      #${PANEL_ID} .fm-collapse-btn {
        background: none;
        border: 1px solid #d1d5db;
        border-radius: 4px;
        width: 22px;
        height: 22px;
        cursor: pointer;
        font-size: 14px;
        line-height: 1;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #6b7280;
        padding: 0;
        flex-shrink: 0;
      }
      #${PANEL_ID} .fm-collapse-btn:hover {
        background: #f3f4f6;
        color: #374151;
      }
      #${PANEL_ID} .fm-body {
        overflow-y: auto;
        padding: 12px;
        flex: 1;
        min-height: 0;
      }
      #${PANEL_ID} .fm-level-row {
        display: flex;
        align-items: baseline;
        gap: 8px;
        margin-bottom: 6px;
      }
      #${PANEL_ID} .fm-level-text {
        font-size: 22px;
        font-weight: 700;
        line-height: 1.1;
      }
      #${PANEL_ID} .fm-score-text {
        font-size: 13px;
        color: #6b7280;
      }
      #${PANEL_ID} .fm-count-text {
        font-size: 12px;
        color: #9ca3af;
        margin-bottom: 8px;
      }
      #${PANEL_ID} .fm-meter-bar-wrap {
        height: 8px;
        background: #e5e7eb;
        border-radius: 4px;
        overflow: hidden;
        margin-bottom: 10px;
      }
      #${PANEL_ID} .fm-meter-bar-fill {
        height: 100%;
        border-radius: 4px;
        transition: width 0.4s ease;
      }
      #${PANEL_ID} .fm-trend-row {
        font-size: 12px;
        color: #6b7280;
        margin-bottom: 10px;
      }
      #${PANEL_ID} .fm-section-label {
        font-size: 11px;
        font-weight: 600;
        color: #9ca3af;
        text-transform: uppercase;
        letter-spacing: 0.07em;
        margin: 10px 0 6px 0;
      }
      #${PANEL_ID} .fm-timeline {
        display: flex;
        align-items: center;
        flex-wrap: wrap;
        gap: 0;
        margin-bottom: 10px;
      }
      #${PANEL_ID} .fm-dot-wrap {
        display: flex;
        align-items: center;
      }
      #${PANEL_ID} .fm-dot {
        width: 24px;
        height: 24px;
        border-radius: 50%;
        display: flex;
        align-items: center;
        justify-content: center;
        color: #fff;
        font-size: 10px;
        font-weight: 700;
        cursor: default;
        flex-shrink: 0;
        border: 2px solid rgba(255,255,255,0.6);
        box-shadow: 0 1px 3px rgba(0,0,0,0.15);
      }
      #${PANEL_ID} .fm-connector {
        width: 12px;
        height: 2px;
        background: #d1d5db;
        flex-shrink: 0;
      }
      #${PANEL_ID} .fm-pills {
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
        margin-bottom: 4px;
      }
      #${PANEL_ID} .fm-pill {
        font-size: 11px;
        font-weight: 600;
        padding: 3px 8px;
        border-radius: 10px;
        color: #fff;
        cursor: default;
      }
      #${PANEL_ID} .fm-phrases {
        list-style: none;
        padding: 0;
        margin: 0 0 6px 0;
      }
      #${PANEL_ID} .fm-phrases li {
        font-size: 12px;
        padding: 2px 0;
        color: #374151;
        display: flex;
        align-items: baseline;
        gap: 6px;
      }
      #${PANEL_ID} .fm-phrase-cat {
        font-size: 10px;
        font-weight: 600;
        color: #9ca3af;
        text-transform: uppercase;
        letter-spacing: 0.05em;
        flex-shrink: 0;
      }
      #${PANEL_ID} .fm-boosts {
        background: #fffbeb;
        border: 1px solid #fde68a;
        border-radius: 6px;
        padding: 6px 8px;
        margin-bottom: 8px;
      }
      #${PANEL_ID} .fm-boosts-title {
        font-size: 11px;
        font-weight: 600;
        color: #92400e;
        margin-bottom: 4px;
      }
      #${PANEL_ID} .fm-boost-item {
        font-size: 11px;
        color: #78350f;
        padding: 1px 0;
      }
      #${PANEL_ID} .fm-status {
        font-size: 11px;
        color: #9ca3af;
        margin-top: 8px;
        border-top: 1px solid #f3f4f6;
        padding-top: 8px;
      }
      #${PANEL_ID} .fm-no-results {
        font-size: 13px;
        color: #6b7280;
        line-height: 1.5;
      }
      #${PANEL_ID} .fm-footer {
        display: flex;
        flex-wrap: wrap;
        gap: 5px;
        padding: 8px 12px;
        background: #f9fafb;
        border-top: 1px solid #e5e7eb;
        flex-shrink: 0;
      }
      #${PANEL_ID} .fm-btn {
        font-size: 11px;
        font-weight: 500;
        padding: 4px 9px;
        border-radius: 5px;
        cursor: pointer;
        border: 1px solid #d1d5db;
        background: #fff;
        color: #374151;
        transition: background 0.15s;
        white-space: nowrap;
      }
      #${PANEL_ID} .fm-btn:hover {
        background: #f3f4f6;
      }
      #${PANEL_ID} .fm-btn-primary {
        background: #2563eb;
        color: #fff;
        border-color: #2563eb;
      }
      #${PANEL_ID} .fm-btn-primary:hover {
        background: #1d4ed8;
      }
      #${PANEL_ID} .fm-copy-feedback {
        color: #2da44e;
        font-size: 11px;
        align-self: center;
        display: none;
      }
      #${PANEL_ID} .fm-context {
        background: #f5f3ff;
        border: 1px solid #ddd6fe;
        border-radius: 6px;
        padding: 8px 10px;
        margin-bottom: 8px;
      }
      #${PANEL_ID} .fm-context-theme {
        margin-bottom: 6px;
      }
      #${PANEL_ID} .fm-context-theme:last-child {
        margin-bottom: 0;
      }
      #${PANEL_ID} .fm-context-theme-name {
        display: block;
        font-size: 11px;
        font-weight: 700;
        color: #5b21b6;
        margin-bottom: 1px;
      }
      #${PANEL_ID} .fm-context-theme-detail {
        display: block;
        font-size: 11px;
        color: #4c1d95;
        line-height: 1.4;
      }
      #${PANEL_ID} .fm-next-steps {
        list-style: none;
        padding: 0;
        margin: 0;
        background: #eff6ff;
        border: 1px solid #bfdbfe;
        border-radius: 6px;
        padding: 6px 10px;
        margin-bottom: 8px;
      }
      #${PANEL_ID} .fm-next-steps li {
        font-size: 11px;
        color: #1e3a8a;
        padding: 2px 0;
        line-height: 1.4;
      }
      #${PANEL_ID} .fm-next-steps li::before {
        content: "→ ";
        font-weight: 700;
      }
      #${PANEL_ID} .fm-followup {
        background: #f0fdf4;
        border: 1px solid #bbf7d0;
        border-radius: 6px;
        padding: 8px 10px;
        margin-bottom: 8px;
      }
      #${PANEL_ID} .fm-followup.fm-followup-overdue {
        background: #fffbeb;
        border-color: #fde68a;
      }
      #${PANEL_ID} .fm-followup-title {
        font-size: 11px;
        font-weight: 700;
        color: #14532d;
        margin-bottom: 3px;
      }
      #${PANEL_ID} .fm-followup.fm-followup-overdue .fm-followup-title {
        color: #92400e;
      }
      #${PANEL_ID} .fm-followup-status {
        font-size: 11px;
        color: #166534;
        line-height: 1.5;
      }
      #${PANEL_ID} .fm-followup.fm-followup-overdue .fm-followup-status {
        color: #92400e;
      }
      #${PANEL_ID} .fm-followup-nudges {
        margin-top: 6px;
        padding-top: 6px;
        border-top: 1px solid #bbf7d0;
      }
      #${PANEL_ID} .fm-followup.fm-followup-overdue .fm-followup-nudges {
        border-top-color: #fde68a;
      }
      #${PANEL_ID} .fm-followup-nudge {
        font-size: 11px;
        color: #14532d;
        line-height: 1.5;
        padding: 1px 0;
      }
      #${PANEL_ID} .fm-followup.fm-followup-overdue .fm-followup-nudge {
        color: #78350f;
      }
      #${PANEL_ID}.fm-collapsed .fm-body {
        display: none;
      }
      #${PANEL_ID}.fm-collapsed .fm-footer {
        display: none;
      }
    `;
    document.head.appendChild(style);
  }

  // ─── Panel ────────────────────────────────────────────────────────────────

  function getOrCreatePanel() {
    let panel = document.getElementById(PANEL_ID);
    if (!panel) {
      panel = document.createElement('div');
      panel.id = PANEL_ID;
      document.body.appendChild(panel);
      initDrag(panel);
      restorePosition(panel);
    }
    return panel;
  }

  function restorePosition(panel) {
    try {
      const saved = localStorage.getItem(POSITION_KEY);
      if (saved) {
        const { top, left, right } = JSON.parse(saved);
        if (top !== undefined) panel.style.top = top;
        if (left !== undefined) {
          panel.style.left = left;
          panel.style.right = 'auto';
        } else if (right !== undefined) {
          panel.style.right = right;
          panel.style.left = 'auto';
        }
      }
    } catch (e) {
      // ignore
    }
  }

  function savePosition(panel) {
    try {
      const rect = panel.getBoundingClientRect();
      localStorage.setItem(POSITION_KEY, JSON.stringify({
        top: rect.top + 'px',
        left: rect.left + 'px',
      }));
    } catch (e) {
      // ignore
    }
  }

  function initDrag(panel) {
    let dragging = false;
    let startX, startY, origLeft, origTop;

    const onMouseMove = (e) => {
      if (!dragging) return;
      const dx = e.clientX - startX;
      const dy = e.clientY - startY;
      const newLeft = Math.max(0, Math.min(window.innerWidth - panel.offsetWidth, origLeft + dx));
      const newTop = Math.max(0, Math.min(window.innerHeight - 60, origTop + dy));
      panel.style.left = newLeft + 'px';
      panel.style.top = newTop + 'px';
      panel.style.right = 'auto';
    };

    const onMouseUp = () => {
      if (dragging) {
        dragging = false;
        savePosition(panel);
      }
    };

    panel.addEventListener('mousedown', (e) => {
      if (!e.target.closest('.fm-header')) return;
      if (e.target.closest('.fm-collapse-btn')) return;
      dragging = true;
      startX = e.clientX;
      startY = e.clientY;
      const rect = panel.getBoundingClientRect();
      origLeft = rect.left;
      origTop = rect.top;
      e.preventDefault();
    });

    document.addEventListener('mousemove', onMouseMove);
    document.addEventListener('mouseup', onMouseUp);
  }

  // ─── Render ───────────────────────────────────────────────────────────────

  function renderPanel(data) {
    const panel = getOrCreatePanel();
    const wasCollapsed = panel.classList.contains('fm-collapsed');

    // Determine display data — use latest message if multiple
    const latestMsg = data.messages.length > 0
      ? data.messages[data.messages.length - 1]
      : null;
    const latestAnalysis = latestMsg ? latestMsg.analysis : null;
    const { themes, nextSteps } = detectFrustrationThemes(data.messages, data.thread || [], data.linearLinks || []);
    const rawTrend = computeTrend(data.messages, data.thread || [], data.isSolved);
    // A known product/engineering issue hasn't truly improved just because the customer's
    // language softened after the HE acknowledged it — override a misleading downward trend.
    const trend = (rawTrend.label === 'Decreasing' && themes.some(t => t.id === '06'))
      ? { label: 'Stable', color: '#6b7280' }
      : rawTrend;
    const levelInfo = latestAnalysis ? getLevel(latestAnalysis.score) : null;

    // Detect whether the HE has already replied (last event in thread is from an agent).
    // When true, swap "Suggested Next Steps" for a follow-up check prompt.
    const thread = data.thread || [];
    const lastThreadEvent = thread.length > 0 ? thread[thread.length - 1] : null;
    const heRepliedLast = !!(lastThreadEvent && lastThreadEvent.type === 'agent');
    const timeSinceReplyHrs = (heRepliedLast && lastThreadEvent.timestamp instanceof Date)
      ? (Date.now() - lastThreadEvent.timestamp.getTime()) / 3600000
      : null;
    const badgeLabel = levelInfo ? levelInfo.label : '—';
    const badgeColor = levelInfo ? levelInfo.color : '#9ca3af';

    let html = `
      <div class="fm-header">
        <span class="fm-header-title">Frustration Signals</span>
        <div class="fm-header-right">
          <span class="fm-level-badge" style="background:${escapeHtml(badgeColor)}">${escapeHtml(badgeLabel)}</span>
          <button class="fm-collapse-btn" title="Toggle panel">${wasCollapsed ? '+' : '–'}</button>
        </div>
      </div>
    `;

    if (data.messages.length === 0) {
      html += `
        <div class="fm-body">
          <div class="fm-no-results">
            Could not detect customer messages.<br>
            <strong>Tip:</strong> Try clicking Rescan after the ticket finishes loading.
          </div>
          ${data.requesterName && data.requesterName !== 'Unknown' ? `<div class="fm-status">Requester: <strong>${escapeHtml(data.requesterName)}</strong></div>` : ''}
        </div>
        <div class="fm-footer">
          <button class="fm-btn fm-btn-primary fm-rescan-btn">Rescan</button>
          <button class="fm-btn fm-debug-btn">Debug DOM</button>
          <button class="fm-btn fm-reset-btn">Reset Panel</button>
        </div>
      `;
    } else {
      const pct = Math.min(100, Math.round((latestAnalysis.score / 16) * 100));

      // Timeline
      let timelineHtml = '';
      if (data.messages.length >= 2) {
        timelineHtml = `<div class="fm-section-label">Progression</div><div class="fm-timeline">`;
        data.messages.forEach((msg, i) => {
          const a = msg.analysis;
          const lvl = getLevel(a.score);
          const ts = formatTimestamp(msg.timestamp);
          // Timestamp in tooltip only — visible labels clutter the timeline for long threads
          const dotTitle = `Reply ${i + 1}: ${lvl.label} (score ${a.score})${ts ? ' · ' + ts : ''}`;
          if (i > 0) {
            timelineHtml += `<div class="fm-dot-wrap"><div class="fm-connector"></div></div>`;
          }
          timelineHtml += `<div class="fm-dot-wrap"><div class="fm-dot" style="background:${escapeHtml(lvl.color)}" title="${escapeHtml(dotTitle)}">${i + 1}</div></div>`;
        });
        timelineHtml += `</div>`;
      }

      // Category pills
      const catEntries = Object.entries(latestAnalysis.categories);
      let pillsHtml = '';
      if (catEntries.length > 0) {
        pillsHtml = `<div class="fm-section-label">Signal Categories</div><div class="fm-pills">`;
        for (const [catKey, phrases] of catEntries) {
          const cat = CATEGORIES[catKey];
          const phraseList = phrases.map(p => escapeHtml(p)).join(', ');
          pillsHtml += `<span class="fm-pill" style="background:${escapeHtml(cat.color)}" title="Matched: ${phraseList}">${escapeHtml(cat.label)}</span>`;
        }
        pillsHtml += `</div>`;
      }

      // Detected phrases
      let phrasesHtml = '';
      if (latestAnalysis.phrases.length > 0) {
        phrasesHtml = `<div class="fm-section-label">Detected Phrases</div><ul class="fm-phrases">`;
        for (const { phrase, category } of latestAnalysis.phrases) {
          const cat = CATEGORIES[category];
          phrasesHtml += `<li><span class="fm-phrase-cat">${escapeHtml(cat.label)}</span>"${escapeHtml(phrase)}"</li>`;
        }
        phrasesHtml += `</ul>`;
      }

      // Boosts
      let boostsHtml = '';
      if (latestAnalysis.boosts.length > 0) {
        boostsHtml = `<div class="fm-boosts"><div class="fm-boosts-title">Score Boosts</div>`;
        for (const boost of latestAnalysis.boosts) {
          boostsHtml += `<div class="fm-boost-item">+${boost.value} — ${escapeHtml(boost.reason)}</div>`;
        }
        boostsHtml += `</div>`;
      }

      // Frustration trigger context
      let contextHtml = '';
      if (themes.length > 0) {
        contextHtml += `<div class="fm-section-label">Frustration Trigger Context</div>`;
        contextHtml += `<div class="fm-context">`;
        for (const t of themes) {
          contextHtml += `<div class="fm-context-theme">
            <span class="fm-context-theme-name">${t.id ? escapeHtml(t.id + ' · ' + t.name) : escapeHtml(t.name)}</span>
            <span class="fm-context-theme-detail">${escapeHtml(t.detail)}</span>
          </div>`;
        }
        contextHtml += `</div>`;
      }

      if (heRepliedLast) {
        // HE has replied — swap next steps for a follow-up check prompt.
        const sla = getTicketSla();
        let isOverdue = false;
        let statusText = 'You replied last — awaiting customer response.';

        if (timeSinceReplyHrs !== null) {
          const timeLabel = formatWaitDuration(timeSinceReplyHrs);
          const overdueThreshold = sla ? sla.hours : 24;
          if (timeSinceReplyHrs > overdueThreshold) {
            isOverdue = true;
            const ctx = sla ? ` (${sla.label} SLA: ${sla.hours}h)` : '';
            statusText = `No customer response in ${timeLabel}${ctx} — consider a follow-up if the ticket is still open.`;
          } else {
            statusText = `You replied last — awaiting customer response (${timeLabel} ago).`;
          }
        }

        const nudges = [];
        const themeIds = themes.map(t => t.id);
        if (levelInfo && (levelInfo.label === 'High' || levelInfo.label === 'Severe')) {
          nudges.push('High-frustration ticket — when the customer replies, confirm all concerns are resolved before closing.');
        }
        if (themeIds.includes('01')) {
          nudges.push('Customer had no prior human contact — verify your reply acknowledged you are a real person taking ownership.');
        }
        if (themeIds.includes('02')) {
          nudges.push('A previous reply may have missed the point — check your response addressed their specific concern directly.');
        }
        if (themeIds.includes('06')) {
          nudges.push('Root cause is a known bug or engineering issue — follow up when there is a fix or update to share.');
        }

        const overdueClass = isOverdue ? ' fm-followup-overdue' : '';
        const titleText = isOverdue ? 'Follow-up may be needed' : 'Awaiting customer response';
        contextHtml += `
          <div class="fm-section-label">Follow-up Check</div>
          <div class="fm-followup${overdueClass}">
            <div class="fm-followup-title">${escapeHtml(titleText)}</div>
            <div class="fm-followup-status">${escapeHtml(statusText)}</div>
            ${nudges.length > 0 ? `<div class="fm-followup-nudges">${nudges.map(n => `<div class="fm-followup-nudge">→ ${escapeHtml(n)}</div>`).join('')}</div>` : ''}
          </div>`;
      } else if (nextSteps.length > 0) {
        contextHtml += `<div class="fm-section-label">Suggested Next Steps</div><ul class="fm-next-steps">`;
        for (const step of nextSteps) {
          contextHtml += `<li>${escapeHtml(step)}</li>`;
        }
        contextHtml += `</ul>`;
      }

      // Trend
      const trendHtml = data.messages.length >= 2
        ? `<div class="fm-trend-row">Trend: <strong style="color:${escapeHtml(trend.color)}">${trend.label === 'Increasing' ? '↑' : trend.label === 'Decreasing' ? '↓' : '→'} ${escapeHtml(trend.label)}</strong></div>`
        : '';

      html += `
        <div class="fm-body">
          <div class="fm-level-row">
            <span class="fm-level-text" style="color:${escapeHtml(levelInfo.color)}">${escapeHtml(levelInfo.label)}</span>
            <span class="fm-score-text">Score: ${latestAnalysis.score}</span>
          </div>
          <div class="fm-count-text">${data.messages.length} customer repl${data.messages.length === 1 ? 'y' : 'ies'} analyzed</div>
          <div class="fm-meter-bar-wrap">
            <div class="fm-meter-bar-fill" style="width:${pct}%;background:${escapeHtml(levelInfo.color)}"></div>
          </div>
          ${trendHtml}
          ${timelineHtml}
          ${pillsHtml}
          ${phrasesHtml}
          ${boostsHtml}
          ${contextHtml}
          <div class="fm-status">${data.isSolved ? '<strong>Ticket Solved</strong> · ' : ''}Requester: <strong>${escapeHtml(data.requesterName)}</strong> · Detection: ${escapeHtml(data.strategy)}</div>
        </div>
        <div class="fm-footer">
          <button class="fm-btn fm-btn-primary fm-rescan-btn">Rescan</button>
          <button class="fm-btn fm-copy-btn">Copy Summary</button>
          <button class="fm-btn fm-reset-btn">Reset Panel</button>
          <span class="fm-copy-feedback">Copied!</span>
        </div>
      `;
    }

    if (data.isExempt) {
      html = `
        <div class="fm-header">
          <span class="fm-header-title">Frustration Signals</span>
          <div class="fm-header-right">
            <span class="fm-level-badge" style="background:#6b7280">N/A</span>
            <button class="fm-collapse-btn" title="Toggle panel">${wasCollapsed ? '+' : '–'}</button>
          </div>
        </div>
        <div class="fm-body">
          <div class="fm-no-results">Stripe partner ticket — frustration scoring does not apply.</div>
        </div>
        <div class="fm-footer">
          <button class="fm-btn fm-btn-primary fm-rescan-btn">Rescan</button>
          <button class="fm-btn fm-reset-btn">Reset Panel</button>
        </div>
      `;
    }

    panel.innerHTML = html;

    if (wasCollapsed) panel.classList.add('fm-collapsed');

    // Wire up buttons
    const collapseBtn = panel.querySelector('.fm-collapse-btn');
    collapseBtn.addEventListener('click', (e) => {
      e.stopPropagation();
      panel.classList.toggle('fm-collapsed');
      collapseBtn.textContent = panel.classList.contains('fm-collapsed') ? '+' : '–';
    });

    panel.querySelector('.fm-rescan-btn')?.addEventListener('click', () => {
      runAnalysis();
    });

    panel.querySelector('.fm-debug-btn')?.addEventListener('click', () => {
      const output = debugDom();
      alert('Debug output logged to browser console (F12 → Console).\n\nQuick summary:\n' + output.split('\n').filter(l => l.includes('FOUND')).join('\n') || 'No matching elements found with any known selector.');
    });

    panel.querySelector('.fm-copy-btn')?.addEventListener('click', () => {
      if (!latestAnalysis) return;
      const catLabels = Object.keys(latestAnalysis.categories)
        .map(k => CATEGORIES[k].label)
        .join(', ');
      const phraseList = latestAnalysis.phrases
        .map(p => p.phrase)
        .join(', ');
      const summary = [
        `Frustration level: ${levelInfo.label} (score: ${latestAnalysis.score})`,
        `Trend: ${trend.label}`,
        `Categories: ${catLabels || 'None'}`,
        `Signals: ${phraseList || 'None'}`,
      ].join('\n');
      navigator.clipboard.writeText(summary).then(() => {
        const fb = panel.querySelector('.fm-copy-feedback');
        if (fb) {
          fb.style.display = 'inline';
          setTimeout(() => { fb.style.display = 'none'; }, 2000);
        }
      }).catch(() => {
        // Fallback
        const ta = document.createElement('textarea');
        ta.value = summary;
        document.body.appendChild(ta);
        ta.select();
        document.execCommand('copy');
        document.body.removeChild(ta);
        const fb = panel.querySelector('.fm-copy-feedback');
        if (fb) {
          fb.style.display = 'inline';
          setTimeout(() => { fb.style.display = 'none'; }, 2000);
        }
      });
    });

    panel.querySelector('.fm-reset-btn')?.addEventListener('click', () => {
      try { localStorage.removeItem(POSITION_KEY); } catch (e) {}
      panel.style.top = '60px';
      panel.style.right = '16px';
      panel.style.left = 'auto';
      runAnalysis();
    });
  }

  // ─── Main analysis runner ─────────────────────────────────────────────────

  function runAnalysis() {
    injectStyles();

    if (isExemptTicket()) {
      renderPanel({ messages: [], strategy: 'exempt', requesterName: getRequesterName() || 'Unknown', thread: [], linearLinks: [], isExempt: true });
      return;
    }

    const { messages: rawMessages, strategy, requesterName, thread, linearLinks } = extractMessages();

    // If nothing found, the conversation may still be rendering — retry automatically.
    // Stop retrying once messages are found or after 4 attempts (~10s total).
    if (rawMessages.length === 0 && retryCount < 4) {
      retryCount++;
      if (retryTimer) clearTimeout(retryTimer);
      retryTimer = setTimeout(runAnalysis, 2500);
    } else if (rawMessages.length > 0) {
      retryCount = 0;
      if (retryTimer) { clearTimeout(retryTimer); retryTimer = null; }
    }

    const messages = rawMessages.map(msg => ({
      ...msg,
      analysis: analyzeMessage(msg.text),
    }));

    const isSolved = ['solved', 'closed'].includes(getTicketStatus());

    if (messages.length > 0) {
      const lastAnalysis = messages[messages.length - 1].analysis;

      // Wait time boost — how long the customer waited for a human HE reply (bots excluded)
      const waitBoosts = computeWaitTimeBoosts(thread, isSolved);
      for (const boost of waitBoosts) {
        lastAnalysis.boosts.push(boost);
        lastAnalysis.score += boost.value;
      }

      if (waitBoosts.length > 0) {
        const updatedLevel = getLevel(lastAnalysis.score);
        lastAnalysis.level = updatedLevel.label;
        lastAnalysis.color = updatedLevel.color;
      }
    }

    renderPanel({ messages, strategy, requesterName, thread, linearLinks, isSolved });
  }

  // ─── SPA Navigation ───────────────────────────────────────────────────────

  let currentTicketId = null;
  let navTimer = null;
  let retryCount = 0;
  let retryTimer = null;

  function getTicketId() {
    const match = window.location.pathname.match(/\/tickets\/(\d+)/);
    return match ? match[1] : null;
  }

  function onPossibleNavigation() {
    const id = getTicketId();
    if (id && id !== currentTicketId) {
      currentTicketId = id;
      retryCount = 0;
      if (retryTimer) { clearTimeout(retryTimer); retryTimer = null; }
      if (navTimer) clearTimeout(navTimer);
      navTimer = setTimeout(() => {
        runAnalysis();
      }, NAV_DELAY);
    }
  }

  function patchHistoryMethod(method) {
    const original = history[method];
    history[method] = function (...args) {
      const result = original.apply(this, args);
      onPossibleNavigation();
      return result;
    };
  }

  function initSpaNavigation() {
    patchHistoryMethod('pushState');
    patchHistoryMethod('replaceState');
    window.addEventListener('popstate', onPossibleNavigation);

    // MutationObserver as a catch-all
    let mutationTimer = null;
    const observer = new MutationObserver(() => {
      if (mutationTimer) return;
      mutationTimer = setTimeout(() => {
        mutationTimer = null;
        onPossibleNavigation();
      }, 500);
    });
    observer.observe(document.body, { childList: true, subtree: true });
  }

  // ─── Bootstrap ────────────────────────────────────────────────────────────

  function init() {
    currentTicketId = getTicketId();
    initSpaNavigation();

    // Initial load
    setTimeout(() => {
      runAnalysis();
    }, INITIAL_DELAY);

    // Also fire on window load in case DOM isn't ready yet
    window.addEventListener('load', () => {
      setTimeout(() => {
        runAnalysis();
      }, LOAD_DELAY);
    });
  }

  if (document.readyState === 'loading') {
    document.addEventListener('DOMContentLoaded', init);
  } else {
    init();
  }

})();
