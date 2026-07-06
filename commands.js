/* ============================================================
   Lives International - Outlook Add-in
   Insereaza automat semnatura la compose/reply/forward,
   pozitionata corect deasupra textului citat.
   ============================================================ */

Office.onReady(() => {
  // Office.js incarcat, functiile sunt asociate mai jos
});

// ---- Configurare ----
const LOGO_URL = "https://medisolassets.blob.core.windows.net/logo-uri/logo-lives-signature.png";
const LINKEDIN_ICON_URL = "https://medisolassets.blob.core.windows.net/logo-uri/linkedin-icon-256x256.png";
const LINKEDIN_COMPANY_URL = "https://www.linkedin.com/company/lives-international/posts/?feedView=all";
const COMPANY_WEBSITE = "https://lives-international.com";
const GRAPH_ENDPOINT = "https://graph.microsoft.com/v1.0/me?$select=displayName,jobTitle,mobilePhone,mail,userPrincipalName";
const EVENT_TIMEOUT_MS = 7000;
const TOKEN_TIMEOUT_MS = 2500;
const GRAPH_TIMEOUT_MS = 3000;

// ---- Handler: se declanseaza automat la fiecare email nou (compose/reply/forward) ----
// function onNewMessageComposeHandler(event) {
//   insertSignature()
//     .then(() => event.completed())
//     .catch((err) => {
//       console.error("Eroare la inserarea automata a semnaturii:", err);
//       event.completed();
//     });
// }
function onNewMessageComposeHandler(event) {
  const finishEvent = createEventFinisher(event);
  const timeoutId = setTimeout(() => {
    setMarkerSignature("EVENT OK, TIMEOUT")
      .catch((err) => {
        console.error("Timeout fallback failed:", err);
      })
      .then(finishEvent, finishEvent);
  }, EVENT_TIMEOUT_MS);

  insertSignature(false)
    .then(() => {
      clearTimeout(timeoutId);
      finishEvent();
    })
    .catch((err) => {
      clearTimeout(timeoutId);
      console.error("Auto signature error:", err);
      let marker = "EVENT OK, GRAPH FAIL";
      if (err && err.stage === "token") {
        marker = "EVENT OK, TOKEN FAIL";
      }
      if (err && err.stage === "token-timeout") {
        marker = "EVENT OK, TOKEN TIMEOUT";
      }
      if (err && err.stage === "graph-timeout") {
        marker = "EVENT OK, GRAPH TIMEOUT";
      }
      if (err && err.stage === "graph-aud") {
        marker = "EVENT OK, TOKEN OK, AUD FAIL";
      }
      setMarkerSignature(marker)
        .catch((fallbackErr) => {
          console.error("Auto fallback failed:", fallbackErr);
        })
        .then(finishEvent, finishEvent);
    });
}
// function onNewMessageComposeHandler(event) {
//   Office.context.mailbox.item.body.setSignatureAsync(
//     "<p><b>TEST EVENT OK 123</b><br/>Lives International</p>",
//     { coercionType: Office.CoercionType.Html },
//     function () {
//       event.completed();
//     }
//   );
// }

// ---- Handler: buton manual din ribbon, pentru fallback ----
function insertSignatureManual(event) {
  insertSignature(true)
    .then(() => event.completed())
    .catch((err) => {
      console.error("Eroare la inserarea manuala a semnaturii:", err);
      event.completed();
    });
}

// ---- Functia principala ----
async function insertSignature(isInteractiveAuth) {
  const token = await getGraphToken(isInteractiveAuth);
  const tokenInfo = getTokenInfo(token);
  if (!tokenInfo.isGraphAudience) {
    const err = new Error(`Token audience is not Graph: ${tokenInfo.aud || "unknown"}`);
    err.stage = "graph-aud";
    throw err;
  }
  const user = await getUserData(token);
  const html = buildSignatureHtml(user);
  return setSignatureHtml(html);
}

// ---- Obtine token SSO pentru Graph API ----
// function getGraphToken() {
//   return new Promise((resolve, reject) => {
//     Office.context.auth.getAccessTokenAsync(
//       { allowSignInPrompt: true, allowConsentPrompt: true },
//       (result) => {
//         if (result.status === Office.AsyncResultStatus.Succeeded) {
//           resolve(result.value);
//         } else {
//           reject(result.error);
//         }
//       }
//     );
//   });
// }

async function getGraphToken(isInteractiveAuth) {
  const authOptions = {
    allowSignInPrompt: !!isInteractiveAuth,
    allowConsentPrompt: !!isInteractiveAuth
  };

  try {
    if (typeof OfficeRuntime !== "undefined" && OfficeRuntime.auth && OfficeRuntime.auth.getAccessToken) {
      try {
        return await withTimeout(
          OfficeRuntime.auth.getAccessToken(authOptions),
          TOKEN_TIMEOUT_MS,
          "token-timeout"
        );
      } catch (runtimeErr) {
        // Classic can hang or fail in OfficeRuntime path; fallback to legacy Office.context.auth.
      }
    }

    return await withTimeout(new Promise((resolve, reject) => {
      Office.context.auth.getAccessTokenAsync(authOptions, (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve(result.value);
        } else {
          reject(result.error);
        }
      });
    }), TOKEN_TIMEOUT_MS, "token-timeout");
  } catch (e) {
    const tokenError = e instanceof Error ? e : new Error(String(e));
    tokenError.stage = tokenError.stage || "token";
    throw tokenError;
  }
}

// ---- Interogheaza Graph API pentru datele userului curent ----
async function getUserData(token) {
  return getUserDataWithXhr(token);
}

function getUserDataWithXhr(token) {
  return new Promise((resolve, reject) => {
    const xhr = new XMLHttpRequest();
    xhr.open("GET", GRAPH_ENDPOINT, true);
    xhr.timeout = GRAPH_TIMEOUT_MS;
    xhr.setRequestHeader("Authorization", `Bearer ${token}`);

    xhr.onreadystatechange = function () {
      if (xhr.readyState !== 4) {
        return;
      }
      if (xhr.status >= 200 && xhr.status < 300) {
        try {
          resolve(JSON.parse(xhr.responseText));
        } catch (parseErr) {
          const err = new Error("Graph response parse error");
          err.stage = "graph";
          reject(err);
        }
        return;
      }
      const err = new Error(`Graph API error: ${xhr.status}`);
      err.stage = "graph";
      reject(err);
    };

    xhr.ontimeout = function () {
      const err = new Error("Graph request timeout");
      err.stage = "graph-timeout";
      reject(err);
    };

    xhr.onerror = function () {
      const err = new Error("Graph network error");
      err.stage = "graph";
      reject(err);
    };

    xhr.send();
  });
}

function withTimeout(promise, timeoutMs, stage) {
  return new Promise((resolve, reject) => {
    const id = setTimeout(() => {
      const err = new Error("Operation timeout");
      err.stage = stage;
      reject(err);
    }, timeoutMs);

    promise
      .then((value) => {
        clearTimeout(id);
        resolve(value);
      })
      .catch((err) => {
        clearTimeout(id);
        reject(err);
      });
  });
}

function getTokenInfo(token) {
  try {
    const parts = (token || "").split(".");
    if (parts.length < 2) {
      return { aud: null, isGraphAudience: false };
    }
    const payloadRaw = parts[1].replace(/-/g, "+").replace(/_/g, "/");
    const payloadJson = JSON.parse(atob(payloadRaw));
    const aud = payloadJson.aud || null;
    const isGraphAudience = aud === "https://graph.microsoft.com" || aud === "00000003-0000-0000-c000-000000000000";
    return { aud, isGraphAudience };
  } catch (err) {
    return { aud: null, isGraphAudience: false };
  }
}

function createEventFinisher(event) {
  let completed = false;
  return function finishEvent() {
    if (completed) {
      return;
    }
    completed = true;
    try {
      event.completed();
    } catch (err) {
      console.error("event.completed() failed:", err);
    }
  };
}

function setSignatureHtml(html) {
  return new Promise((resolve, reject) => {
    Office.context.mailbox.item.body.setSignatureAsync(
      html,
      { coercionType: Office.CoercionType.Html },
      (result) => {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
          resolve();
        } else {
          reject(result.error);
        }
      }
    );
  });
}

function setMarkerSignature(markerText) {
  return setSignatureHtml(`<p><b>${markerText}</b><br/>Lives International</p>`);
}

// ---- Construieste HTML-ul semnaturii, cu logica conditionala (client-side!) ----
function buildSignatureHtml(user) {
  const displayName = user.displayName || "";
  const title = user.jobTitle || "";
  const mobile = user.mobilePhone || "";
  const email = user.mail || user.userPrincipalName || "";

  // Linia de telefon apare DOAR daca exista numarul - imposibil de facut in transport rules!
  const mobileLine = mobile
    ? `<p style='margin:3px 0;font-size:9pt;color:#404040;font-family:Calibri,sans-serif'>m: <a href='tel:${mobile}' style='color:#184479;text-decoration:none'>${mobile}</a></p>`
    : "";

  const titleLine = title
    ? `<p style='margin:0;font-size:10pt;font-weight:bold;color:#404040;font-family:Calibri,sans-serif'>${title}</p>`
    : "";

  return `<div style='max-width:520px;font-family:Calibri,sans-serif'>
<table cellpadding='0' cellspacing='0' border='0' style='border-collapse:collapse'>
<tr>
<td style='padding-right:18px;vertical-align:top;width:160px'>
<a href='${COMPANY_WEBSITE}' target='_blank'>
<img src='${LOGO_URL}' width='150' height='78' alt='Lives International' style='display:block;border:0'/>
</a>
</td>
<td style='padding-left:18px;border-left:2px solid #184479;vertical-align:top'>
<p style='margin:0;font-size:13pt;color:#184479;font-family:Calibri,sans-serif'>${displayName}</p>
${titleLine}
<p style='margin:0;font-size:9pt;color:#404040;font-family:Calibri,sans-serif'>Lives International</p>
${mobileLine}
<p style='margin:0;font-size:9pt;color:#404040;font-family:Calibri,sans-serif'>e: <a href='mailto:${email}' style='color:#184479;text-decoration:none'>${email}</a></p>
<p style='margin:6px 0 0 0'><a href='${LINKEDIN_COMPANY_URL}' target='_blank'><img src='${LINKEDIN_ICON_URL}' width='15' height='15' alt='LinkedIn' style='border:0;vertical-align:middle'/></a></p>
</td>
</tr>
</table>
</div>`;
}

// Inregistram functiile pentru Office (event-based + ribbon button)
Office.actions.associate("onNewMessageComposeHandler", onNewMessageComposeHandler);
Office.actions.associate("insertSignatureManual", insertSignatureManual);
