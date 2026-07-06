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
  insertSignature()
    .then(() => event.completed())
    .catch((err) => {
      console.error("Auto signature error:", err);
      // fallback vizibil ca sa stii ca eventul ruleaza
      Office.context.mailbox.item.body.setSignatureAsync(
        "<p><b>EVENT OK, GRAPH FAIL</b></p>",
        { coercionType: Office.CoercionType.Html },
        () => event.completed()
      );
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
  insertSignature()
    .then(() => event.completed())
    .catch((err) => {
      console.error("Eroare la inserarea manuala a semnaturii:", err);
      event.completed();
    });
}

// ---- Functia principala ----
async function insertSignature() {
  const token = await getGraphToken();
  const user = await getUserData(token);
  const html = buildSignatureHtml(user);

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

async function getGraphToken() {
  try {
    // In multe build-uri Outlook event-based, asta e calea corecta
    return await OfficeRuntime.auth.getAccessToken({
      allowSignInPrompt: true,
      allowConsentPrompt: true
    });
  } catch (e) {
    throw e;
  }
}

// ---- Interogheaza Graph API pentru datele userului curent ----
async function getUserData(token) {
  const response = await fetch(GRAPH_ENDPOINT, {
    headers: { Authorization: `Bearer ${token}` },
  });
  if (!response.ok) {
    throw new Error(`Graph API error: ${response.status}`);
  }
  return response.json();
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
