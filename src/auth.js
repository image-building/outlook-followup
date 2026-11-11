/* Microsoft Graph sign-in.
 * Preferred: Nested App Authentication (NAA) — Outlook signs you in with the account already in Outlook.
 * Fallback for older Outlook builds: an Office dialog that runs a normal MSAL redirect sign-in. */
window.FUAuth = (function () {
  "use strict";
  var cfg = window.FU_CONFIG;
  var pca = null;
  var mode = null; // 'naa' | 'dialog'
  var dialogToken = null; // { token, exp }

  async function init() {
    if (mode) return mode;
    var naaOk = Office.context.requirements && Office.context.requirements.isSetSupported("NestedAppAuth", "1.1");
    if (naaOk) {
      try {
        pca = await msal.createNestablePublicClientApplication({
          auth: { clientId: cfg.clientId, authority: cfg.authority }
        });
        mode = "naa";
        return mode;
      } catch (e) {
        console.warn("NAA init failed, using dialog sign-in", e);
      }
    }
    mode = "dialog";
    return mode;
  }

  /** interactive=false: never show UI (used for automatic background syncs). */
  async function getToken(interactive) {
    await init();
    var req = { scopes: cfg.scopes };
    if (mode === "naa") {
      try {
        return (await pca.acquireTokenSilent(req)).accessToken;
      } catch (e) {
        if (!interactive) throw needsSignIn(e);
        return (await pca.acquireTokenPopup(req)).accessToken;
      }
    }
    if (dialogToken && dialogToken.exp > Date.now() + 60000) return dialogToken.token;
    if (!interactive) throw needsSignIn();
    return await signInWithDialog();
  }

  function needsSignIn(cause) {
    var err = new Error("Sign-in required");
    err.code = "SIGNIN_REQUIRED";
    err.cause = cause;
    return err;
  }

  function signInWithDialog() {
    return new Promise(function (resolve, reject) {
      if (!Office.context.ui || !Office.context.ui.displayDialogAsync) {
        return reject(new Error("This Outlook app can't sign in to the add-in yet. Update Outlook, or use the add-in on your Mac/PC — tracking syncs across devices."));
      }
      var url = cfg.baseUrl + "/src/auth-dialog.html";
      Office.context.ui.displayDialogAsync(url, { height: 60, width: 30, promptBeforeOpen: false }, function (res) {
        if (res.status !== Office.AsyncResultStatus.Succeeded) return reject(new Error(res.error.message));
        var dialog = res.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
          dialog.close();
          try {
            var msg = JSON.parse(arg.message);
            if (msg.error) return reject(new Error(msg.error));
            dialogToken = { token: msg.token, exp: msg.exp };
            resolve(msg.token);
          } catch (e) { reject(e); }
        });
        dialog.addEventHandler(Office.EventType.DialogEventReceived, function () {
          reject(new Error("Sign-in window was closed"));
        });
      });
    });
  }

  return { init: init, getToken: getToken };
})();
