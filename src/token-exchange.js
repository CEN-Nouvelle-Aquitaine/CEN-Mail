(async () => {
  const params = new URLSearchParams(location.search);
  const tokenUrl = params.get("tokenUrl");

  const form = document.getElementById("tokenform");
  form.action = tokenUrl;

  const fields = {
    client_id: params.get("client_id"),
    code: params.get("code"),
    redirect_uri: params.get("redirect_uri"),
    grant_type: "authorization_code",
    code_verifier: params.get("code_verifier"),
    scope: params.get("scope"),
  };

  for (const [key, value] of Object.entries(fields)) {
    const input = document.createElement("input");
    input.type = "hidden";
    input.name = key;
    input.value = value;
    form.appendChild(input);
  }

  // Soumettre le formulaire dans l'iframe
  form.submit();

  // Attendre que l'iframe charge la réponse JSON
  const iframe = document.getElementById("tokenframe");
  iframe.addEventListener("load", async () => {
    try {
      const text = iframe.contentDocument.body.textContent || iframe.contentDocument.body.innerText || "";
      console.log("[TokenExchange] Response:", text.substring(0, 100));
      const data = JSON.parse(text);
      await browser.storage.local.set({ _graphTokenResult: { data, ts: Date.now() } });
    } catch(e) {
      console.error("[TokenExchange] Error:", e.message);
      await browser.storage.local.set({ _graphTokenResult: { error: e.message, ts: Date.now() } });
    }

    // Fermer cet onglet
    try {
      const tab = await browser.tabs.getCurrent();
      if (tab) browser.tabs.remove(tab.id);
    } catch(e) {}
  });
})();
