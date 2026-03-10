(() => {
  const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

  // Snapshot URL params before Supabase init (some clients may scrub/alter the URL).
  const INITIAL_SEARCH = window.location.search || "";
  const INITIAL_HASH = window.location.hash || "";

  const supabase = window.supabase.createClient(SUPABASE_URL, SUPABASE_ANON_KEY, {
    auth: {
      detectSessionInUrl: false,
      persistSession: true,
    },
  });

  const resetStatus = document.getElementById("resetStatus");
  const resetForm = document.getElementById("resetForm");
  const newPassword = document.getElementById("newPassword");
  const confirmPassword = document.getElementById("confirmPassword");
  const resetSubmit = document.getElementById("resetSubmit");

  const setStatus = (message, tone = "") => {
    if (!resetStatus) return;
    resetStatus.textContent = message;
    resetStatus.classList.remove("ok", "error");
    if (tone) resetStatus.classList.add(tone);
  };

  const setLoading = (isLoading) => {
    if (resetSubmit) resetSubmit.disabled = isLoading;
    if (newPassword) newPassword.disabled = isLoading;
    if (confirmPassword) confirmPassword.disabled = isLoading;
  };

  const getHashParams = () => {
    const hash = INITIAL_HASH.startsWith("#") ? INITIAL_HASH.slice(1) : INITIAL_HASH;
    return new URLSearchParams(hash);
  };

  const getQueryParams = () => new URLSearchParams(INITIAL_SEARCH || "");

  const hasParam = (value) => typeof value === "string" && value.trim().length > 0;

  const scrubUrl = () => {
    if (window.history && window.history.replaceState) {
      window.history.replaceState(null, "", window.location.pathname);
    }
  };

  const ensureSession = async () => {
    const { data } = await supabase.auth.getSession();
    return data?.session || null;
  };

  const bootstrapSession = async () => {
    const hashParams = getHashParams();
    const queryParams = getQueryParams();

    const typeFromHash = hashParams.get("type");
    const typeFromQuery = queryParams.get("type");
    const type = (typeFromQuery || typeFromHash || "").trim();

    const accessToken = queryParams.get("access_token") || hashParams.get("access_token");
    const refreshToken = queryParams.get("refresh_token") || hashParams.get("refresh_token");

    const code = queryParams.get("code") || hashParams.get("code");
    const tokenHash = queryParams.get("token_hash") || hashParams.get("token_hash");
    const token = queryParams.get("token") || hashParams.get("token");

    if (!resetForm) return false;

    setStatus("Checking reset link…");
    setLoading(true);
    let error = null;

    try {
      const existingSession = await ensureSession();
      if (existingSession) {
        scrubUrl();
        setStatus("Set a new password to complete the reset.");
        resetForm.classList.remove("hidden");
        return true;
      }

      // Prefer the official parser when available (handles both fragments and query params).
      try {
        const parsed = await supabase.auth.getSessionFromUrl({ storeSession: true });
        if (parsed?.data?.session) {
          scrubUrl();
          setStatus("Set a new password to complete the reset.");
          resetForm.classList.remove("hidden");
          return true;
        }
        if (parsed?.error) error = parsed.error;
      } catch {
        // Ignore and fall back to manual handlers below.
      }

      if (hasParam(code)) {
        const result = await supabase.auth.exchangeCodeForSession(code);
        error = result.error;
      } else if ((typeFromHash === "recovery" || type === "recovery") && hasParam(accessToken)) {
        const result = await supabase.auth.setSession({
          access_token: accessToken,
          refresh_token: refreshToken || "",
        });
        error = result.error;
      } else if ((hasParam(tokenHash) || hasParam(token)) && hasParam(type)) {
        const result = await supabase.auth.verifyOtp({
          ...(hasParam(tokenHash) ? { token_hash: tokenHash } : { token }),
          type,
        });
        error = result.error;
      } else {
        error = new Error("Missing recovery parameters.");
      }
    } catch (err) {
      if (String(err?.name || "") === "AbortError") {
        await new Promise((resolve) => setTimeout(resolve, 300));
        try {
          if (hasParam(code)) {
            const retry = await supabase.auth.exchangeCodeForSession(code);
            error = retry.error;
          } else if ((typeFromHash === "recovery" || type === "recovery") && hasParam(accessToken)) {
            const retry = await supabase.auth.setSession({
              access_token: accessToken,
              refresh_token: refreshToken || "",
            });
            error = retry.error;
          } else if ((hasParam(tokenHash) || hasParam(token)) && hasParam(type)) {
            const retry = await supabase.auth.verifyOtp({
              ...(hasParam(tokenHash) ? { token_hash: tokenHash } : { token }),
              type,
            });
            error = retry.error;
          } else {
            error = new Error("Missing recovery parameters.");
          }
        } catch (retryErr) {
          error = retryErr;
        }
      } else {
        error = err;
      }
    } finally {
      setLoading(false);
    }

    const session = await ensureSession();
    if (error || !session) {
      const debug = `Params: code=${hasParam(code) ? "yes" : "no"}, access_token=${
        hasParam(accessToken) ? "yes" : "no"
      }, refresh_token=${hasParam(refreshToken) ? "yes" : "no"}, token_hash=${hasParam(tokenHash) ? "yes" : "no"}, token=${
        hasParam(token) ? "yes" : "no"
      }, type=${type || "(none)"}`;
      const detail = error?.message ? ` (${error.message})` : "";
      setStatus(`Invalid or expired reset link.${detail} ${debug}`, "error");
      resetForm.classList.remove("hidden");
      return false;
    }

    scrubUrl();
    setStatus("Set a new password to complete the reset.");
    resetForm.classList.remove("hidden");
    return true;
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    const password = newPassword ? newPassword.value.trim() : "";
    const confirm = confirmPassword ? confirmPassword.value.trim() : "";

    let session = await ensureSession();
    if (!session) {
      await bootstrapSession();
      session = await ensureSession();
      if (!session) {
        setStatus("Reset link not verified. Please open the reset link again or request a new one.", "error");
        return;
      }
    }

    if (!password) {
      setStatus("Please enter a new password.", "error");
      return;
    }
    if (password !== confirm) {
      setStatus("Passwords do not match.", "error");
      return;
    }

    setLoading(true);
    const { error } = await supabase.auth.updateUser({ password });
    setLoading(false);

    if (error) {
      setStatus(`Reset failed: ${error.message}`, "error");
      return;
    }

    setStatus("Password updated. Redirecting...", "ok");
    window.setTimeout(() => {
      window.location.href = "./";
    }, 1200);
  };

  if (resetForm) resetForm.addEventListener("submit", handleSubmit);
  bootstrapSession();
})();
