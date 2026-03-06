(() => {
  const SUPABASE_URL = "https://bsceqirconhqmxwipbyl.supabase.co";
  const SUPABASE_ANON_KEY = "sb_publishable_zE9Cz_GnZIRluKPkr41RxA_EqCZxVgp";

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
    const hash = window.location.hash.startsWith("#") ? window.location.hash.slice(1) : "";
    return new URLSearchParams(hash);
  };

  const bootstrapSession = async () => {
    const params = getHashParams();
    const type = params.get("type");
    const accessToken = params.get("access_token");
    const refreshToken = params.get("refresh_token");

    if (type !== "recovery" || !accessToken || !refreshToken) {
      setStatus("Invalid or expired reset link.", "error");
      if (resetForm) resetForm.classList.add("hidden");
      return false;
    }

    let error = null;
    try {
      const result = await supabase.auth.setSession({
        access_token: accessToken,
        refresh_token: refreshToken,
      });
      error = result.error;
    } catch (err) {
      if (String(err?.name || "") === "AbortError") {
        await new Promise((resolve) => setTimeout(resolve, 300));
        try {
          const retry = await supabase.auth.setSession({
            access_token: accessToken,
            refresh_token: refreshToken,
          });
          error = retry.error;
        } catch (retryErr) {
          error = retryErr;
        }
      } else {
        error = err;
      }
    }

    if (error) {
      setStatus("Invalid or expired reset link.", "error");
      if (resetForm) resetForm.classList.add("hidden");
      return false;
    }

    if (window.history && window.history.replaceState) {
      window.history.replaceState(null, "", window.location.pathname);
    }

    setStatus("Set a new password to complete the reset.");
    return true;
  };

  const handleSubmit = async (event) => {
    event.preventDefault();
    const password = newPassword ? newPassword.value.trim() : "";
    const confirm = confirmPassword ? confirmPassword.value.trim() : "";

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
