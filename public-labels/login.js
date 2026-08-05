const form = document.getElementById("loginForm");
const username = document.getElementById("username");
const password = document.getElementById("password");
const loginButton = document.getElementById("loginButton");
const loginError = document.getElementById("loginError");
const togglePassword = document.getElementById("togglePassword");

function destination() {
  const requested =
    new URLSearchParams(window.location.search).get("next") || "/dashboard";
  return requested.startsWith("/") &&
    !requested.startsWith("//") &&
    !requested.startsWith("/login")
    ? requested
    : "/dashboard";
}

function showError(message) {
  loginError.textContent = message;
  loginError.hidden = false;
}

togglePassword.addEventListener("click", () => {
  const visible = password.type === "text";
  password.type = visible ? "password" : "text";
  togglePassword.textContent = visible ? "Show" : "Hide";
  togglePassword.setAttribute(
    "aria-label",
    visible ? "Show password" : "Hide password",
  );
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  loginError.hidden = true;
  if (!username.value.trim() || !password.value) {
    showError("Enter both your username and password.");
    return;
  }

  loginButton.disabled = true;
  loginButton.textContent = "Signing in...";
  try {
    const response = await fetch("/api/dashboard-login", {
      method: "POST",
      headers: { "Content-Type": "application/json" },
      body: JSON.stringify({
        username: username.value.trim(),
        password: password.value,
      }),
    });
    const result = await response.json().catch(() => ({}));
    if (!response.ok) throw new Error(result.error || "Unable to sign in.");
    window.location.replace(destination());
  } catch (error) {
    showError(error.message);
    loginButton.disabled = false;
    loginButton.textContent = "Sign in";
  }
});

fetch("/api/dashboard-session")
  .then((response) => response.json())
  .then((session) => {
    if (session.authenticated) window.location.replace(destination());
  })
  .catch(() => {});
