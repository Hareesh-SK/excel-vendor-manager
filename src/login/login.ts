document.addEventListener("DOMContentLoaded", () => {
  const loginBtn = document.getElementById("loginBtn") as HTMLButtonElement;
  const usernameInput = document.getElementById("username") as HTMLInputElement;
  const passwordInput = document.getElementById("password") as HTMLInputElement;
  const errorMsg = document.getElementById("login-error") as HTMLParagraphElement;

  loginBtn.onclick = () => {
    const username = usernameInput.value;
    const password = passwordInput.value;

    if (username === "admin" && password === "admin123") {
      localStorage.setItem("isLoggedIn", "true");
      window.location.href = "taskpane.html";
    } else {
      errorMsg.innerText = "Invalid credentials!";
    }
  };
});
