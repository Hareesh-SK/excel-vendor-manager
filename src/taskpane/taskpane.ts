/* global Excel, Office, document, window */

// Make sure Office is ready before accessing the DOM or Excel APIs
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const isLoggedIn = localStorage.getItem("isLoggedIn") === "true";
    if (!isLoggedIn) {
      window.location.href = "login.html";
      return;
    }

    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Assign functions to window so they're available in HTML onclick
    (window as any).run = run;
    (window as any).logout = logout;

    document.getElementById("run").onclick = run;
    document.getElementById("logoutBtn").onclick = logout;
  }
});

// Excel logic when "Run" button is clicked
async function run() {
  try {
    await Excel.run(async (context) => {
      const range = context.workbook.getSelectedRange();

      // Load address and modify the fill color
      range.load("address");
      range.format.fill.color = "yellow";

      await context.sync();
      console.log(`The range address was ${range.address}.`);
    });
  } catch (error) {
    console.error(error);
  }
}

// Logout function
function logout() {
  localStorage.removeItem("isLoggedIn");
  window.location.href = "login.html";
}
