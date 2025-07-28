/* global Excel, Office, document, window */

type Vendor = {
  name: string;
  paymentType: string;
  account: number;
};

const VENDOR_KEY = "vendors";

// Initialize
Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const isLoggedIn = localStorage.getItem("isLoggedIn") === "true";
    if (!isLoggedIn) {
      window.location.href = "login.html";
      return;
    }

    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "flex";

    (document.getElementById("logoutBtn") as HTMLButtonElement).onclick = logout;
    (document.getElementById("addVendorBtn") as HTMLButtonElement).onclick = addVendor;

    loadVendors();
  }
});

function logout() {
  localStorage.removeItem("isLoggedIn");
  window.location.href = "login.html";
}

function addVendor() {
  const name = (document.getElementById("vendorName") as HTMLInputElement).value.trim();
  const paymentType = (document.getElementById("paymentType") as HTMLSelectElement).value;
  const account = parseInt((document.getElementById("account") as HTMLSelectElement).value);

  if (!name) {
    alert("Vendor name is required");
    return;
  }

  const vendor: Vendor = { name, paymentType, account };
  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");
  vendors.push(vendor);
  localStorage.setItem(VENDOR_KEY, JSON.stringify(vendors));

  (document.getElementById("vendorName") as HTMLInputElement).value = "";
  loadVendors();
}

function loadVendors() {
  const vendorList = document.getElementById("vendorList")!;
  vendorList.innerHTML = "";

  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");

  vendors.forEach((vendor, index) => {
    const li = document.createElement("li");
    li.className = "ms-ListItem";
    li.innerHTML = `
      ${vendor.name} - ${vendor.paymentType} - Account ${vendor.account}
      <button class="ms-Button ms-Button--small" data-index="${index}">
        <span class="ms-Button-label">Delete</span>
      </button>
    `;
    li.querySelector("button")!.onclick = () => deleteVendor(index);
    vendorList.appendChild(li);
  });
}

function deleteVendor(index: number) {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");
  vendors.splice(index, 1);
  localStorage.setItem(VENDOR_KEY, JSON.stringify(vendors));
  loadVendors();
}
