/* global Excel, Office, document, window */

type Vendor = {
  name: string;
  paymentType: string;
  account: number;
  skipNextPayment?: boolean;
  lastPaid?: string;
};

type Payment = {
  vendorName: string;
  amount: number;
  account: number;
  date: string;
};

type PendingPayment = {
  vendorName: string;
  amount: number;
  account: number;
  date: string;
};

const VENDOR_KEY = "vendors";
const BALANCE_KEY = "accountBalances";
const PAYMENT_HISTORY_KEY = "paymentHistory";
const PENDING_PAYMENTS_KEY = "pendingPayments";
const BASE_PAYMENT = 100;

let accountBalances = {
  1: 200000,
  2: 200000,
};

let selectedVendorName = "";

Office.onReady((info) => {
  if (info.host === Office.HostType.Excel) {
    const isLoggedIn = localStorage.getItem("isLoggedIn") === "true";
    if (!isLoggedIn) {
      window.location.href = "login.html";
      return;
    }

    document.getElementById("sideload-msg")!.style.display = "none";
    document.getElementById("app-body")!.style.display = "block";

    (document.getElementById("logoutBtn") as HTMLButtonElement).onclick = logout;
    (document.getElementById("addVendorBtn") as HTMLButtonElement).onclick = addVendor;
    (document.getElementById("exportBtn") as HTMLButtonElement).onclick = generateStatement;
    (document.getElementById("runScheduledPaymentsBtn") as HTMLButtonElement).onclick = runScheduledPayments;
    (document.getElementById("payNowBtn") as HTMLButtonElement).onclick = triggerOnDemandPayment;

    loadVendors();
    loadBalances();
    displayPendingPayments();
    loadPaymentHistory();
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

  const vendor: Vendor = { name, paymentType, account, skipNextPayment: false };
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
      <button class="ms-Button ms-Button--small delete-btn">
        <span class="ms-Button-label">Delete</span>
      </button>
      <button class="ms-Button ms-Button--small edit-btn">Edit</button>
    `;

    li.querySelector(".delete-btn")!.addEventListener("click", () => deleteVendor(index));
    li.querySelector(".edit-btn")!.addEventListener("click", () => editVendor(vendor.name));

    vendorList.appendChild(li);
  });
}

function deleteVendor(index: number) {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");
  vendors.splice(index, 1);
  localStorage.setItem(VENDOR_KEY, JSON.stringify(vendors));
  loadVendors();
}

let vendorToEditIndex = -1;

function editVendor(name: string) {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");
  const index = vendors.findIndex(v => v.name === name);
  if (index === -1) return alert("Vendor not found");

  const vendor = vendors[index];
  vendorToEditIndex = index;

  // Set values into the form
  (document.getElementById("vendorNameInput") as HTMLInputElement).value = vendor.name;
  (document.getElementById("paymentTypeInput") as HTMLInputElement).value = vendor.paymentType;
  (document.getElementById("accountInput") as HTMLInputElement).value = vendor.account.toString();

  document.getElementById("editVendorModal")!.style.display = "block";
}


function saveEditedVendor() {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");

  const name = (document.getElementById("vendorNameInput") as HTMLInputElement).value;
  const paymentType = (document.getElementById("paymentTypeInput") as HTMLInputElement).value;
  const account = parseInt((document.getElementById("accountInput") as HTMLInputElement).value || "0");

  if (!name || !paymentType || !account) {
    alert("Please fill all fields");
    return;
  }

  const vendor = vendors[vendorToEditIndex];
  vendor.name = name;
  vendor.paymentType = paymentType;
  vendor.account = account;

  localStorage.setItem(VENDOR_KEY, JSON.stringify(vendors));
  loadVendors();

  document.getElementById("editVendorModal")!.style.display = "none";
}


function closeEditModal() {
  (document.getElementById("editVendorModal") as HTMLDivElement).style.display = "none";
}

(window as any).editVendor = editVendor;
(window as any).saveEditedVendor = saveEditedVendor;
(window as any).closeEditModal = closeEditModal;


function loadBalances() {
  const saved = localStorage.getItem(BALANCE_KEY);
  if (saved) {
    accountBalances = JSON.parse(saved);
  }
  updateBalanceUI();
}

function updateBalanceUI() {
  document.getElementById("account1-balance")!.textContent = `$${accountBalances[1].toLocaleString()}`;
  document.getElementById("account2-balance")!.textContent = `$${accountBalances[2].toLocaleString()}`;
}

function deductFromAccount(account: number, amount: number): boolean {
  if (accountBalances[account] >= amount) {
    accountBalances[account] -= amount;
    localStorage.setItem(BALANCE_KEY, JSON.stringify(accountBalances));
    updateBalanceUI();
    return true;
  }
  return false;
}

function logPayment(vendorName: string, amount: number, account: number) {
  const history: Payment[] = JSON.parse(localStorage.getItem(PAYMENT_HISTORY_KEY) || "[]");
  history.push({ vendorName, amount, account, date: new Date().toISOString() });
  localStorage.setItem(PAYMENT_HISTORY_KEY, JSON.stringify(history));
  loadPaymentHistory();
}

function loadPaymentHistory() {
  const list = document.getElementById("paymentHistoryList")!;
  list.innerHTML = "";
  const history: Payment[] = JSON.parse(localStorage.getItem(PAYMENT_HISTORY_KEY) || "[]");
  history.slice().reverse().forEach(p => {
    const li = document.createElement("li");
    li.textContent = `${p.vendorName} - $${p.amount} - Account ${p.account} - ${new Date(p.date).toLocaleString()}`;
    list.appendChild(li);
  });
}

function isAlternateFriday(lastPaid?: string): boolean {
  const today = new Date();
  if (today.getDay() !== 5) return false;
  if (!lastPaid) return true;
  const last = new Date(lastPaid);
  const diffDays = Math.floor((today.getTime() - last.getTime()) / (1000 * 60 * 60 * 24));
  return diffDays >= 14;
}

function runScheduledPayments() {
  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");
  const today = new Date();
  const pending: PendingPayment[] = JSON.parse(localStorage.getItem(PENDING_PAYMENTS_KEY) || "[]");

  vendors.forEach(vendor => {
    if ((vendor.paymentType === "Weekly" && today.getDay() === 5) ||
        (vendor.paymentType === "Biweekly" && isAlternateFriday(vendor.lastPaid))) {
      if (vendor.skipNextPayment) {
        vendor.skipNextPayment = false;
        return;
      }

      const success = deductFromAccount(vendor.account, BASE_PAYMENT);
      if (success) {
        logPayment(vendor.name, BASE_PAYMENT, vendor.account);
        vendor.lastPaid = new Date().toISOString();
      } else {
        pending.push({ vendorName: vendor.name, amount: BASE_PAYMENT, account: vendor.account, date: today.toISOString() });
      }
    }
  });

  localStorage.setItem(VENDOR_KEY, JSON.stringify(vendors));
  localStorage.setItem(PENDING_PAYMENTS_KEY, JSON.stringify(pending));
  displayPendingPayments();
}

function triggerOnDemandPayment() {

  const vendors: Vendor[] = JSON.parse(localStorage.getItem(VENDOR_KEY) || "[]");
  const onDemandVendors = vendors.filter(v => v.paymentType === "On-Demand");
  const errorDiv = document.getElementById("errorMessage");
  if (onDemandVendors.length === 0) {
       if (errorDiv) {
      errorDiv.textContent = "No On-Demand vendors found.";
    }
    return;
  }
 
    if (errorDiv) {
    errorDiv.textContent = "";
  }
  const confirmPay = confirm("Are you sure you want to pay all On-Demand vendors now?");
  if (!confirmPay) return;

  onDemandVendors.forEach(vendor => {
    const success = deductFromAccount(vendor.account, BASE_PAYMENT);
    if (success) {
      logPayment(vendor.name, BASE_PAYMENT, vendor.account);
      vendor.skipNextPayment = true;
    } else {
      const pending: PendingPayment[] = JSON.parse(localStorage.getItem(PENDING_PAYMENTS_KEY) || "[]");
      pending.push({ vendorName: vendor.name, amount: BASE_PAYMENT, account: vendor.account, date: new Date().toISOString() });
      localStorage.setItem(PENDING_PAYMENTS_KEY, JSON.stringify(pending));
    }
  });

  localStorage.setItem(VENDOR_KEY, JSON.stringify(vendors));
  displayPendingPayments();
}

function displayPendingPayments() {
  const list = document.getElementById("pendingPaymentsList");
  if (!list) return;
  list.innerHTML = "";
  const pending: PendingPayment[] = JSON.parse(localStorage.getItem(PENDING_PAYMENTS_KEY) || "[]");
  pending.slice().reverse().forEach(p => {
    const li = document.createElement("li");
    li.textContent = `${p.vendorName} - $${p.amount} - Account ${p.account} - ${new Date(p.date).toLocaleString()}`;
    list.appendChild(li);
  });
}

function generateStatement() {
  const history: Payment[] = JSON.parse(localStorage.getItem(PAYMENT_HISTORY_KEY) || "[]");

  if (history.length === 0) {
    alert("No payments to generate.");
    return;
  }

  Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    sheet.getRange("A1:D1").values = [["Vendor Name", "Amount", "Account", "Date"]];
    const data = history.map(p => [p.vendorName, p.amount, p.account, new Date(p.date).toLocaleString()]);
    sheet.getRangeByIndexes(1, 0, data.length, 4).values = data;
    sheet.getRange("F1").values = [["Account 1 Balance"]];
    sheet.getRange("F2").values = [[accountBalances[1]]];
    sheet.getRange("G1").values = [["Account 2 Balance"]];
    sheet.getRange("G2").values = [[accountBalances[2]]];

    await context.sync();
    alert("Statement generated in Excel.");
  }).catch((err) => {
    console.error(err);
    alert("Error generating statement: " + err);
  });
}
