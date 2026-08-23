const GOOGLE_SCRIPT_URL = "https://script.google.com/macros/s/AKfycby2co1jbRf9qqbhMcQ15ARL13CQRAmJga2-Af5eMq5wm5wXN6iiRpHzMJCQUYaq8fuO/exec";

const form = document.getElementById("tshirtForm");
const submitBtn = document.getElementById("submitBtn");
const popup = document.getElementById("successPopup");
const popupClose = document.getElementById("popupClose");
const previewNumber = document.getElementById("previewNumber");
const previewName = document.getElementById("previewName");
const jerseyPreview = document.getElementById("jerseyPreview");

const fields = {
  fullName: document.getElementById("fullName"),
  phone: document.getElementById("phone"),
  size: document.getElementById("size"),
  sleeveLength: document.getElementById("sleeveLength"),
  backNumber: document.getElementById("backNumber"),
  backName: document.getElementById("backName"),
};

const sleeveClassMap = {
  Full: "sleeve-full",
  Half: "sleeve-half",
  "Three-Quarter Sleeve": "sleeve-three-quarter",
};

function setError(input, message) {
  const field = input.closest(".field");
  const errorEl = field.querySelector(".error");
  field.classList.toggle("invalid", Boolean(message));
  errorEl.textContent = message || "";
}

function validateName(value) {
  if (!value.trim()) return "Please enter your full name.";
  if (value.trim().length < 2) return "Please enter a valid full name.";
  return "";
}

function validatePhone(value) {
  if (!/^[6-9]\d{9}$/.test(value)) {
    return "Please enter a valid 10-digit mobile number.";
  }
  return "";
}

function validateSelect(value, label) {
  return value ? "" : `Please select a ${label}.`;
}

function validateBackNumber(value) {
  if (value === "") return "Please enter a back number.";
  if (!/^\d{1,3}$/.test(value)) {
    return "Back number may contain digits only.";
  }
  const num = Number(value);
  if (num > 99) {
    return "Back number must be between 0 and 99.";
  }
  return "";
}

function validateBackName(value) {
  if (!value.trim()) return "Please enter a back name.";
  if (!/^[A-Za-z ]+$/.test(value.trim())) {
    return "Back name may contain letters only.";
  }
  return "";
}

function validateForm() {
  const errors = {
    fullName: validateName(fields.fullName.value),
    phone: validatePhone(fields.phone.value),
    size: validateSelect(fields.size.value, "size"),
    sleeveLength: validateSelect(fields.sleeveLength.value, "sleeve length"),
    backNumber: validateBackNumber(fields.backNumber.value),
    backName: validateBackName(fields.backName.value),
  };

  Object.keys(errors).forEach((key) => setError(fields[key], errors[key]));
  return !Object.values(errors).some(Boolean);
}

function updatePreview() {
  const number = fields.backNumber.value;
  const name = fields.backName.value.trim().toUpperCase();
  previewNumber.textContent = number === "" ? "00" : number;
  previewName.textContent = name || "NAME";

  jerseyPreview.classList.remove("sleeve-full", "sleeve-half", "sleeve-three-quarter");
  const sleeveClass = sleeveClassMap[fields.sleeveLength.value];
  if (sleeveClass) {
    jerseyPreview.classList.add(sleeveClass);
  }
}

function showPopup() {
  popup.hidden = false;
  popupClose.focus();
}

function hidePopup() {
  popup.hidden = true;
}

function setLoading(isLoading) {
  submitBtn.disabled = isLoading;
  submitBtn.classList.toggle("loading", isLoading);
}

function getFormPayload() {
  return {
    fullName: fields.fullName.value.trim(),
    phone: fields.phone.value.trim(),
    size: fields.size.value,
    sleeveLength: fields.sleeveLength.value,
    backNumber: String(fields.backNumber.value),
    backName: fields.backName.value.trim().toUpperCase(),
  };
}

async function saveToGoogleSheet(payload) {
  if (!GOOGLE_SCRIPT_URL || GOOGLE_SCRIPT_URL.includes("PASTE_YOUR_GOOGLE")) {
    console.warn("Google Apps Script URL is not set yet. Showing the success message locally.");
    return;
  }

  const body = new FormData();
  Object.entries(payload).forEach(([key, value]) => body.append(key, value));

  await fetch(GOOGLE_SCRIPT_URL, {
    method: "POST",
    mode: "no-cors",
    body,
  });
}

fields.backNumber.addEventListener("input", () => {
  fields.backNumber.value = fields.backNumber.value.replace(/\D/g, "").slice(0, 3);
  updatePreview();
});
fields.sleeveLength.addEventListener("change", updatePreview);
fields.backName.addEventListener("input", () => {
  fields.backName.value = fields.backName.value.toUpperCase();
  updatePreview();
});

fields.phone.addEventListener("input", () => {
  fields.phone.value = fields.phone.value.replace(/\D/g, "").slice(0, 10);
});

Object.values(fields).forEach((input) => {
  input.addEventListener("blur", () => {
    if (input === fields.fullName) setError(input, validateName(input.value));
    if (input === fields.phone) setError(input, validatePhone(input.value));
    if (input === fields.size) setError(input, validateSelect(input.value, "size"));
    if (input === fields.sleeveLength) setError(input, validateSelect(input.value, "sleeve length"));
    if (input === fields.backNumber) setError(input, validateBackNumber(input.value));
    if (input === fields.backName) setError(input, validateBackName(input.value));
  });
});

form.addEventListener("submit", async (event) => {
  event.preventDefault();
  if (!validateForm()) return;

  setLoading(true);
  try {
    await saveToGoogleSheet(getFormPayload());
    form.reset();
    updatePreview();
    showPopup();
  } catch (error) {
    console.error(error);
    alert("Unable to submit. Please check your internet connection and try again.");
  } finally {
    setLoading(false);
  }
});

popupClose.addEventListener("click", hidePopup);
popup.addEventListener("click", (event) => {
  if (event.target === popup) hidePopup();
});

document.addEventListener("keydown", (event) => {
  if (event.key === "Escape" && !popup.hidden) hidePopup();
});

updatePreview();
