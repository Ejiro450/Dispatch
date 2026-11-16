// EAM Dispatch Board settings. Only this file changes between setups.
window.EAM_CONFIG = {
  // Application (client) ID from the Microsoft Entra app registration.
  clientId: "",
  // Directory (tenant) ID from the same page. "organizations" also works.
  tenant: "organizations",
  // Leave empty when everyone signs in as the admin Outlook account.
  // If staff sign in with their own accounts instead, put the admin address here
  // (the calendar must then be shared with them with edit rights).
  mailbox: ""
};
