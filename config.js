// EAM Dispatch Board settings. Only this file changes between setups.
window.EAM_CONFIG = {
  // Application (client) ID from the Microsoft Entra app registration.
  clientId: "f274a355-d7ab-4c65-ad48-9b047e51685e",
  // Directory (tenant) ID from the same page. "organizations" also works.
  tenant: "4d3d8101-8c25-4a28-a975-edabfa5e83cd",
  // Leave empty when everyone signs in as the admin Outlook account.
  // If staff sign in with their own accounts instead, put the admin address here
  // (the calendar must then be shared with them with edit rights).
  mailbox: ""
};
