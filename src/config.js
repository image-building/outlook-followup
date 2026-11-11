/* Filled in by `node scripts/configure.js <baseUrl> <clientId> <tenantId>` (see README). */
window.FU_CONFIG = {
  clientId: "YOUR-ENTRA-APP-CLIENT-ID",
  baseUrl: "https://YOUR-HOST/outlook-followup",
  authority: "https://login.microsoftonline.com/organizations",
  scopes: ["User.Read", "Mail.ReadWrite", "Calendars.ReadWrite", "MailboxSettings.ReadWrite"],
  category: "Follow-up",
  defaults: { days: 14, reminderTime: "09:00" }
};
