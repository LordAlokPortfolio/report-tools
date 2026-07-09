// Microsoft Graph / Microsoft To Do integration — STUBBED.
//
// This scaffold does not sign in or call Graph yet. All task data lives in
// localStorage on the current device only (see taskpane.js). To turn this
// into the real cross-device version described in the product spec:
//
//   1. Register an app in Azure AD / Microsoft Entra admin center:
//      https://entra.microsoft.com -> App registrations -> New registration.
//        - Redirect URI: add the SPA redirect for this task pane's origin,
//          e.g. https://<your-host>/taskpane.html
//        - API permissions (delegated): Tasks.ReadWrite, Mail.Read
//   2. Fill in GRAPH_CONFIG.clientId and GRAPH_CONFIG.tenantId below.
//   3. Add the MSAL browser SDK to taskpane.html:
//        <script src="https://alcdn.msauth.net/browser/2.x.x/js/msal-browser.min.js"></script>
//   4. Implement signIn()/getAccessToken() using msal.PublicClientApplication.
//   5. Implement the Graph calls stubbed below.
//
// Until then, every method here resolves to a "not connected" result so the
// rest of the app can run in local-only mode without throwing.

const GRAPH_CONFIG = {
  clientId: "", // TODO: paste Azure AD app (client) ID here
  tenantId: "common",
  scopes: ["Tasks.ReadWrite", "Mail.Read"],
  todoListName: "Boss replies",
};

const GraphClient = {
  isConfigured() {
    return Boolean(GRAPH_CONFIG.clientId);
  },

  isSignedIn() {
    // TODO: return true once MSAL session/account is established.
    return false;
  },

  async signIn() {
    // TODO: wire up msal.PublicClientApplication.loginPopup({ scopes: GRAPH_CONFIG.scopes })
    console.warn("[graph] signIn() not implemented — running in local-only mode.");
    return { signedIn: false, reason: "not-configured" };
  },

  async signOut() {
    // TODO: msalInstance.logoutPopup()
  },

  async getAccessToken() {
    // TODO: msalInstance.acquireTokenSilent({ scopes: GRAPH_CONFIG.scopes, account })
    return null;
  },

  // Ensures a "Boss replies" list exists in Microsoft To Do and returns its id.
  async ensureTaskList() {
    // TODO: GET  https://graph.microsoft.com/v1.0/me/todo/lists
    //       POST https://graph.microsoft.com/v1.0/me/todo/lists  { displayName: GRAPH_CONFIG.todoListName }
    return null;
  },

  // Creates a Microsoft To Do task mirroring a scratchpad task.
  async createTask(_task) {
    // TODO: POST https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks
    //   {
    //     title: task.title,
    //     body: { content: task.note, contentType: "text" },
    //     status: task.status === "done" ? "completed" : "notStarted",
    //     dueDateTime: task.dueDate ? { dateTime: task.dueDate, timeZone: "UTC" } : undefined,
    //     linkedResources: [{ webUrl: task.sourceEmail?.webLink, applicationName: "Outlook", displayName: task.sourceEmail?.subject }]
    //   }
    console.warn("[graph] createTask() not implemented — task saved locally only.", _task.id);
    return null;
  },

  async updateTask(_todoTaskId, _patch) {
    // TODO: PATCH https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks/{taskId}
    return null;
  },

  async deleteTask(_todoTaskId) {
    // TODO: DELETE https://graph.microsoft.com/v1.0/me/todo/lists/{listId}/tasks/{taskId}
    return null;
  },

  // Resolves a permanent, cross-device web link for the currently open message.
  async getMessageWebLink(_itemId) {
    // TODO: GET https://graph.microsoft.com/v1.0/me/messages/{id}?$select=webLink
    // Office.js's item.itemId is only stable enough for this device/session
    // in some hosts, so the durable link should come from Graph once wired up.
    return null;
  },
};
