Office.onReady(function () {
  Office.actions.associate("openZendesk", openZendesk);
});

function openZendesk(event) {
  Office.context.ui.openBrowserWindow("https://pickeringcollege.zendesk.com/");
  event.completed();
}
