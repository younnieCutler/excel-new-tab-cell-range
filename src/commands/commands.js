Office.onReady(() => {
    Office.actions.associate('openInCellFocus', openInCellFocus);
});

function openInCellFocus(event) {
    // Show taskpane first, then trigger capture
    Office.addin.showAsTaskpane().then(() => {
        // captureSelectedRange is exposed by taskpane.js in the Shared Runtime context
        if (typeof window.captureSelectedRange === 'function') {
            window.captureSelectedRange({ preferTrackedSelection: true });
        }
    });
    event.completed();
}
