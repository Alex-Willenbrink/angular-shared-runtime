
const SetRuntimeVisibleHelper = (visible) => {
    let p;
    if (visible) {
        p = Office.addin.showAsTaskpane();
    } else {
        p = Office.addin.hide();
    }

    return p
        .then(() => {
            return visible;
        })
        .catch(error => {
            return error.code;
        });
};

Office.onReady(() => {
    console.log("office is ready");
    // If needed, Office.js is ready to be called
});

/**
* Shows a notification when the add-in command is executed.
* @param event
*/
function action(event) {
    // Your code goes here

    // Be sure to indicate when the add-in command function is complete
    event.completed();
}

function btnOpenTaskpane(event) {
    console.log('Open task pane button pressed');
    // Your code goes here
    SetRuntimeVisibleHelper(true);
    g.state.isTaskpaneOpen = true;
    event.completed();
}

function btnCloseTaskpane(event) {
    console.log('Open task pane button pressed');
    // Your code goes here
    SetRuntimeVisibleHelper(false);
    g.state.isTaskpaneOpen = false;
    event.completed();
}

function getGlobal() {
    return typeof self !== "undefined"
        ? self
        : typeof window !== "undefined"
            ? window
            : typeof global !== "undefined"
                ? global
                : undefined;
}

const g = getGlobal();

// the add-in command functions need to be available in global scope
g.action = action;
g.btnopentaskpane = btnOpenTaskpane;
g.btnclosetaskpane = btnCloseTaskpane;
