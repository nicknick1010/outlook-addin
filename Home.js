Office.initialize = function (reason) {
    $(document).ready(function () {
        // Add event handler for item send
        Office.context.mailbox.item.addHandlerAsync(Office.EventType.ItemSend, handleItemSend);
    });
};

function handleItemSend(eventArgs) {
    // Display a confirmation dialog
    var userConfirmation = confirm("Do you really want to send this message?");
    if (userConfirmation) {
        eventArgs.completed({ allowEvent: true });
    } else {
        eventArgs.completed({ allowEvent: false });
    }
}
