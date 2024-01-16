'use client';

import Script from 'next/script';

export default function Page() {
  return (
    <>
      <Script
        src='https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js'
        // strategy='beforeInteractive'
        onReady={() => {
          Office.onReady(() => {
            function insertOrEditField(event: Office.AddinCommands.Event) {
              const url = 'https://app.cloudfiles.build/office/mail/dialogs';
              Office.context.ui.displayDialogAsync(url, { height: 50, width: 50 }, function (result) {
                var dialog = result.value;

                dialog.addEventHandler(Office.EventType.DialogMessageReceived, function (arg) {
                  // Check if the received event has the 'message' property
                  if ('message' in arg) {
                    var messageFromDialog = arg.message;
                    console.log('Message from dialog: ' + messageFromDialog);
                  } else {
                    // Handle the case where 'message' property is not present (optional)
                    console.warn("Unexpected event format. No 'message' property found.");
                  }
                });
              });
            }
            Office.actions.associate('insertOrEditField', insertOrEditField);

            function getGlobal() {
              return typeof self !== 'undefined'
                ? self
                : typeof window !== 'undefined'
                ? window
                : typeof global !== 'undefined'
                ? global
                : undefined;
            }

            const g = getGlobal() as any;

            console.log(g);

            // The add-in command functions need to be available in global scope
            g.action = insertOrEditField;
          });
        }}
      />
      {/* {isOfficeInitialized ? children : <h1>Loading...</h1>} */}
      {/* <Script src='https://app.cloudfiles.build/office/commands.js' /> */}
    </>
  );
}
