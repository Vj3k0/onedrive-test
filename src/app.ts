/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#run').click(run);
    });
  };

  async function run() {
    
    await Word.run(async (context) => {
      /**
       * Insert your Word code here
       */
      await context.sync();

      Office.context.document.settings.refreshAsync(function () {
        var foo = Office.context.document.settings.get('hello');
        if (!foo) {
          Office.context.document.settings.set('hello', 'world');
          Office.context.document.settings.saveAsync(function (asyncResult) {
            $('#content').append('Settings saved with status: ' + asyncResult.status);
          });
        }
        else {
          $('#content').append('Value found: ' + foo);
        }
      });

    });
    
  }
})();
