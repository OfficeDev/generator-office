/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  // The initialize function must be run each time a new page is loaded
  Office.initialize = function (reason) {
    $(document).ready(function () {
      app.initialize();
      $('#get-started').click(runGetStarted);
    });
  };

  function runGetStarted() {
    return <%= hostDisplayName %>.run(function (context) {
      /**
       * Insert your <%= host %> code here
       */
      return context.sync();
    });
  }

})();