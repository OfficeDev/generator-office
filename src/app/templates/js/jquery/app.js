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
    return <%= host %>.run(function (context) {
      /**
       * Insert your <%= host %> code here
       */
      return context.sync();
    });
  }

})();