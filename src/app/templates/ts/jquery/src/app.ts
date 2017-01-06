(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#get-started').click(runGetStarted);
    });
  };

  // Reads data from current document selection and displays a notification
  function runGetStarted() {
    return <%= host %>.run(async (context) => {
      /**
       * Insert your <%= host %> code here
       */
      return await context.sync();
    });
  }
})();
