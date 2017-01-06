(() => {
  // The initialize function must be run each time a new page is loaded
  Office.initialize = (reason) => {
    $(document).ready(() => {
      $('#get-started').click(runGetStarted);
    });
  };

  function runGetStarted() {
    return <%= host %>.run(async (context) => {
      /**
       * Insert your <%= host %> code here
       */
      return await context.sync();
    });
  }
})();
