/*
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {

  // create
  angular
    .module('<%= projectInternalName %>', [])
    .controller('HomeController', [HomeController])
    .config(['$logProvider', function ($logProvider) {
      // set debug logging to on
      if ($logProvider.debugEnabled) {
        $logProvider.debugEnabled(true);
      }
    }]);

  /**
   * Home Controller
   */
  function HomeController() {
    this.title = 'Home';
    console.log(`${this.title} is ready!`);

    this.run = async () => {
      await <%= hostDisplayName %>.run(async (context) => {
        /**
         * Insert your <%= host %> code here
         */
        await context.sync();
      });
    }
  }

  // when Office has initalized, manually bootstrap the app
  Office.initialize = function () {
    angular.bootstrap(document.body, ['<%= projectInternalName %>']);
  };

})();