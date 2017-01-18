/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

declare var angular: any;

(function(){

  // create
  var officeAddin = angular.module('<%= projectInternalName %>', []);

  // configure
  officeAddin.config(['$logProvider', function($logProvider) {
    // set debug logging to on
    if ($logProvider.debugEnabled) {
      $logProvider.debugEnabled(true);
    }
  }]);

  // when Office has initalized, manually bootstrap the app
  Office.initialize = function() {
    angular.bootstrap(document.body, ['officeAddin']);
  };

})();