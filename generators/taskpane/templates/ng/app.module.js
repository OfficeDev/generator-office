(function(){
  'use strict';

  // create
  var officeAddin = angular.module('officeAddin', [
    'ngRoute',
    'ngSanitize'
  ]);

  // configure
  officeAddin.config(['$logProvider', function($logProvider){
    // set debug logging to on
    if ($logProvider.debugEnabled) {
      $logProvider.debugEnabled(true);
    }
  }]);

  // when Office has initalized, manually bootstrap the app
  Office.initialize = function(){
    console.log('>>> Office.initialize()');
    angular.bootstrap(document.getElementById('container'), ['officeAddin']);
  };

})();
