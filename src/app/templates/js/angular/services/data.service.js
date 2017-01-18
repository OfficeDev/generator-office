/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

'use strict';

(function () {

  angular
    .module('<%= projectInternalName %>')
    .service('DataService', ['$q', DataService]);

  /**
   * Data Service.
   */
  function DataService($q) {

    // public signature of the service
    return {
      getData: getData
    };

    function getData() {
      var deferred = $q.defer();

      deferred.resolve([
        {
          propertyOne: 'valueOne',
          propertyTwo: 'valueTwo',
        }
      ]);

      return deferred.promise;
    }
  }
})();
