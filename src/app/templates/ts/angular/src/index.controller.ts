/* 
 * Copyright (c) Microsoft Corporation. All rights reserved. Licensed under the MIT license.
 * See LICENSE in the project root for license information.
 */

(function () {
  angular
    .module('<%= projectInternalName %>')
    .controller('HomeController', ['DataService', HomeController]);

  /**
   * Home Controller
   */
  function HomeController(DataService) {
    this.title = 'home controller';
    this.dataService = DataService;
    this.dataObject = {};
    this.getDataFromService();
  }

  HomeController.prototype.getDataFromService = function () {
    var self = this;
    this.dataService.getData().then(function (response) {
      self.dataObject = response;
    });
  }

})();
