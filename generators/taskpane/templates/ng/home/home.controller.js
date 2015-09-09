(function () {
  'use strict';

  angular.module('officeAddin')
    .controller('homeController', ['dataService', homeController]);

  /**
   * Controller constructor
   */
  function homeController(dataService) {
    var vm = this;
    vm.title = 'home controller';
    vm.dataObject = {};

    getDataFromService();

    function getDataFromService() {
      dataService.getData()
        .then(function (response) {
          vm.dataObject = response;
        });
    }
  }

})();