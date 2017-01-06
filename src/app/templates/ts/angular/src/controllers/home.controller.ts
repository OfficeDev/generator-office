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
