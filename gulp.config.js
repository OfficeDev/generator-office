'use strict';

var GulpConfig = (function(){
  var config = {
    generatorJs: './generators/**/*.js',
    testJs: './test/**/.js',

    allJs: [
      './*.js',
      './generators/**/*.js',
      './test/**/*.js',
      '!**/scripts/*.js'
    ]
  };

  return config;
})();

module.exports = GulpConfig;
