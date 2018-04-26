var httpserver = require('http-server');
var newfunctions = require("./newfunctions");
var number = newfunctions.addTen(20);
var isEven = newfunctions.isEvenYo(number);
server = httpserver.createServer(function(request,response){
    }).listen(8080, "localhost");
    console.log('Server running at http://localhost:8080/');
    console.log(number);
    console.log(isEven);

