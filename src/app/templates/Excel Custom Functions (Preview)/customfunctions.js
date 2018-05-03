function ADD42(num) {
	return num + 42;
}

function ADD42ASYNC(num) {
	// waits 1 second before returning the result
	return new OfficeExtension.Promise(function(resolve) {
		setTimeout(function() {
			resolve(num + 42);
		}, 1000);
	});
}

function ISEVEN(num) {
	return num % 2 == 0;
}

function GETDAY() {
	var d = new Date();
	var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
	return days[d.getDay()];
}

function INCREMENTVALUE(increment, caller){    
	var result = 0;
	var myInterval = setInterval(function(){
		result += increment;
		caller.setResult(result);
	}, 1000);
	caller.onCanceled = function(){
		clearInterval(myInterval);
	}
}

function SECONDHIGHEST(range){
	var highest = range[0][0], secondHighest = range[0][0];
	for(var i = 0; i < range.length;i++){
		for(var j = 0; j < range[i].length;j++){
			if(range[i][j] >= highest){
				secondHighest = highest;
				highest = range[i][j];
			}
			else if(range[i][j] >= secondHighest){
				secondHighest = range[i][j];
			}
		}
	}
	return secondHighest;
}

