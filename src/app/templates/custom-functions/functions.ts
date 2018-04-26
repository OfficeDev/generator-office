function add10(num) {
    return num + 10;
}
/**
 * Add twenty to the number.
 * @param num
 */
function add20(num: number) {
    return num + 20;
}

function isEven(num: number) {
    return num % 2 == 0;
}

function test(str: string) {
    return str.length;
}

function getDay() {
    var d = new Date();
    var days = ["Sunday", "Monday", "Tuesday", "Wednesday", "Thursday", "Friday", "Saturday"];
    return days[d.getDay()];
}

function getDummyString(len: number) {
    var base = "123456789 ";
    var div = Math.floor(len / 10);
    var rem = len - (div * 10);
    var result = "";
    for (var i = 0; i < div; i++) {
        result += base;
    }
    result += base.substr(0, rem);
    return result;
}
