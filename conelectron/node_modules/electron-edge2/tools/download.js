var http = require('https');

var url = 'https://iojs.org/dist/v' + process.argv[3] + '/' 
	+ (process.argv[2] === 'x86' ? '/win-x86/iojs.exe' : '/win-x64/iojs.exe');

console.log(url);
http.get(url, function (res) {
	if (res.statusCode !== 200)
		throw new Error('HTTP response status code ' + res.statusCode);

	var stream = require('fs').createWriteStream(process.argv[4] + '/iojs.exe');
	res.pipe(stream);
});
