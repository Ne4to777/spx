export function typeOf(variable) {
	return Object.prototype.toString.call(variable).slice(8, -1).toLowerCase();
}

export function log() {
	console.log('------- Begin');
	for (var arg of arguments) {
		console.log(arg);
	}
	console.log('------- End');
}