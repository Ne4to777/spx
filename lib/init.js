const fs = require('fs');
const Cpass = require('cpass').Cpass;
const rl = require('readline').createInterface({
	input: process.stdin,
	output: process.stdout
});

const promisify = f => {
	const argsDeclared = [];
	for (let i = 0, fLength = f.length - 1; i < fLength; i++) argsDeclared.push(void 0);
	return (...args) => new Promise((resolve, reject) => f(...argsDeclared.map((u, i) => args[i]), x => resolve(x)))
}

const question = promisify(rl.question.bind(rl));

fs.stat('./config/private.json', async (err, stats) => {
	if (err) {
		const username = await question('Username: ');
		const password = await question('Password: ');
		const path = await question('Project absolute path (f.e. "Z:/a/b"): ');
		const filename = await question('Output filename: ');
		const library = await question('Library name: ')
		await fs.writeFileSync('./config/private.json', JSON.stringify({
			siteUrl: 'http://aura.dme.aero.corp',
			strategy: 'OnpremiseUserCredentials',
			domain: 'dme',
			username: username,
			password: new Cpass().encode(password),
			path: path.replace(/\//g, '\\'),
			filename: filename,
			library: library
		}));
	}
	rl.close();
})