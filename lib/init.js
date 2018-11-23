const fs = require('fs');
const Cpass = require('cpass').Cpass;
const rl = require('readline').createInterface({
	input: process.stdin,
	output: process.stdout
});

fs.stat('./config/private.json', (err, stats) => {
	if (err) {
		rl.question('Username: ', (username) => {
			rl.question('Password: ', (password) => {
				rl.question('Project absolute path (f.e. "Z:/a/b"): ', (path) => {
					fs.writeFileSync('./config/private.json', JSON.stringify({
						siteUrl: 'http://aura.dme.aero.corp',
						strategy: 'OnpremiseUserCredentials',
						domain: 'dme',
						username: username,
						password: new Cpass().encode(password),
						path: path.replace(/\//g, '\\')
					}));
					rl.close();
				})
			})
		});
	} else {
		rl.close();
	}
})