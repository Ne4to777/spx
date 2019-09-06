const fs = require('fs')
const { Cpass } = require('cpass')
const rl = require('readline').createInterface({
	input: process.stdin,
	output: process.stdout
})

const question = msg => new Promise(resolve => rl.question(msg, resolve))

fs.stat('./dev/private.json', async (err) => {
	if (err) {
		await fs.writeFileSync(
			'./dev/private.json',
			JSON.stringify({
				siteUrl: (await question('Host: ')),
				strategy: 'OnpremiseUserCredentials',
				domain: (await question('Domain: ')),
				username: await question('Username: '),
				password: new Cpass().encode(await question('Password: ')),
				path:
					(await question('Project absolute path: ')).replace(/\//g, '\\'),
				customUsersWeb: (await question('Custom users web: ')),
				customUsersList: (await question('Custom users list: ')),
				defaultUsersList: (await question(
					'Default users list (User Information List): '
				)) || 'User Information List'
			})
		)
	}
	rl.close()
})
