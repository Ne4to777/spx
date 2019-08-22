const fs = require('fs')
const { Cpass } = require('cpass')
const rl = require('readline').createInterface({
	input: process.stdin,
	output: process.stdout
})

const promisify = f => {
	const argsDeclared = []
	for (let i = 0, fLength = f.length - 1; i < fLength; i += 1) argsDeclared.push(undefined)
	return (...args) => new Promise((resolve) => f(...argsDeclared.map((u, i) => args[i]), x => resolve(x)))
}

const question = promisify(rl.question.bind(rl))

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
				filename: (await question('Output filename (index.js): ')) || 'index.js',
				library: (await question('Library name: ')),
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
