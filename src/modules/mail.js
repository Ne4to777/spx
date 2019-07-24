import axios from 'axios'

import {
	ACTION_TYPES,
	AbstractBox,
	getArray,
	map,
	arrayInit,
	arrayLast,
	prop,
	getInstance,
	switchCase,
	typeOf,
	identity,
	report
} from '../lib/utility'

const EMAIL_RE = /@.+\./

const lifter = switchCase(typeOf)({
	object: identity,
	default: () => ({
		To: ''
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value, lifter)
	}

	getCount() {
		return this.isArray ? this.value.length : 1
	}
}

class Mail {
	constructor(parent, params) {
		this.name = 'mail'
		this.parent = parent
		this.box = getInstance(Box)(params)
		this.count = this.box.getCount()
		this.user = this.parent.user.of
	}

	async	send(opts = {}) {
		const {
			isFake,
			detailed
		} = opts
		const result = await this.box.chain(async element => {
			const {
				From,
				To,
				Subject,
				Body
			} = element
			const recievers = getArray(To)
			if (!recievers.length) throw new Error('Recievers are missed')
			if (!Subject) throw new Error('Subject is missed')
			if (!Body) throw new Error('Body is missed')

			const availableRecievers = []
			const missedUsers = []

			map(el => (EMAIL_RE.test(el) ? availableRecievers : missedUsers).push(el))(recievers)

			let recieversEmails = availableRecievers
			let senderEmail = From

			if (!From) {
				if (missedUsers.length) {
					const [sender, missedRecievers] = await Promise.all([
						this.user()
							.getCurrent({
								...opts,
								isSP: true,
								view: 'Email'
							}),
						this.user(missedUsers)
							.get({
								...opts,
								isSP: true,
								view: 'EMail'
							})
					])
					senderEmail = sender.Email
					recieversEmails = recieversEmails.concat(missedRecievers.map(prop('EMail')))
				} else {
					senderEmail = (await this.user()
						.getCurrent({
							...opts,
							isSP: true,
							view: 'Email'
						}))
						.Email
				}
			} else if (!EMAIL_RE.test(From)) {
				const foundUsers = await this.user(missedUsers.concat(From))
					.get({
						...opts,
						isSP: true,
						view: 'EMail'
					})
				recieversEmails = recieversEmails.concat(arrayInit(foundUsers).map(prop('EMail')))
				senderEmail = arrayLast(foundUsers).EMail
			} else if (missedUsers.length) {
				recieversEmails = recieversEmails
					.concat((await this.user(missedUsers)
						.get({
							...opts,
							isSP: true,
							view: 'EMail'
						}))
						.map(prop('EMail')))
			}

			if (isFake) return undefined
			const response = await axios({
				url: '/_api/SP.Utilities.Utility.SendEmail',
				headers: {
					Accept: 'application/json;odata=verbose',
					'content-type': 'application/json;odata=verbose'
				},
				method: 'POST',
				data: JSON.stringify({
					properties: {
						__metadata: { type: 'SP.Utilities.EmailProperties' },
						From: senderEmail,
						To: { results: recieversEmails },
						Subject,
						Body
					}
				})
			})

			if (detailed) {
				report(`Email is sent\nSender: ${senderEmail}\nRecievers: ${recieversEmails.join(', ')}`, opts)
			}

			return response
		})

		report(`${ACTION_TYPES.send} ${this.count} ${this.name}(s)`, opts)

		return result
	}

	of(params) {
		return getInstance(this.constructor)(params)
	}
}

export default getInstance(Mail)
