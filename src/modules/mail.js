import {
	ACTION_TYPES,
	AbstractBox,
	getArray,
	map,
	getInstance,
	switchType,
	identity,
	report,
	executorREST,
	prepareResponseREST
} from '../lib/utility'


const EMAIL_RE = /\S+@\S+\.\S+/
const KEY_PROP = 'To'

const lifter = switchType({
	object: identity,
	default: () => ({
		[KEY_PROP]: ''
	})
})

class Box extends AbstractBox {
	constructor(value = '') {
		super(value, lifter)
		this.joinProp = KEY_PROP
		this.prop = KEY_PROP
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
		this.user = parent.user.bind(parent)
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

			const recieversEmails = availableRecievers
			let senderEmail = From

			if (!From) {
				const sender = await this.user().get({ isSP: true, view: 'Email' })
				senderEmail = sender.Email
			}
			if (isFake) return undefined

			const response = await executorREST('/', {
				url: '/_api/SP.Utilities.Utility.SendEmail',
				headers: {
					Accept: 'application/json;odata=verbose',
					'Content-Type': 'application/json;odata=verbose'
				},
				method: 'POST',
				body: JSON.stringify({
					properties: {
						__metadata: {
							type: 'SP.Utilities.EmailProperties'
						},
						From: senderEmail,
						To: {
							results: recieversEmails
						},
						Subject,
						Body
					}
				})
			})

			if (detailed) {
				report(`Mail is sent\nSender: ${senderEmail}\nRecievers: ${recieversEmails.join(', ')}`, opts)
			}

			return prepareResponseREST(response, opts)
		})
		report(`${ACTION_TYPES.send} ${this.count} ${this.name}${this.count > 1 ? '(s)' : ''}`, opts)
		return result
	}

	of(params) {
		return getInstance(this.constructor)(params)
	}
}

export default getInstance(Mail)
