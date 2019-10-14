import {
	AbstractBox,
	prepareResponseJSOM,
	load,
	executorJSOM,
	getInstance,
	switchType,
	deep1Iterator
} from '../lib/utility'

const KEY_PROP = 'Url'

const lifter = switchType({
	object: item => Object.assign({}, item),
	string: (item = '') => ({ [KEY_PROP]: item }),
	default: () => ({ [KEY_PROP]: '/' })
})

class Attachment {
	constructor(parent, attachments) {
		this.name = 'attachment'
		this.parent = parent
		this.box = getInstance(AbstractBox)(attachments, lifter)
		this.item = parent.box.getHead()

		this.iterator = deep1Iterator({
			contextUrl: parent.parent.contextUrl,
			elementBox: this.box
		})
	}

	async	get(opts = {}) {
		const { clientContexts, result } = await this.iterator(({
			clientContext,
			element
		}) => load(clientContext, this.getSPObject(element, clientContext), opts))

		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, opts)
	}

	getSPObject(element, clientContext) {
		const url = element[KEY_PROP]
		const attachments = this.parent.getSPObject(this.item, clientContext).get_attachmentFiles()
		return url && url !== '/' ? attachments.getByFileName(url) : attachments
	}
}

export default getInstance(Attachment)
