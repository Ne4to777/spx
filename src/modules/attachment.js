/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
	prepareResponseJSOM,
	load,
	executorJSOM,
	getInstance,
	switchCase,
	typeOf,
	deep1Iterator
} from '../lib/utility'

const lifter = switchCase(typeOf)({
	object: item => Object.assign({}, item),
	string: (item = '') => ({ Url: item }),
	default: () => ({ Url: '/' })
})

class Attachment {
	constructor(parent, attachments) {
		this.name = 'attachment'
		this.parent = parent
		this.box = getInstance(AbstractBox)(attachments, lifter)
		this.listProp = parent.parent.box.getHeadPropValue()
		this.item = parent.box.getHead()
		this.getItemSPObject = parent.getSPObject.bind(parent)
		this.getListSPObject = parent.parent.getSPObject.bind(parent.parent)
		this.getContextSPObject = parent.parent.parent.getSPObject.bind(parent.parent.parent)

		this.iterator = deep1Iterator({
			contextUrl: parent.parent.contextUrl,
			elementBox: this.box
		})
	}

	async	get(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(this.listProp, contextSPObject)
			const itemSPObject = this.getItemSPObject(this.item, listSPObject)
			const spObject = this.getSPObject(element, itemSPObject)
			return load(clientContext, spObject, opts)
		})

		await Promise.all(clientContexts.map(executorJSOM))

		return prepareResponseJSOM(result, opts)
	}

	getSPObject(element, itemSPObject) {
		const url = element.Url
		const attachments = itemSPObject.get_attachmentFiles()
		return url && url !== '/' ? attachments.getByFileName(url) : attachments
	}
}

export default getInstance(Attachment)
