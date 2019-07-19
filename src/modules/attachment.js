/* eslint class-methods-use-this:0 */
import {
	AbstractBox,
	prepareResponseJSOM,
	load,
	executorJSOM,
	methodEmpty,
	getInstance,
	switchCase,
	typeOf,
	ifThen,
	isArrayFilled,
	map,
	constant,
	deep1Iterator
} from '../lib/utility'

const liftItemType = switchCase(typeOf)({
	object: item => Object.assign({}, item),
	string: (item = '') => ({ Url: item }),
	default: () => ({ Url: undefined })
})

class Box extends AbstractBox {
	constructor(value) {
		super(value)
		this.value = this.isArray
			? ifThen(isArrayFilled)([map(liftItemType), constant([liftItemType()])])(value)
			: liftItemType(value)
	}
}

class Attachment {
	constructor(parent, attachments) {
		this.name = 'attachment'
		this.parent = parent
		this.box = getInstance(Box)(attachments)
		this.listUrl = parent.listUrl
		this.itemUrl = parent.box.head().Url
		this.getItemSPObject = this.parent.getSPObject
		this.getListSPObject = this.parent.parent.getSPObject
		this.getContextSPObject = this.parent.parent.parent.getSPObject

		this.iterator = deep1Iterator({
			contextUrl: parent.parent.contextUrl,
			elementBox: this.box
		})
	}

	async	get(opts = {}) {
		const { clientContexts, result } = await this.iterator(({ clientContext, element }) => {
			const contextSPObject = this.getContextSPObject(clientContext)
			const listSPObject = this.getListSPObject(this.listUrl, contextSPObject)
			const itemSPObject = this.getItemSPObject(this.itemUrl, listSPObject)
			const spObject = this.getSPObject(element)(itemSPObject)
			return load(clientContext)(spObject)(opts)
		})

		await Promise.all(clientContexts.map(clientContext => executorJSOM(clientContext)(opts)))

		return prepareResponseJSOM(opts)(result)
	}

	getSPObject(element, itemSPObject) {
		const url = element.Url
		const attachments = this.getSPObjectCollection(itemSPObject)
		return url ? attachments.getByFileName(url) : attachments
	}

	getSPObjectCollection(url) {
		return methodEmpty('get_attachmentFiles')(url)
	}
}

export default getInstance(Attachment)
