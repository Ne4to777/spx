import * as utility from './../utility';

export default class User {
	constructor(parent, elementUrl) {
		this.parent = parent;
		this.elementUrl = elementUrl;
		this.contextUrl = this.parent.contextUrl;
		this.clientContext = this.parent.clientContext;
		this.executeBinded = this.execute.bind(this, 'user', this.contextUrl, this.elementUrl);
	}
	async get(opts) {
		return this.executeBinded(this._getSPObject.bind(this), opts);
	}

	_getSPObject(elementUrl) {
		return elementUrl ? this.clientContext.get_web().getUserById(elementUrl) : this.clientContext.get_web().get_currentUser();
	}
}

let legacy = {
	async getUserCurrent(opts = {}) {
		let cached = opts.cached;
		let cache = this.getUserCurrent.cache;
		if (cached && cache) {
			return cache;
		} else {
			opts.caml = 'Number uid Eq <UserID/>';
			let users = await this.getItems(this.LIBRARY.users, opts);
			let preparedResponse = this.prepareResponse(users[0], opts);
			cached && (this.getUserCurrent.cache = preparedResponse);
			return preparedResponse;
		}
	},

	async getUsers(users, opts = {}) {
		if (typeOf(users) !== 'array') users = [users];
		let userIds = [];
		(function flatternArrayR(items) {
			let item;
			for (let i = 0; i < items.length; i++) {
				item = items[i];
				if (typeOf(item) === 'array') {
					flatternArrayR(item)
				} else {
					switch ((item).constructor.getName()) {
						case 'String':
						case 'Number':
							userIds.push(item);
							break;
						case 'Object':
							item.uid && userIds.push(item.uid);
							break;
						case 'SP.User':
							userIds.push(item.get_id())
							break;
						case 'SP.FieldUserValue':
							userIds.push(item.get_lookupId())
							break;
						case 'SP.ListItem':
							let uid = item.get_item('uid');
							uid && userIds.push(uid);
							break;
					}
				}
			}
		}(users));
		opts.caml = `number uid in ${userIds}`;
		return await this.getItems(this.LIBRARY.users, opts)
	},

	async getUserByUid(uid) {
		return await this.getItems(this.LIBRARY.users, {
			caml: `number uid in ${uid}`
		})
	},

	async getUserByLogin(login) {
		return await this.getItems(this.LIBRARY.users, {
			caml: `login in ${login}`
		})
	},

	async getUserByUidOrLogin(uidOrLogin) {
		if (typeOf(uidOrLogin) === 'number' || ~~uidOrLogin) {
			return await this.getUserByUid(uidOrLogin)
		} else {
			return await this.getUserByLogin(uidOrLogin)
		}
	},

	async getUserByName(name) {
		return await this.getItems(this.LIBRARY.users, {
			caml: `Title Contains ${name}`
		})
	}
}