import * as utility from './../utility';
import spx from './../modules/site';

const USER_WEB = '/AM';
const USER_LIST = 'UsersAD';
const USER_LIST_GUID = '/b327d30a-b9bf-4728-a3c1-a6b4f0253ff2';

export default class User {
	constructor(parent, elementUrl) {
		this._parent = parent;
		this._elementUrl = elementUrl;
		this._elementUrlIsArray = typeOf(this._elementUrl) === 'array';
	}

	// Inteface

	async get(opts) {
		let el = typeOf(this._elementUrl) === 'array' ? this._elementUrl[0] : this._elementUrl;
		if (typeOf(el) === 'nubmer' || ~~el == el) {
			return this.getByUid(opts);
		} else {
			return /\s/.test(el) ? this.getByName(opts) : this.getByLogin(opts);
		}
	}

	async getByUid(opts) {
		let el = this._elementUrl;
		const list = spx(USER_WEB).list(USER_LIST);
		if (el) {
			if (typeOf(el) !== 'array') el = [el];
			const userIds = [];
			const flatternArrayR = items => {
				for (let item of items) {
					if (typeOf(item) === 'array') {
						flatternArrayR(item)
					} else {
						switch ((item).constructor.getName()) {
							case 'String':
							case 'Number': userIds.push(item); break;
							case 'Object': item.uid && userIds.push(item.uid); break;
							case 'SP.User': userIds.push(item.get_id()); break;
							case 'SP.FieldUserValue': userIds.push(item.get_lookupId()); break;
							case 'SP.ListItem':
								let uid = item.get_item('uid');
								uid && userIds.push(uid);
						}
					}
				}
			}
			flatternArrayR(el);
			return list.item(`Number uid Eq ${userIds}`).get(opts);
		} else {
			return list.item(`Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`).get(opts);
		}
	}

	async getByLogin(opts) {
		let el = this._elementUrl;
		const list = spx(USER_WEB).list(USER_LIST);
		if (el) {
			if (typeOf(el) !== 'array') el = [el];
			const userLogins = [];
			const flatternArrayR = items => {
				for (let item of items) {
					if (typeOf(item) === 'array') {
						flatternArrayR(item)
					} else {
						switch ((item).constructor.getName()) {
							case 'String': userLogins.push(item); break;
							case 'Object': item.Login && userLogins.push(item.Login); break;
							case 'SP.User': userLogins.push(item.get_loginName()); break;
							case 'SP.ListItem':
								let login = item.get_item('Login');
								login && userLogins.push(login);
						}
					}
				}
			}
			flatternArrayR(el);
			return list.item(`Login Eq ${userLogins}`).get(opts);
		} else {
			return list.item(`Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`).get(opts);
		}
	}

	async getByName(opts) {
		let el = this._elementUrl;
		const list = spx(USER_WEB).list(USER_LIST);
		if (el) {
			if (typeOf(el) !== 'array') el = [el];
			const userNames = [];
			const flatternArrayR = items => {
				for (let item of items) {
					if (typeOf(item) === 'array') {
						flatternArrayR(item)
					} else {
						switch ((item).constructor.getName()) {
							case 'String': userNames.push(item); break;
							case 'Object': item.Title && userNames.push(item.Title); break;
							case 'SP.User': userNames.push(item.get_Title()); break;
							case 'SP.ListItem':
								let title = item.get_item('Title');
								title && userNames.push(title);
						}
					}
				}
			}
			flatternArrayR(el);
			return list.item(`Title Eq ${userNames}`).get(opts);
		} else {
			return list.item(`Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`).get(opts);
		}
	}

	async getCurrent(opts) {
		const clientContext = utility.getClientContext('/');
		const user = clientContext.get_web().get_currentUser();
		utility.load(clientContext, user);
		await utility.executeQueryAsync(clientContext);
		return spx(USER_WEB).list(USER_LIST).item(`Number uid Eq ${user.get_id()}`).get(opts);
	}

	async create(opts = {}) {
		const el = this._elementUrl;
		opts.view = ['ID', 'uid'];

		const isOk = this._elementUrlIsArray ? this._elementUrlIsArray.reduce((acc, el) => {
			if (acc && el && (!el.uid || !el.Title)) acc = false;
			return acc
		}, true) : (!!el.uid && !!el.Title);
		if (isOk) {
			return spx(USER_WEB).list(USER_LIST).item(el).create(opts);
		} else {
			return new Promise((resolve, reject) => { reject(new Error('missing uid or Title')) });
		}
	}

	async update(opts = {}) {
		let elements = this._elementUrl;
		opts.view = ['ID', 'uid'];
		if (!this._elementUrlIsArray) elements = [elements];
		const isOk = elements.reduce((acc, el) => {
			if (acc && el && !el.uid) acc = false;
			return acc
		}, true)
		const ids = elements.map(el => el.uid)
		if (isOk) {
			const users = await spx.user(ids).getByUid({ view: ['ID', 'uid'], groupBy: 'uid' });
			const usersToUpdate = elements.map(el => {
				const userID = users[el.uid] ? users[el.uid].ID : void 0;
				if (userID) {
					el.ID = userID;
					delete el.uid;
					return el;
				}
			})
			return spx(USER_WEB).list(USER_LIST).item(usersToUpdate).update(opts);
		} else {
			return new Promise((resolve, reject) => { reject(new Error('missing uid or Title')) });
		}
	}

	async updateSP(opts) {
		let elements = this._elementUrl;
		if (!this._elementUrlIsArray) elements = [elements];
		const usersToUpdate = elements.map(el => {
			if (userID) {
				el.ID = el.uid;;
				delete el.uid;
				return el;
			}
		})
		return spx('/').list(USER_LIST_GUID).item(usersToUpdate).update(opts);
	}

	async deleteWithMissedUid(opts) {
		const users = await spx(USER_WEB).list(USER_LIST).item('Number uid IsNull').get(opts);
		return spx(USER_WEB).list(USER_LIST).item(users.map(el => el.ID)).delete(opts);
	}

	// Internal

	get _name() { return 'user' }
}