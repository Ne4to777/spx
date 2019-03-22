import {
	getClientContext,
	executeJSOM,
	prop,
	prepareResponseJSOM,
	flatten,
	isArray,
	isNumber,
	typeOf
} from './../lib/utility'

import site from './../modules/site';

const USER_WEB = 'AM';
const USER_LIST = 'UsersAD';
const USER_LIST_GUID = 'b327d30a-b9bf-4728-a3c1-a6b4f0253ff2';

const NAME = 'user';

const getByUid = isUsersArray => items => async (opts = {}) => {
	let users;
	const userIds = [];
	const { isSP } = opts;
	for (const item of items) {
		switch ((item).constructor.getName()) {
			case 'String':
			case 'Number': userIds.push(item); break;
			case 'Object': item.uid && userIds.push(item.uid); break;
			case 'SP.User': userIds.push(item.get_id()); break;
			case 'SP.FieldUserValue': userIds.push(item.get_lookupId()); break;
			case 'SP.ListItem':
				const uid = item.get_item('uid');
				uid && userIds.push(uid);
		}
	}
	if (isSP) {
		users = site().list(USER_LIST_GUID).item(userIds)
	} else {
		users = site(USER_WEB).list(USER_LIST).item(`Number uid In ${userIds}`);
	}
	const elements = await users.get(opts);
	return isUsersArray || elements.length > 1 ? elements : elements[0];
}

const getByLogin = isUsersArray => items => async (opts = {}) => {
	let users;
	const userLogins = [];
	const { isSP } = opts;
	for (const item of items) {
		switch ((item).constructor.getName()) {
			case 'String': userLogins.push(item); break;
			case 'Object': item.Login && userLogins.push(item.Login); break;
			case 'SP.User': userLogins.push(item.get_loginName()); break;
			case 'SP.ListItem':
				const login = item.get_item('Login');
				login && userLogins.push(login);
		}
	}
	if (isSP) {
		users = site().list(USER_LIST_GUID).item(`LoginName In ${userLogins}`)
	} else {
		users = site(USER_WEB).list(USER_LIST).item(`Login In ${userLogins}`);
	}
	const elements = await users.get(opts);
	return isUsersArray || elements.length > 1 ? elements : elements[0];
}

const getByName = items => (opts = {}) => {
	let list;
	const userNames = [];
	const { isSP } = opts;
	for (const item of items) {
		switch ((item).constructor.getName()) {
			case 'String': userNames.push(item); break;
			case 'Object': item.Title && userNames.push(item.Title); break;
			case 'SP.User': userNames.push(item.get_title()); break;
			case 'SP.ListItem':
				const title = item.get_item('Title');
				title && userNames.push(title);
		}
	}
	if (isSP) {
		list = site().list(USER_LIST_GUID)
	} else {
		list = site(USER_WEB).list(USER_LIST)
	}
	return list.item(`Title BeginsWith ${userNames}`).get(opts);
}

const getByEMail = isUsersArray => items => async (opts = {}) => {
	let list;
	const userEMails = [];
	const { isSP } = opts;
	for (const item of items) {
		switch ((item).constructor.getName()) {
			case 'String': userEMails.push(item); break;
			case 'Object': item.EMail && userEMails.push(item.EMail); break;
			case 'SP.User': userEMails.push(item.get_email()); break;
			case 'SP.ListItem':
				const title = item.get_item('EMail');
				title && userEMails.push(title);
		}
	}
	if (isSP) {
		list = site().list(USER_LIST_GUID)
	} else {
		list = site(USER_WEB).list(USER_LIST)
	}
	const elements = await list.item(`EMail In ${userEMails}`).get(opts);
	return isUsersArray || elements.length > 1 ? elements : elements[0];
}

const getAll = (opts = {}) => {
	const { isSP } = opts;
	if (isSP) {
		list = site().list(USER_LIST_GUID);
	} else {
		list = site(USER_WEB).list(USER_LIST);
		caml = `Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`
	}
	return list.item(caml).get(opts);
}

// Interface

const user = users => {
	const isUsersArray = isArray(users);
	const elements = users ? (isUsersArray ? flatten(users) : [users]) : [];
	return {
		get: (isUsersArray => elements => (opts = {}) => {
			if (elements.length) {
				const el = elements[0];
				const values = elements;
				return isNumber(el) || ~~el == el || (typeOf(el) === 'object' && el.uid)
					? getByUid(isUsersArray)(values)(opts)
					: /\s/.test(el) || /[а-яА-ЯЁё]/.test(el)
						? getByName(values)(opts)
						: /@/.test(el)
							? getByEMail(isUsersArray)(values)(opts)
							: getByLogin(isUsersArray)(values)(opts);
			} else {
				return getAll(opts);
			}
		})(isUsersArray)(elements),
		create: (elements => opts => {
			const usersToCreate = elements.filter(el => el.uid && el.Title);
			return usersToCreate.length
				? site(USER_WEB).list(USER_LIST).item(elements).create(opts)
				: new Promise((resolve, reject) => reject(new Error('missing uid or Title')));
		})(elements),
		getByUid: getByUid(isUsersArray)(elements),
		getByLogin: getByLogin(isUsersArray)(elements),
		getByEMail: getByEMail(isUsersArray)(elements),
		getByName: getByName(elements),

		update: (isUsersArray => elements => async (opts = {}) => {
			const { isSP } = opts;
			if (isSP) {
				const usersToUpdate = elements.reduce((acc, el) => {
					if (el.uid) {
						el.ID = el.uid;
						delete el.uid;
						acc.push(el);
					}
					return acc;
				}, []);
				const results = await site().list(USER_LIST_GUID).item(usersToUpdate).update(opts);
				return isUsersArray || results.length > 1 ? results : results[0];
			} else {
				const ids = elements.filter(el => !!el.uid);
				if (ids.length) {
					const users = await site.user(ids).get({ view: ['ID', 'uid'], groupBy: 'uid' });
					const usersToUpdate = elements.reduce((acc, el) => {
						const userID = users[el.uid] ? users[el.uid].ID : void 0;
						if (userID) {
							el.ID = userID;
							delete el.uid;
							acc.push(el);
						}
						return acc
					}, [])
					const results = await site(USER_WEB).list(USER_LIST).item(usersToUpdate).update({ ...opts, view: ['ID', 'uid'] });
					return isUsersArray || results.length > 1 ? results : results[0];
				} else {
					return new Promise((resolve, reject) => { reject(new Error('missing uid')) });
				}
			}
		})(isUsersArray)(elements),
		deleteWithMissedUid: async opts => {
			const users = await site(USER_WEB).list(USER_LIST).item('Number uid IsNull').get(opts);
			return site(USER_WEB).list(USER_LIST).item(users.map(prop('ID'))).delete(opts);
		}
	}
}

user.get = async (opts = {}) => {
	const { fromSP } = opts;
	if (fromSP) {
		const clientContext = getClientContext('/');
		const user = clientContext.get_web().get_currentUser();
		return prepareResponseJSOM(opts)(await executeJSOM(clientContext)(user)(opts));
	} else {
		const uid = window._spPageContextInfo ? window._spPageContextInfo.userId : (await user.get({ view: 'Id', fromSP: true })).Id;
		return (await site(USER_WEB).list(USER_LIST).item(`Number uid Eq ${uid}`).get(opts))[0];
	}
}

export default user