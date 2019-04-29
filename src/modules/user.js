import {
	getClientContext,
	executeJSOM,
	prop,
	prepareResponseJSOM,
	flatten,
	isArray,
	isNumber,
	typeOf,
	reduce
} from './../lib/utility'

import web from './../modules/web';

const NAME = 'user';

const getCustomUsersList = _ => web.customUsersList;
const getDefaultUsersList = _ => web.defaultUsersList;

const getByUid = isUsersArray => items => async (opts = {}) => {
	const userIds = reduce(acc => item => {
		switch ((item).constructor.getName()) {
			case 'String':
			case 'Number': acc.push(item); break;
			case 'Object': item.uid && acc.push(item.uid); break;
			case 'SP.User': acc.push(item.get_id()); break;
			case 'SP.FieldUserValue': acc.push(item.get_lookupId()); break;
			case 'SP.ListItem': const uid = item.get_item('uid'); uid && acc.push(uid);
		}
		return acc;
	})([])(items);
	const userList = getCustomUsersList();
	const users = opts.isSP || !userList ? getDefaultUsersList().item(userIds) : userList.item(`Number uid In ${userIds}`);
	const elements = await users.get(opts);
	return isUsersArray || elements.length > 1 ? elements : elements[0];
}

const getByLogin = isUsersArray => items => async (opts = {}) => {
	const userLogins = reduce(acc => item => {
		switch ((item).constructor.getName()) {
			case 'String': acc.push(item); break;
			case 'Object': item.Login && acc.push(item.Login); break;
			case 'SP.User': acc.push(item.get_loginName()); break;
			case 'SP.ListItem': const login = item.get_item('Login'); login && acc.push(login);
		}
		return acc;
	})([])(items);
	const userList = getCustomUsersList();
	const users = opts.isSP || !userList ? getDefaultUsersList().item(`UserName In ${userLogins}`) : userList.item(`Login In ${userLogins}`);
	const elements = await users.get(opts);
	return isUsersArray || elements.length > 1 ? elements : elements[0];
}

const getByName = items => (opts = {}) => {
	const userNames = reduce(acc => item => {
		switch ((item).constructor.getName()) {
			case 'String': acc.push(item); break;
			case 'Object': item.Title && acc.push(item.Title); break;
			case 'SP.User': acc.push(item.get_title()); break;
			case 'SP.ListItem': const title = item.get_item('Title'); title && acc.push(title);
		}
		return acc;
	})([])(items);
	const userList = getCustomUsersList();
	const list = opts.isSP || !userList ? getDefaultUsersList() : userList;
	return list.item(`Title BeginsWith ${userNames}`).get(opts);
}

const getByEMail = isUsersArray => items => async (opts = {}) => {
	const userEMails = reduce(acc => item => {
		switch ((item).constructor.getName()) {
			case 'String': acc.push(item); break;
			case 'Object': item.EMail && acc.push(item.EMail); break;
			case 'SP.User': acc.push(item.get_email()); break;
			case 'SP.ListItem':
				const title = item.get_item('EMail');
				title && acc.push(title);
		}
		return acc;
	})([])(items);
	const userList = getCustomUsersList();
	const list = opts.isSP || !userList ? getDefaultUsersList() : userList;
	const elements = await list.item(`EMail In ${userEMails}`).get(opts);
	return isUsersArray || elements.length > 1 ? elements : elements[0];
}

const getAll = (opts = {}) => {
	const userList = getCustomUsersList();
	(opts.isSP || !userList ? getDefaultUsersList().item() : userList.item(`Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`)).get(opts);
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
						: /@.+\./.test(el)
							? getByEMail(isUsersArray)(values)(opts)
							: getByLogin(isUsersArray)(values)(opts);
			} else {
				return getAll(opts);
			}
		})(isUsersArray)(elements),
		create: (elements => opts => {
			const userList = getCustomUsersList();
			if (!userList) throw new Error('Custom user list is missed');
			const usersToCreate = elements.filter(el => el.uid && el.Title);
			return usersToCreate.length
				? userList.item(elements).create(opts)
				: new Promise((resolve, reject) => reject('missing uid or Title'));
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
				const results = await getDefaultUsersList().item(usersToUpdate).update(opts);
				return isUsersArray || results.length > 1 ? results : results[0];
			} else {
				const userList = getCustomUsersList();
				if (!userList) throw new Error('Custom user list is missed');
				const ids = elements.filter(el => !!el.uid);
				if (ids.length) {
					const users = await web.user(ids).get({ view: ['ID', 'uid'], groupBy: 'uid' });
					const usersToUpdate = elements.reduce((acc, el) => {
						const userID = users[el.uid] ? users[el.uid].ID : void 0;
						if (userID) {
							el.ID = userID;
							delete el.uid;
							acc.push(el);
						}
						return acc
					}, [])
					const results = await userList.item(usersToUpdate).update({ ...opts, view: ['ID', 'uid'] });
					return isUsersArray || results.length > 1 ? results : results[0];
				} else {
					return new Promise((resolve, reject) => { reject('missing uid') });
				}
			}
		})(isUsersArray)(elements),
		deleteWithMissedUid: async opts => {
			const userList = getCustomUsersList();
			if (!userList) throw new Error('Custom user list is missed');
			const users = await userList.item('Number uid IsNull').get(opts);
			return userList.item(users.map(prop('ID'))).delete(opts);
		}
	}
}

user.get = async (opts = {}) => {
	const { isSP } = opts;
	const userList = getCustomUsersList();
	if (isSP || !userList) {
		const clientContext = getClientContext('/');
		const user = clientContext.get_web().get_currentUser();
		return prepareResponseJSOM(opts)(await executeJSOM(clientContext)(user)(opts));
	} else {
		const uid = window._spPageContextInfo ? window._spPageContextInfo.userId : (await user.get({ view: 'Id', isSP: true })).Id;
		return (await userList.item(`Number uid Eq ${uid}`).get(opts))[0];
	}
}

export default user