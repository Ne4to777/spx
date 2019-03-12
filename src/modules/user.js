import {
	getClientContext,
	executeJSOM,
	prop,
	prepareResponseJSOM,
	flatten,
	isArray,
	isNumber
} from './../utility'

import site from './../modules/site';

const USER_WEB = '/AM';
const USER_LIST = 'UsersAD';
const USER_LIST_GUID = '/b327d30a-b9bf-4728-a3c1-a6b4f0253ff2';



const NAME = 'user';

const getByUid = isUsersArray => items => async opts => {
	const userIds = [];
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
	const elements = await site(USER_WEB).list(USER_LIST).item(`Number uid In ${userIds}`).get(opts);
	return isUsersArray ? elements : elements[0];
}

const getByLogin = isUsersArray => items => async opts => {
	const userLogins = [];
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
	const elements = await spx(USER_WEB).list(USER_LIST).item(`Login In ${userLogins}`).get(opts);
	return isUsersArray ? elements : elements[0];
}

const getByName = items => opts => {
	const userNames = [];
	for (const item of items) {
		switch ((item).constructor.getName()) {
			case 'String': userNames.push(item); break;
			case 'Object': item.Title && userNames.push(item.Title); break;
			case 'SP.User': userNames.push(item.get_Title()); break;
			case 'SP.ListItem':
				const title = item.get_item('Title');
				title && userNames.push(title);
		}
	}
	return spx(USER_WEB).list(USER_LIST).item(`Title In ${userNames}`).get(opts);
}

// Interface

const user = users => {
	const isUsersArray = isArray(users);
	const elements = isUsersArray ? flatten(users) : [users];
	return {
		get: (isUsersArray => elements => opts => {
			const list = spx(USER_WEB).list(USER_LIST);
			if (elements.length) {
				const el = elements[0];
				const values = elements;
				return isNumber(el) || ~~el == el
					? getByUid(isUsersArray)(values)(opts) : /\s/.test(el)
						? getByName(values)(opts) : getByLogin(isUsersArray)(values)(opts);
			} else {
				return list.item(`Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`).get(opts);
			}
		})(isUsersArray)(elements),
		getSP: (isUsersArray => elements => opts => {
			const list = spx().list(USER_LIST_GUID);
			if (elements.length) {
				const el = elements[0];
				const values = elements;
				return isNumber(el) || ~~el == el
					? getByUid(isUsersArray)(values)(opts) : /\s/.test(el)
						? getByName(values)(opts) : getByLogin(isUsersArray)(values)(opts);
			} else {
				return list.item(`Email IsNotNull && (deleted IsNull && (Position Neq Неактивный сотрудник && Position Neq Резерв))`).get(opts);
			}
		})(isUsersArray)(elements),
		create: (elements => opts => {
			const usersToCreate = elements.filter(el => el.uid && el.Title);
			return usersToCreate.length
				? spx(USER_WEB).list(USER_LIST).item(elements).create(opts)
				: new Promise((resolve, reject) => reject(new Error('missing uid or Title')));
		})(elements),
		getByUid: getByUid(isUsersArray)(elements),
		getByLogin: getByLogin(isUsersArray)(elements),
		getByName: getByName(elements),

		update: (elements => async opts => {
			const ids = elements.filter(el => !!el.uid);
			if (ids.length) {
				const users = await spx.user(ids).get({ view: ['ID', 'uid'], groupBy: 'uid' });
				const usersToUpdate = elements.reduce((acc, el) => {
					const userID = users[el.uid] ? users[el.uid].ID : void 0;
					if (userID) {
						el.ID = userID;
						delete el.uid;
						acc.push(el);
					}
					return acc
				}, [])
				return spx(USER_WEB).list(USER_LIST).item(usersToUpdate).update({ ...opts, view: ['ID', 'uid'] });
			} else {
				return new Promise((resolve, reject) => { reject(new Error('missing uid')) });
			}
		})(elements),

		updateSP: (elements => opts => {
			const usersToUpdate = elements.reduce((acc, el) => {
				if (userID) {
					el.ID = el.uid;;
					delete el.uid;
					acc.push(el);
				}
				return acc;
			}, [])
			return spx('/').list(USER_LIST_GUID).item(usersToUpdate).update(opts);
		})(elements),

		deleteWithMissedUid: async opts => {
			const users = await spx(USER_WEB).list(USER_LIST).item('Number uid IsNull').get(opts);
			return spx(USER_WEB).list(USER_LIST).item(users.map(prop('ID'))).delete(opts);
		}
	}
}

user.getSP = async opts => {
	const clientContext = getClientContext('/');
	const user = clientContext.get_web().get_currentUser();
	return prepareResponseJSOM(opts)(await executeJSOM(clientContext)(user)(opts));
}

user.get = async opts => {
	const uid = window._spPageContextInfo ? window._spPageContextInfo.userId : (await user.getSP({ view: 'Id' })).Id;
	return (await spx(USER_WEB).list(USER_LIST).item(`Number uid Eq ${uid}`).get(opts))[0];
}

export default user