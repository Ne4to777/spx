//  ===============================================================================================================
//  =============     =====    ====  =======  ===      ===        =====  =====  =======  ==        ===      =======
//  ============  ===  ===  ==  ===   ======  ==  ====  =====  =======    ====   ======  =====  =====  ====  ======
//  ===========  ========  ====  ==    =====  ==  ====  =====  ======  ==  ===    =====  =====  =====  ====  ======
//  ===========  ========  ====  ==  ==  ===  ===  ==========  =====  ====  ==  ==  ===  =====  ======  ===========
//  ===========  ========  ====  ==  ===  ==  =====  ========  =====  ====  ==  ===  ==  =====  ========  =========
//  ===========  ========  ====  ==  ====  =  =======  ======  =====        ==  ====  =  =====  ==========  =======
//  ===========  ========  ====  ==  =====    ==  ====  =====  =====  ====  ==  =====    =====  =====  ====  ======
//  ============  ===  ===  ==  ===  ======   ==  ====  =====  =====  ====  ==  ======   =====  =====  ====  ======
//  =============     =====    ====  =======  ===      ======  =====  ====  ==  =======  =====  ======      =======
//  ===============================================================================================================

export const REQUEST_TIMEOUT = 3600000;
export const MAX_ITEMS_LIMIT = 100000;
export const REQUEST_BUNDLE_MAX_SIZE = 252;
export const ACTION_TYPES = {
	create: 'created',
	update: 'updated',
	delete: 'deleted',
	recycle: 'recycled',
	get: 'get',
	copy: 'copied',
	move: 'moved',
	restore: 'restored',
	clear: 'cleared'
}
export const IS_DELETE_ACTION = {
	delete: true,
	recycle: true
}

export const CHUNK_MAX_LENGTH = 500;
export const TYPES = {
	list: [{
		name: 'genericList',
		description: 'Custom list'
	}, {
		name: 'documentLibrary',
		description: 'Document library'
	}, {
		name: 'pictureLibrary',
		description: 'Picture library'
	}],
	column: ['Text', 'Note', 'Number', 'Choice', 'DateTime', 'Lookup', 'User'],
	listCodes: {
		100: 'genericList',
		101: 'documentLibrary'
	}
}

export const ACTION_TYPES_TO_UNSET = {
	create: true,
	update: true,
	delete: true,
	recycle: true,
	restore: true
}

export const FILE_LIST_TEMPLATES = {
	101: true,
	109: true,
	110: true,
	113: true,
	114: true,
	116: true,
	119: true,
	121: true,
	122: true,
	123: true,
	175: true,
	851: true,
	10102: true
}

export const LIBRARY_STANDART_COLUMN_NAMES = {
	AppAuthor: true,
	AppEditor: true,
	Author: true,
	CheckedOutTitle: true,
	CheckedOutUserId: true,
	CheckoutUser: true,
	ContentTypeId: true,
	Created: true,
	Created_x0020_By: true,
	Created_x0020_Date: true,
	DocConcurrencyNumber: true,
	Editor: true,
	FSObjType: true,
	FileDirRef: true,
	FileLeafRef: true,
	FileRef: true,
	File_x0020_Size: true,
	File_x0020_Type: true,
	FolderChildCount: true,
	GUID: true,
	HTML_x0020_File_x0020_Type: true,
	ID: true,
	InstanceID: true,
	IsCheckedoutToLocal: true,
	ItemChildCount: true,
	Last_x0020_Modified: true,
	MetaInfo: true,
	Modified: true,
	Modified_x0020_By: true,
	Order: true,
	ParentLeafName: true,
	ParentVersionString: true,
	ProgId: true,
	ScopeId: true,
	SortBehavior: true,
	SyncClientId: true,
	TemplateUrl: true,
	UniqueId: true,
	VirusStatus: true,
	WorkflowInstanceID: true,
	WorkflowVersion: true,
	owshiddenversion: true,
	xd_ProgID: true,
	xd_Signature: true,
	_CheckinComment: true,
	_CopySource: true,
	_HasCopyDestinations: true,
	_IsCurrentVersion: true,
	_Level: true,
	_ModerationComments: true,
	_ModerationStatus: true,
	_SharedFileIndex: true,
	_SourceUrl: true,
	_UIVersion: true,
	_UIVersionString: true,
}

export const ROOT_WEB_DUMMY = '@ROOT_WEB@';

//  ============================================================
//  ================        ==  ====  ==       ===        ======
//  ===================  =====   ==   ==  ====  ==  ============
//  ===================  ======  ==  ===  ====  ==  ============
//  ===================  ======  ==  ===  ====  ==  ============
//  ===================  =======    ====       ===      ========
//  ===================  ========  =====  ========  ============
//  ===================  ========  =====  ========  ============
//  ===================  ========  =====  ========  ============
//  ===================  ========  =====  ========        ======
//  ============================================================

export const typeOf = variable => Object.prototype.toString.call(variable).slice(8, -1).toLowerCase();

//  =================================================================
//  =============     ===  ====  ==       ===       ===  ====  ======
//  ============  ===  ==  ====  ==  ====  ==  ====  ==   ==   ======
//  ===========  ========  ====  ==  ====  ==  ====  ===  ==  =======
//  ===========  ========  ====  ==  ===   ==  ===   ===  ==  =======
//  ===========  ========  ====  ==      ====      ======    ========
//  ===========  ========  ====  ==  ====  ==  ====  =====  =========
//  ===========  ========  ====  ==  ====  ==  ====  =====  =========
//  ============  ===  ==   ==   ==  ====  ==  ====  =====  =========
//  =============     ====      ===  ====  ==  ====  =====  =========
//  =================================================================

/**
 * curry :: ((a,b) → c) → (a → b → c)
 *
 * 2-args currying
 *
 * @param {*} b
 * @param {*} c
 * @returns {Function}
 */
export const curry = f => (x, y) => f(x)(y);

//  ===================================================================================================================
//  =============     =====    ====  =====  ==      ====    ==  =======  =====  =====        ====    ====       =======
//  ============  ===  ===  ==  ===   ===   ==  ===  ====  ===   ======  ====    =======  ======  ==  ===  ====  ======
//  ===========  ========  ====  ==  =   =  ==  ====  ===  ===    =====  ===  ==  ======  =====  ====  ==  ====  ======
//  ===========  ========  ====  ==  == ==  ==  ===  ====  ===  ==  ===  ==  ====  =====  =====  ====  ==  ===   ======
//  ===========  ========  ====  ==  =====  ==      =====  ===  ===  ==  ==  ====  =====  =====  ====  ==      ========
//  ===========  ========  ====  ==  =====  ==  ===  ====  ===  ====  =  ==        =====  =====  ====  ==  ====  ======
//  ===========  ========  ====  ==  =====  ==  ====  ===  ===  =====    ==  ====  =====  =====  ====  ==  ====  ======
//  ============  ===  ===  ==  ===  =====  ==  ===  ====  ===  ======   ==  ====  =====  ======  ==  ===  ====  ======
//  =============     =====    ====  =====  ==      ====    ==  =======  ==  ====  =====  =======    ====  ====  ======
//  ===================================================================================================================

/**
 * I_Combinator :: a → a
 *
 * identity
 *
 * @param {*} a
 * @returns {*} a
 */
export const I_Combinator = _ => _;
export const identity = I_Combinator;

/**
 * K_Combinator :: a → b → a
 *
 * constant
 *
 * @param {*} a
 * @returns {*} a
 */
export const K_Combinator = x => _ => x;
export const constant = K_Combinator;

/**
 * A_Combinator :: (a → b) → a → b
 *
 * apply
 * @param {Function} f
 * @returns {Function} f
 */
export const A_Combinator = f => x => f(x);

/**
 * U_Combinator :: (a → b) → b
 *
 * fix-point
 * @param {Function} f
 * @returns {Function} f
 */
export const U_Combinator = f => f(f);

/**
 * Y_Combinator :: (a → b) → b
 *
 * fix-point
 * @param {Function} f
 * @returns {Function} f
 */
export const Y_Combinator = f => U_Combinator(g => f(x => g(g)(x)));
export const fix = Y_Combinator;

/**
 * C_Combinator :: (a → b → c) → b → a → c
 *
 * flip
 * @param {Function} f
 * @returns {Function} f
 */

export const C_Combinator = f => x => y => f(y)(x);
export const flip = C_Combinator;

/**
 * S_Combinator :: (a → b → c) → (a → b) → a → c
 *
 * substitution
 * @param {Function} f
 * @returns {Function} f
 */
export const S_Combinator = f => g => x => f(x)(g(x));
export const substitution = S_Combinator;

/**
 * S_CombinatorI :: (a → b) → (a → b → c) → a → c
 *
 * inverted S_Combinator
 * @param {Function} f
 * @returns {Function} f
 */
export const S_CombinatorI = f => g => x => g(x)(f(x));
export const substitutionI = S_CombinatorI;

/**
 * S_CombinatorAsync :: (a → b) → (a → b → c) → a → c
 *
 * async S_Combinator
 * @param {Function} f
 * @returns {Function} f
 */
export const S_CombinatorAsync = f => g => async x => f(x)(await g(x));
export const substitutionAsync = S_CombinatorAsync;

/**
 * S_CombinatorIAsync :: (a → b) → (a → b → c) → a → c
 *
 * async S_CombinatorI
 * @param {Function} f
 * @returns {Function} f
 */
export const S_CombinatorIAsync = f => g => async x => g(x)(await f(x));
export const substitutionIAsync = S_CombinatorIAsync;

export const overstep = f => x => (f(x), x);
export const functionSum = f => x => y => x + f(y);

//  =======================================================================================================
//  =============     =====    ====  =======  ==       ===    ==        ==    ====    ====  =======  ======
//  ============  ===  ===  ==  ===   ======  ==  ====  ===  ======  ======  ====  ==  ===   ======  ======
//  ===========  ========  ====  ==    =====  ==  ====  ===  ======  ======  ===  ====  ==    =====  ======
//  ===========  ========  ====  ==  ==  ===  ==  ====  ===  ======  ======  ===  ====  ==  ==  ===  ======
//  ===========  ========  ====  ==  ===  ==  ==  ====  ===  ======  ======  ===  ====  ==  ===  ==  ======
//  ===========  ========  ====  ==  ====  =  ==  ====  ===  ======  ======  ===  ====  ==  ====  =  ======
//  ===========  ========  ====  ==  =====    ==  ====  ===  ======  ======  ===  ====  ==  =====    ======
//  ============  ===  ===  ==  ===  ======   ==  ====  ===  ======  ======  ====  ==  ===  ======   ======
//  =============     =====    ====  =======  ==       ===    =====  =====    ====    ====  =======  ======
//  =======================================================================================================

export const ifThen = predicate => ([onTrue, onFalse]) => x => predicate(x) ? onTrue(x) : onFalse ? onFalse(x) : x;
export const switchCase = condition => cases => x => {
	const caseF = cases[condition(x)];
	return caseF ? caseF(x) : cases.default ? cases.default(x) : void 0
}


//  ===============================================================================
//  ===========  =======  ==  ====  ==  =====  ==      ====        ==       =======
//  ===========   ======  ==  ====  ==   ===   ==  ===  ===  ========  ====  ======
//  ===========    =====  ==  ====  ==  =   =  ==  ====  ==  ========  ====  ======
//  ===========  ==  ===  ==  ====  ==  == ==  ==  ===  ===  ========  ===   ======
//  ===========  ===  ==  ==  ====  ==  =====  ==      ====      ====      ========
//  ===========  ====  =  ==  ====  ==  =====  ==  ===  ===  ========  ====  ======
//  ===========  =====    ==  ====  ==  =====  ==  ====  ==  ========  ====  ======
//  ===========  ======   ==   ==   ==  =====  ==  ===  ===  ========  ====  ======
//  ===========  =======  ===      ===  =====  ==      ====        ==  ====  ======
//  ===============================================================================

export const sum = x => y => x + y;
export const gt = x => y => x < y;

//  ==========================================================================
//  ============      ===        ==       ===    ==  =======  ===      =======
//  ===========  ====  =====  =====  ====  ===  ===   ======  ==   ==   ======
//  ===========  ====  =====  =====  ====  ===  ===    =====  ==  ====  ======
//  ============  ==========  =====  ===   ===  ===  ==  ===  ==  ============
//  ==============  ========  =====      =====  ===  ===  ==  ==  ============
//  ================  ======  =====  ====  ===  ===  ====  =  ==  ===   ======
//  ===========  ====  =====  =====  ====  ===  ===  =====    ==  ====  ======
//  ===========  ====  =====  =====  ====  ===  ===  ======   ==   ==   ======
//  ============      ======  =====  ====  ==    ==  =======  ===      =======
//  ==========================================================================

export const stringTest = re => str => re.test(str);
export const stringReplace = re => to => str => str.replace(re, to);
export const stringMatch = re => str => str.match(re) || [];
export const stringCut = re => stringReplace(re)('');
export const stringSplit = re => str => str.split(re);
export const stringTrim = str => str.trim();

//  =================================================================
//  ==============  =====       ===       ======  =====  ====  ======
//  =============    ====  ====  ==  ====  ====    ====   ==   ======
//  ============  ==  ===  ====  ==  ====  ===  ==  ====  ==  =======
//  ===========  ====  ==  ===   ==  ===   ==  ====  ===  ==  =======
//  ===========  ====  ==      ====      ====  ====  ====    ========
//  ===========        ==  ====  ==  ====  ==        =====  =========
//  ===========  ====  ==  ====  ==  ====  ==  ====  =====  =========
//  ===========  ====  ==  ====  ==  ====  ==  ====  =====  =========
//  ===========  ====  ==  ====  ==  ====  ==  ====  =====  =========
//  =================================================================


export const getArray = x => typeOf(x) === 'array' ? x : x ? [x] : [];
export const getArrayLength = xs => xs.length;
export const map = f => xs => xs.map(f);
export const filter = f => xs => xs.filter(f);
export const slice = (from, to) => xs => xs.slice(from, to);
export const join = delim => xs => xs.join(delim);
export const removeEmpties = filter(x => !!x);
export const concat = array => x => array.concat(x);
export const reduce = f => init => xs => xs.reduce(curry(f), switchCase(typeOf)({
	object: constant({}),
	array: constant([]),
	default: identity
})(init));
export const reduceDirty = f => init => xs => xs.reduce(curry(f), init);
export const flatten = reduce(acc => pipe([ifThen(isArray)([flatten, identity]), concat(acc)]))([]);
export const arrayHead = xs => xs[0];
export const arrayTail = ([h, ...t]) => t;
export const arrayLast = xs => xs[xs.length - 1];
export const arrayInit = slice(0, -1);
export const EMPTY_ARRAY = _ => [];
export const arrayHas = x => pipe([filter(isEqual(x)), isArrayFilled]);
export const chunkArrayFrom = i => size => reduce(acc => x => {
	acc[i] === void 0 && (acc[i] = []);
	const chunk = acc[i];
	chunk.push(x);
	chunk.length === size && i++;
	return acc;
})([])

export const chunkArray = chunkArrayFrom(0);

//  ===========================================================================
//  =============    ====      ========    ==        ====     ===        ======
//  ============  ==  ===  ===  ========  ===  =========  ===  =====  =========
//  ===========  ====  ==  ====  =======  ===  ========  ===========  =========
//  ===========  ====  ==  ===  ========  ===  ========  ===========  =========
//  ===========  ====  ==      =========  ===      ====  ===========  =========
//  ===========  ====  ==  ===  ========  ===  ========  ===========  =========
//  ===========  ====  ==  ====  ==  ===  ===  ========  ===========  =========
//  ============  ==  ===  ===  ===  ===  ===  =========  ===  =====  =========
//  =============    ====      =====     ====        ====     ======  =========
//  ===========================================================================

export const forIn = f => o => {
	for (const prop in o) f(prop)(o[prop]);
	return o
}
export const methodEmpty = m => o => o[m] ? o[m]() : o;
export const method = m => arg => o => o[m] ? o[m](arg) : o;
export const methodI = m => o => arg => o[m] ? o[m](arg) : o;
export const methodIOverstep = m => arg => o => (method(m)(arg)(o), o);
export const apply = m => o => args => o[m].apply(o, args);
export const prop = name => o => o[name];
export const keys = o => Object.keys(o);
export const getInstance = constructor => (...args) => new constructor(...args);
export const getInstanceEmpty = constructor => new constructor();
export const switchProp = props => x => {
	for (const prop in props) {
		if (x[prop]) return props[prop](x);
		if (x.default) return x.default(x)
	}
}
export const NULL = _ => null;
export const climb = f => fix(fr => ([h, ...t]) => o => t.length ? fr(t)(f(h)(o)) : f(h)(o));


//  =========================================================================================================================
//  =============     =====    ====  =====  ==       =====    =====      ===    ==        ==    ====    ====  =======  ======
//  ============  ===  ===  ==  ===   ===   ==  ====  ===  ==  ===  ====  ===  ======  ======  ====  ==  ===   ======  ======
//  ===========  ========  ====  ==  =   =  ==  ====  ==  ====  ==  ====  ===  ======  ======  ===  ====  ==    =====  ======
//  ===========  ========  ====  ==  == ==  ==  ====  ==  ====  ===  ========  ======  ======  ===  ====  ==  ==  ===  ======
//  ===========  ========  ====  ==  =====  ==       ===  ====  =====  ======  ======  ======  ===  ====  ==  ===  ==  ======
//  ===========  ========  ====  ==  =====  ==  ========  ====  =======  ====  ======  ======  ===  ====  ==  ====  =  ======
//  ===========  ========  ====  ==  =====  ==  ========  ====  ==  ====  ===  ======  ======  ===  ====  ==  =====    ======
//  ============  ===  ===  ==  ===  =====  ==  =========  ==  ===  ====  ===  ======  ======  ====  ==  ===  ======   ======
//  =============     =====    ====  =====  ==  ==========    =====      ===    =====  =====    ====    ====  =======  ======
//  =========================================================================================================================

export const pipe = reduce(acc => f => x => f(acc(x)))(identity);
export const pipeAsync = reduce(acc => f => async x => f(await acc(x)))(identity);

//  =============================================================
//  ===========  ==========    =====      ===    ====     =======
//  ===========  =========  ==  ===   ==   ===  ====  ===  ======
//  ===========  ========  ====  ==  ====  ===  ===  ============
//  ===========  ========  ====  ==  =========  ===  ============
//  ===========  ========  ====  ==  =========  ===  ============
//  ===========  ========  ====  ==  ===   ===  ===  ============
//  ===========  ========  ====  ==  ====  ===  ===  ============
//  ===========  =========  ==  ===   ==   ===  ====  ===  ======
//  ===========        ====    =====      ===    ====     =======
//  =============================================================

export const toBoolean = x => !!x;
export const not = x => !x;
export const and = x => y => toBoolean(x && y);
export const or = x => y => toBoolean(x || y);
export const TRUE = _ => true;
export const FALSE = _ => false;
export const andArray = reduce(and)(true);
export const orArray = reduce(or)(false);
export const isEqual = sample => x => x === sample;
export const isNotEqual = pipe([isEqual, not]);
export const isNumber = x => typeOf(x) === 'number';
export const isString = x => typeOf(x) === 'string';
export const isRegExp = x => typeOf(x) === 'regexp';
export const isFunction = x => typeOf(x) === 'function';
export const isIterator = x => typeOf(x) === 'iterator';
export const isArray = x => typeOf(x) === 'array';
export const isObject = x => typeOf(x) === 'object';
export const isGUID = stringTest(/^[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}$/);

//  ========================================================================================================
//  ===========        ==   ==   ==    ===      ===        =====  =====  =======  ====     ===        ======
//  ===========  =========  ==  ====  ===  ====  =====  =======    ====   ======  ===  ===  ==  ============
//  ===========  =========  ==  ====  ===  ====  =====  ======  ==  ===    =====  ==  ========  ============
//  ===========  ==========    =====  ====  ==========  =====  ====  ==  ==  ===  ==  ========  ============
//  ===========      =======  ======  ======  ========  =====  ====  ==  ===  ==  ==  ========      ========
//  ===========  ==========    =====  ========  ======  =====        ==  ====  =  ==  ========  ============
//  ===========  =========  ==  ====  ===  ====  =====  =====  ====  ==  =====    ==  ========  ============
//  ===========  =========  ==  ====  ===  ====  =====  =====  ====  ==  ======   ===  ===  ==  ============
//  ===========        ==  ====  ==    ===      ======  =====  ====  ==  =======  ====     ===        ======
//  ========================================================================================================

export const isNull = x => x === null;
export const isNotNull = pipe([isNull, not]);
export const isUndefined = x => x === void 0;
export const isDefined = pipe([isUndefined, not]);
export const isZero = x => x === 0;
export const isNotZero = pipe([isZero, not]);
export const isNaN = x => x !== x;
export const isNotNaN = pipe([isNaN, not]);
export const isStringEmpty = x => x === '';
export const isStringFilled = pipe([isStringEmpty, not]);
export const isArrayFilled = pipe([prop('length'), toBoolean]);
export const isArrayEmpty = pipe([isArrayFilled, not]);
export const isObjectFilled = pipe([keys, isArrayFilled]);
export const isObjectEmpty = pipe([keys, isArrayEmpty]);
export const isExists = x => isDefined(x) && isNotNull(x);
export const isNotExists = pipe([isExists, not]);
export const isFilled = ifThen(isExists)([
	switchCase(typeOf)({
		number: x => isNotZero(x) && isNotNaN(x),
		string: isStringFilled,
		array: isArrayFilled,
		object: isObjectFilled,
		null: FALSE,
		default: TRUE
	}),
	toBoolean
])
export const isNotFilled = pipe([isFilled, not]);

export const hasOwnProp = method('hasOwnProperty');
export const hasProp = name => o => o[name];
export const isPropExists = name => pipe([prop(name), isExists])
export const isPropFilled = name => pipe([prop(name), isFilled])

//  =============================================
//  ===========        ==       ===  ====  ======
//  ==============  =====  ====  ==   ==   ======
//  ==============  =====  ====  ===  ==  =======
//  ==============  =====  ===   ===  ==  =======
//  ==============  =====      ======    ========
//  ==============  =====  ====  =====  =========
//  ==============  =====  ====  =====  =========
//  ==============  =====  ====  =====  =========
//  ==============  =====  ====  =====  =========
//  =============================================

export const tryCatch = tryer => catcher => data => {
	try { return tryer(data) } catch (err) { return catcher(err)(data) }
}

export const throwError = msg => { throw new Error(msg) };
export const throwCatchedError = err => msg => { throw new Error(msg + '\n' + err) };

//  ========================================================================================
//  =============     =====    ====  =======  ===      =====    ====  ========        ======
//  ============  ===  ===  ==  ===   ======  ==  ====  ===  ==  ===  ========  ============
//  ===========  ========  ====  ==    =====  ==  ====  ==  ====  ==  ========  ============
//  ===========  ========  ====  ==  ==  ===  ===  =======  ====  ==  ========  ============
//  ===========  ========  ====  ==  ===  ==  =====  =====  ====  ==  ========      ========
//  ===========  ========  ====  ==  ====  =  =======  ===  ====  ==  ========  ============
//  ===========  ========  ====  ==  =====    ==  ====  ==  ====  ==  ========  ============
//  ============  ===  ===  ==  ===  ======   ==  ====  ===  ==  ===  ========  ============
//  =============     =====    ====  =======  ===      =====    ====        ==        ======
//  ========================================================================================

export const inspect = overstep(console.log);

export const log = (...args) => {
	console.log('------- Begin');
	args.map(el => console.log(el));
	console.log('------- End');
	return args.length > 1 ? args : args[0]
}

//  =====================================================================================
//  ============      ===       =====    ====  ====  ==       ===        ==       =======
//  ===========   ==   ==  ====  ===  ==  ===  ====  ==  ====  ==  ========  ====  ======
//  ===========  ====  ==  ====  ==  ====  ==  ====  ==  ====  ==  ========  ====  ======
//  ===========  ========  ===   ==  ====  ==  ====  ==  ====  ==  ========  ===   ======
//  ===========  ========      ====  ====  ==  ====  ==       ===      ====      ========
//  ===========  ===   ==  ====  ==  ====  ==  ====  ==  ========  ========  ====  ======
//  ===========  ====  ==  ====  ==  ====  ==  ====  ==  ========  ========  ====  ======
//  ===========   ==   ==  ====  ===  ==  ===   ==   ==  ========  ========  ====  ======
//  ============      ===  ====  ====    =====      ===  ========        ==  ====  ======
//  =====================================================================================

const groupSimple = by => reduce(acc => el => {
	const elValue = el[by];
	const trueValue = elValue === void 0 ? '' : elValue.get_lookupId ? elValue.get_lookupId() : elValue;
	const groupValue = acc[trueValue];
	acc[trueValue] = groupValue === void 0
		? [el] : isArray(groupValue)
			? concat(groupValue)(el) : [groupValue, el];
	return acc
})({})

const mapper = f => fix(fR => acc => switchCase(typeOf)({
	array: f,
	object: el => {
		for (let prop in el) {
			const childEl = el[prop];
			acc[prop] = isArray(childEl) ? f(childEl) : fR({})(childEl)
		}
		return acc;
	},
	default: identity
}))({})

export const grouper = pipe([getArray, flip(pipe([getArray, reduceDirty(flip(pipe([groupSimple, mapper])))]))]);

//  =============================================
//  ===========  ====  ==       ===  ============
//  ===========  ====  ==  ====  ==  ============
//  ===========  ====  ==  ====  ==  ============
//  ===========  ====  ==  ===   ==  ============
//  ===========  ====  ==      ====  ============
//  ===========  ====  ==  ====  ==  ============
//  ===========  ====  ==  ====  ==  ============
//  ===========   ==   ==  ====  ==  ============
//  ============      ===  ====  ==        ======
//  =============================================

export const removeEmptyUrls = filter(x => !!x.Url);
export const removeDuplicatedUrls = pipe([reduce(acc => x => (acc[x.Url] = x, acc))({}), Object.values]);
export const prependSlash = ifThen(stringTest(/^\//))([identity, sum('/')]);
export const popSlash = stringCut(/\/$/);
export const pushSlash = str => str + '/';
export const shiftSlash = stringCut(/^\//);
export const mergeSlashes = stringReplace(/\/\/+/g)('/');
export const urlSplit = stringSplit('/');
export const getTitleFromUrl = pipe([popSlash, urlSplit, arrayLast])
export const urlJoin = join('/');
export const getParentUrl = pipe([popSlash, urlSplit, arrayInit, urlJoin]);
export const getFolderFromUrl = ifThen(stringTest(/\./))([getParentUrl, popSlash]);
export const getFilenameFromUrl = ifThen(stringTest(/\./))([getTitleFromUrl, NULL]);
export const hasUrlTailSlash = stringTest(/\/$/);

export const getListRelativeFolder = webUrl => listUrl => elementUrl => {
	if (elementUrl) {
		if (elementUrl === '/') return '/'
		return shiftSlash(arrayLast(stringSplit('@list@')(stringReplace(listUrl)('@list@')(stringReplace(shiftSlash(webUrl))('@web@')(elementUrl)))))
	}
}

export const getWebRelativeFolder = webUrl => elementUrl => {
	if (elementUrl) {
		if (elementUrl === '/') return '/'
		return shiftSlash(arrayLast(stringSplit('@web@')(stringReplace(shiftSlash(webUrl))('@web@')(elementUrl))))
	}
}

const liftContextUrlType = switchCase(typeOf)({
	object: contextUrl => {
		const newContext = Object.assign({}, contextUrl);
		newContext.Url = shiftSlash(newContext.Url);
		if (!contextUrl.Url && contextUrl.Title) newContext.Url = contextUrl.Title;
		if (!contextUrl.Title && contextUrl.Url) newContext.Title = getTitleFromUrl(contextUrl.Url);
		return newContext
	},
	string: contextUrl => ({
		Url: shiftSlash(contextUrl),
		Title: getTitleFromUrl(contextUrl)
	}),
	default: _ => ({
		Url: ROOT_WEB_DUMMY,
		Title: ''
	})
})

const liftUrlType = switchCase(typeOf)({
	object: item => {
		if (!item.Url) item.Url = item.Folder ? `${item.Folder}/${item.Content.name}` : item.Content.name;
		return item;
	},
	string: item => ({
		Url: item,
		Title: getTitleFromUrl(item),
		ServerRelativeUrl: item
	}),
	number: item => ({
		Url: item,
		Title: item,
		ServerRelativeUrl: item
	}),
	default: _ => ({
		Url: void 0,
		ServerRelativeUrl: void 0
	})
})

const liftListType = switchCase(typeOf)({
	object: list => {
		const newList = Object.assign({}, list);
		if (!list.Title) newList.Title = list.EntityTypeName || list.Url;
		return newList
	},
	string: list => ({
		Url: list,
		Title: list
	}),
	default: _ => ({
		Url: '',
		Title: ''
	})
})

const liftColumnType = switchCase(typeOf)({
	object: column => {
		const newColumn = Object.assign({}, column);
		if (!column.Title) newColumn.Title = column.EntityPropertyName || column.InternalName || column.StaticName;
		if (!column.Type) newColumn.Type = 'Text';
		return newColumn
	},
	string: column => ({
		Title: column,
		Type: 'Text'
	}),
	default: _ => ({
		Title: ''
	})
})

const liftItemType = switchCase(typeOf)({
	object: identity,
	string: item => ({
		ID: item
	}),
	number: item => ({
		ID: item
	}),
	default: _ => ({
		ID: void 0
	})
})

const normalizeContextUrl = pipe([
	liftContextUrlType,
	substitution(o => url => (o.Url = url, o))
		(pipe([
			prop('Url'),
			overstep(pipe([
				mergeSlashes,
				ifThen(isEqual('/'))([
					constant(`${ROOT_WEB_DUMMY}/`),
					shiftSlash
				])
			]))
		]))
])
const normalizeContextUrls = ifThen(isArrayFilled)([
	pipe([
		map(normalizeContextUrl),
		removeEmptyUrls,
		removeDuplicatedUrls
	]),
	constant([liftContextUrlType()])
])

export class ContextUrlBox {
	constructor(value) {
		this.isArray = isArray(value);
		this.value = this.isArray ? normalizeContextUrls(value) : normalizeContextUrl(value);
	}
	async chainAsync(f) {
		return await this.isArray ? Promise.all(map(f)(this.value)) : f(this.value);
	}
	chain(f) {
		return this.isArray ? map(f)(this.value) : f(this.value);
	}
	join() {
		return this.isArray ? join(', ')(map(prop('Url'))(this.value)) : this.value.Url;
	}
	getLength() {
		return this.isArray ? this.value.length : 1;
	}
	getHead() {
		return this.isArray ? this.value[0] : this.value;
	}
}

const normalizeUrl = pipe([
	liftUrlType,
	substitution(o => url => (o.Url = url, o))
		(pipe([
			prop('Url'),
			overstep(ifThen(isExists)([
				pipe([
					mergeSlashes,
					shiftSlash
				])
			]))
		]))
])
const normalizeUrls = ifThen(isArrayFilled)([
	pipe([
		map(normalizeUrl),
		removeEmptyUrls,
		removeDuplicatedUrls
	]),
	constant([liftUrlType()])
])

export class Box {
	constructor(value, type) {
		this.isArray = isArray(value);
		this.prop = 'Url';
		this.joinProp = 'Url';
		this.value = switchCase(identity)({
			list: _ => {
				this.joinProp = 'Title';
				return this.isArray ? map(liftListType)(value) : liftListType(value)
			},
			column: _ => {
				this.prop = 'Title';
				this.joinProp = 'Title';
				return this.isArray ? map(liftColumnType)(value) : liftColumnType(value)
			},
			item: _ => {
				this.prop = 'ID';
				this.joinProp = 'ID';
				return this.isArray ? map(liftItemType)(value) : liftItemType(value)
			},
			folder: _ => {
				this.joinProp = 'ServerRelativeUrl';
				return this.isArray ? normalizeUrls(value) : normalizeUrl(value)
			},
			file: _ => {
				this.joinProp = 'ServerRelativeUrl';
				return this.isArray ? normalizeUrls(value) : normalizeUrl(value)
			},
			default: _ => this.isArray ? normalizeUrls(value) : normalizeUrl(value)
		})(type)
	}
	async chainAsync(f) {
		return await this.isArray ? Promise.all(map(f)(this.value)) : f(this.value);
	}
	chain(f) {
		return this.isArray ? map(f)(this.value) : f(this.value);
	}
	map(f) {
		return new Box(this.isArray ? map(f)(this.value) : f(this.value));
	}
	join() {
		return this.isArray ? join(', ')(map(prop(this.joinProp))(this.value)) : this.value[this.joinProp];
	}
	getValues() {
		return this.isArray ? this.value : [this.value];
	}
	getLength() {
		return this.isArray ? this.value.length : 1;
	}
	getHead() {
		return this.isArray ? this.value[0] : this.value;
	}
}


//  ========================================================================================
//  =============     =====    ====  =======  ==        ==        ==   ==   ==        ======
//  ============  ===  ===  ==  ===   ======  =====  =====  =========  ==  ======  =========
//  ===========  ========  ====  ==    =====  =====  =====  =========  ==  ======  =========
//  ===========  ========  ====  ==  ==  ===  =====  =====  ==========    =======  =========
//  ===========  ========  ====  ==  ===  ==  =====  =====      =======  ========  =========
//  ===========  ========  ====  ==  ====  =  =====  =====  ==========    =======  =========
//  ===========  ========  ====  ==  =====    =====  =====  =========  ==  ======  =========
//  ============  ===  ===  ==  ===  ======   =====  =====  =========  ==  ======  =========
//  =============     =====    ====  =======  =====  =====        ==  ====  =====  =========
//  ========================================================================================

const newClientContext = getInstance(SP.ClientContext);

export const getClientContext = pipe([
	pipe([
		stringReplace(ROOT_WEB_DUMMY)('/'),
		mergeSlashes,
		popSlash,
		prependSlash,
		newClientContext
	]),
	overstep(method('set_requestTimeout')(REQUEST_TIMEOUT))
])

//  ==================================================================================================
//  ===========       ===        ===      ===       =====    ====  =======  ===      ===        ======
//  ===========  ====  ==  ========  ====  ==  ====  ===  ==  ===   ======  ==  ====  ==  ============
//  ===========  ====  ==  ========  ====  ==  ====  ==  ====  ==    =====  ==  ====  ==  ============
//  ===========  ===   ==  =========  =======  ====  ==  ====  ==  ==  ===  ===  =======  ============
//  ===========      ====      =======  =====       ===  ====  ==  ===  ==  =====  =====      ========
//  ===========  ====  ==  =============  ===  ========  ====  ==  ====  =  =======  ===  ============
//  ===========  ====  ==  ========  ====  ==  ========  ====  ==  =====    ==  ====  ==  ============
//  ===========  ====  ==  ========  ====  ==  =========  ==  ===  ======   ==  ====  ==  ============
//  ===========  ====  ==        ===      ===  ==========    ====  =======  ===      ===        ======
//  ==================================================================================================

const getSPObjectValues = asItem => ifThen(isExists)([
	pipe([
		ifThen(constant(asItem))([methodEmpty('get_listItemAllFields')]),
		switchProp({
			get_fieldValues: methodEmpty('get_fieldValues'),
			get_objectData: pipe([methodEmpty('get_objectData'), methodEmpty('get_properties')]),
			default: tryCatch(pipe([JSON.stringify, sum(`Wrong spObject: `), throwError]))(throwError)
		})
	])
])

const getRESTValues = pipe([
	ifThen(hasProp('body'))([
		pipe([
			prop('body'),
			ifThen(isString)([
				JSON.parse,
			])
		]),
		prop('data'),
	]),
	ifThen(hasProp('d'))([
		pipe([
			prop('d'),
			ifThen(hasProp('results'))([
				prop('results')
			])
		])
	])
])


export const prepareResponseJSOM = (opts = {}) =>
	ifThen(isArray)([
		pipe([
			flatten,
			ifThen(constant(opts.expanded))([
				identity,
				pipe([
					map(getSPObjectValues(opts.asItem)),
					ifThen(constant(opts.groupBy))([
						grouper(opts.groupBy)
					])
				])
			])
		]),
		ifThen(constant(opts.expanded))([
			identity,
			getSPObjectValues(opts.asItem)
		])
	])

export const prepareResponseREST = (opts = {}) => ifThen(isArray)([
	pipe([
		flatten,
		pipe([
			ifThen(constant(opts.expanded))([
				identity
			]),
			map(getRESTValues),
			ifThen(constant(opts.groupBy))([
				grouper(opts.groupBy)
			])
		])
	]),
	getRESTValues
])

//  =======================================================
//  ===========  ==========    =======  =====       =======
//  ===========  =========  ==  =====    ====  ====  ======
//  ===========  ========  ====  ===  ==  ===  ====  ======
//  ===========  ========  ====  ==  ====  ==  ====  ======
//  ===========  ========  ====  ==  ====  ==  ====  ======
//  ===========  ========  ====  ==        ==  ====  ======
//  ===========  ========  ====  ==  ====  ==  ====  ======
//  ===========  =========  ==  ===  ====  ==  ====  ======
//  ===========        ====    ====  ====  ==       =======
//  =======================================================

const getViewOption = ifThen(isObjectFilled)([
	ifThen(isPropFilled('view'))([
		ifThen(isPropFilled('groupBy'))([
			opts => concat(getArray(opts.view))(getArray(opts.groupBy)),
			pipe([prop('view'), getArray])
		]),
		EMPTY_ARRAY
	]),
	EMPTY_ARRAY
])

export const load = clientContext => spObject => (opts = {}) => ifThen(hasProp('getEnumerator'))([
	pipe([
		data => pipe([constant(getViewOption(opts)), ifThen(isArrayFilled)([view => [data, `Include(${view})`], constant([data])])])(data),
		apply('loadQuery')(clientContext),
	]),
	overstep(pipe([
		data => pipe([constant(getViewOption(opts)), ifThen(isArrayFilled)([view => [data, view], constant([data])])])(data),
		apply('load')(clientContext)
	]))
])(spObject)

//  ===============================================================================================
//  ===========        ==   ==   ==        ====     ===  ====  ==        ====    ====       =======
//  ===========  =========  ==  ===  =========  ===  ==  ====  =====  ======  ==  ===  ====  ======
//  ===========  =========  ==  ===  ========  ========  ====  =====  =====  ====  ==  ====  ======
//  ===========  ==========    ====  ========  ========  ====  =====  =====  ====  ==  ===   ======
//  ===========      =======  =====      ====  ========  ====  =====  =====  ====  ==      ========
//  ===========  ==========    ====  ========  ========  ====  =====  =====  ====  ==  ====  ======
//  ===========  =========  ==  ===  ========  ========  ====  =====  =====  ====  ==  ====  ======
//  ===========  =========  ==  ===  =========  ===  ==   ==   =====  ======  ==  ===  ====  ======
//  ===========        ==  ====  ==        ====     ====      ======  =======    ====  ====  ======
//  ===============================================================================================

export const executorJSOM = clientContext => (opts = {}) => new Promise((resolve, reject) => {
	clientContext.executeQueryAsync(_ => resolve(), (sender, args) => {
		const { silent, silentErrors } = opts;
		if (!silent && !silentErrors) {
			console.error(`\nMessage: ${
				args.get_message().replace(/\n{1,}/g, ' ')}\nValue: ${
				args.get_errorValue()}\nType: ${
				args.get_errorTypeName()}\nId: ${
				args.get_errorTraceCorrelationId()}`);
		}
		reject(args)
	})
})

export const executeJSOM = clientContext => spObject => async opts => {
	const spObjects = load(clientContext)(spObject)(opts);
	await executorJSOM(clientContext)(opts);
	return spObjects
}

export const executorREST = contextUrl => (opts = {}) => pipe([
	mergeSlashes,
	popSlash,
	prependSlash,
	getInstance(SP.RequestExecutor),
	executor => new Promise((resolve, reject) => executor.executeAsync({
		...opts,
		method: pipe([prop('method'), ifThen(stringTest(/post/i))([constant('POST'), constant('GET')])])(opts),
		success: resolve,
		error: res => {
			const { silent, silentErrors } = opts;
			if (!silent && !silentErrors) {
				const body = res.body;
				let msg = body;
				if (typeOf(res.body) === 'string') {
					try {
						msg = JSON.parse(res.body).error.message.value
					} catch (err) { }
				}
				console.error(`\nMessage: ${
					res.statusText}\nCode: ${
					res.statusCode}\nValue: ${
					msg}`)
			}
			reject(res)
		}
	}))
])(contextUrl)

//  =============================================================================
//  ===========      =======  ======      ===        ===       ========  ========
//  ===========  ===  =====    ====  ====  ==  ========  =====  ======   ========
//  ===========  ====  ===  ==  ===  ====  ==  ========  ============    ========
//  ===========  ===  ===  ====  ===  =======  ========       ======  =  ========
//  ===========      ====  ====  =====  =====      ====   ===  ====  ==  ========
//  ===========  ===  ===        =======  ===  ========  =====  ==  ===  ========
//  ===========  ====  ==  ====  ==  ====  ==  ========  =====  ==         ======
//  ===========  ===  ===  ====  ==  ====  ==  =========  ===   =======  ========
//  ===========      ====  ====  ===      ===        ====     =========  ========
//  =============================================================================

export const convertFileContent = switchCase(typeOf)({
	arraybuffer: pipe([
		getInstance(Uint8Array),
		reduce(functionSum(String.fromCharCode))(''),
		btoa
	]),
	default: tryCatch(btoa)(err => identity)
})

//  ===============================================================================================
//  ============      ===       =====    ====      ========    ==        ====     ===        ======
//  ===========  ====  ==  ====  ===  ==  ===  ===  ========  ===  =========  ===  =====  =========
//  ===========  ====  ==  ====  ==  ====  ==  ====  =======  ===  ========  ===========  =========
//  ============  =======  ====  ==  ====  ==  ===  ========  ===  ========  ===========  =========
//  ==============  =====       ===  ====  ==      =========  ===      ====  ===========  =========
//  ================  ===  ========  ====  ==  ===  ========  ===  ========  ===========  =========
//  ===========  ====  ==  ========  ====  ==  ====  ==  ===  ===  ========  ===========  =========
//  ===========  ====  ==  =========  ==  ===  ===  ===  ===  ===  =========  ===  =====  =========
//  ============      ===  ==========    ====      =====     ====        ====     ======  =========
//  ===============================================================================================

const setItemSP = name => item => value => item.set_item(name, value);

export const getSPFolderByUrl = url => ifThen(constant(url))([
	climb(
		url => pipe([
			methodEmpty('get_folders'),
			method('getByUrl')(url)
		]))
		(pipe([
			stringReplace(/\/\/+/)('/'),
			stringSplit('/')
		])(url)),
	identity
])


export const setItem = fieldsInfo => fields => spObject => {
	for (const prop in fields) {
		const fieldValues = fields[prop];
		const fieldInfoArray = fieldsInfo[prop];
		if (isArray(fieldInfoArray)) {
			const fieldInfo = fieldInfoArray[0];
			const set = setItemSP(fieldInfo.InternalName)(spObject);
			const setLookupAndUser = f => constructor => pipe([f(constructor), set]);
			switch (fieldInfo.TypeAsString) {
				case 'Lookup': setLookupAndUser(setLookup)(SP.FieldLookupValue)(fieldValues); break;
				case 'LookupMulti': setLookupAndUser(setLookupMulti)(SP.FieldLookupValue)(getArray(fieldValues)); break;
				case 'User': setLookupAndUser(setLookup)(SP.FieldUserValue)(fieldValues); break;
				case 'UserMulti': setLookupAndUser(setLookupMulti)(SP.FieldUserValue)(getArray(fieldValues)); break;
				default: set(fieldValues);
			}
		}
	}
	spObject.update();
	return spObject
}

export const setLookupMulti = constructor => pipe([
	reduce(acc => ifThen(isExists)([
		pipe([
			setLookup(constructor),
			concat(acc)
		])
	]))([]),
	ifThen(isArrayFilled)([
		identity,
		NULL
	])
])

const setLookupValue = constructor => value => pipe([
	getInstanceEmpty,
	overstep(method('set_lookupId')(value)),
])(constructor);

export const setLookup = constructor => pipe([
	ifThen(isNull)([
		NULL,
		ifThen(hasProp('get_lookupId'))([
			pipe([
				methodEmpty('get_lookupId'),
				setLookupValue(constructor)
			]),
			pipe([
				parseInt,
				ifThen(constant(and(isNumber)(gt(0))))([
					setLookupValue(constructor),
					NULL
				])
			])
		])
	])
])



export const setFields = source => target => {
	for (let prop in source) {
		const value = source[prop];
		if (value !== void 0 && target[prop]) target[prop](value)
	}
	return target
}

export const getContext = methodEmpty('get_context');

export const getWeb = methodEmpty('get_web');


//  =======================================================
//  ===========        ==        ===      ===        ======
//  ==============  =====  ========  ====  =====  =========
//  ==============  =====  ========  ====  =====  =========
//  ==============  =====  =========  ==========  =========
//  ==============  =====      =======  ========  =========
//  ==============  =====  =============  ======  =========
//  ==============  =====  ========  ====  =====  =========
//  ==============  =====  ========  ====  =====  =========
//  ==============  =====        ===      ======  =========
//  =======================================================

export const assert = msg => bool => console.assert(bool === true, msg);

export const assertProp = prop => o => assert(`object has no property "${prop}"`)(hasOwnProp(prop)(o))

export const assertProps = props => o => {
	for (const prop of props) assertProp(prop)(o)
	return o
}

export const assertObject = props => name => async promise => {
	const el = await promise;
	assert(`${name} is not an object`)(isObject(el));
	assert(`${name} is empty object`)(isObjectFilled(el));
	assertProps(props)(el);
	return el
}

export const assertCollection = props => name => async promise => {
	const el = await promise;
	assert(`${name} collection is not an array`)(isArray(el));
	assert(`${name} collection is an empty array`)(isArrayFilled(el));
	map(pipe([isObjectFilled, assert(`${name} collection element is empty`)]))(el);
	assertProps(props)(el[0]);
	return el
}

export const testIsOk = name => _ => console.log(`${name} is OK`);