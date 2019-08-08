import axios from 'axios'
//  ================================================================================================
//  =======     ====    ===  =======  ==      ==        ====  ====  =======  =        ==      ======
//  ======  ===  ==  ==  ==   ======  =  ====  ====  ======    ===   ======  ====  ====  ====  =====
//  =====  =======  ====  =    =====  =  ====  ====  =====  ==  ==    =====  ====  ====  ====  =====
//  =====  =======  ====  =  ==  ===  ==  =========  ====  ====  =  ==  ===  ====  =====  ==========
//  =====  =======  ====  =  ===  ==  ====  =======  ====  ====  =  ===  ==  ====  =======  ========
//  =====  =======  ====  =  ====  =  ======  =====  ====        =  ====  =  ====  =========  ======
//  =====  =======  ====  =  =====    =  ====  ====  ====  ====  =  =====    ====  ====  ====  =====
//  ======  ===  ==  ==  ==  ======   =  ====  ====  ====  ====  =  ======   ====  ====  ====  =====
//  =======     ====    ===  =======  ==      =====  ====  ====  =  =======  ====  =====      ======
//  ================================================================================================

export const REQUEST_TIMEOUT = 3600000
export const MAX_ITEMS_LIMIT = 100000
export const REQUEST_BUNDLE_MAX_SIZE = 252
export const REQUEST_LIST_FOLDER_CREATE_BUNDLE_MAX_SIZE = 126
export const REQUEST_LIST_FOLDER_UPDATE_BUNDLE_MAX_SIZE = 82
export const REQUEST_LIST_FOLDER_DELETE_BUNDLE_MAX_SIZE = 240
export const CACHE_RETRIES_LIMIT = 3
export const ACTION_TYPES = {
	create: 'created',
	update: 'updated',
	delete: 'deleted',
	recycle: 'recycled',
	get: 'get',
	copy: 'copied',
	move: 'moved',
	restore: 'restored',
	clear: 'cleared',
	erase: 'erased',
	send: 'sent'
}
export const IS_DELETE_ACTION = {
	delete: true,
	recycle: true
}

export const CHUNK_MAX_LENGTH = 500
export const TYPES = {
	list: [
		{
			name: 'genericList',
			description: 'Custom list'
		},
		{
			name: 'documentLibrary',
			description: 'Document library'
		},
		{
			name: 'pictureLibrary',
			description: 'Picture library'
		}
	],
	column: ['Text', 'Note', 'Number', 'Choice', 'DateTime', 'Lookup', 'User'],
	listCodes: {
		100: 'genericList',
		101: 'documentLibrary'
	}
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

	ImageHeight: true,
	ImageWidth: true,
	PreviewExists: true,
	ThumbnailExists: true
}

export const ROOT_WEB_DUMMY = '@ROOT_WEB@'

//  =============================================
//  =====        =  ====  =       ==        =====
//  ========  ====   ==   =  ====  =  ===========
//  ========  =====  ==  ==  ====  =  ===========
//  ========  =====  ==  ==  ====  =  ===========
//  ========  ======    ===       ==      =======
//  ========  =======  ====  =======  ===========
//  ========  =======  ====  =======  ===========
//  ========  =======  ====  =======  ===========
//  ========  =======  ====  =======        =====
//  =============================================

export const typeOf = (variable) => Object.prototype.toString.call(variable).slice(8, -1).toLowerCase()

//  ======================================================
//  =======     ==  ====  =       ==       ==  ====  =====
//  ======  ===  =  ====  =  ====  =  ====  =   ==   =====
//  =====  =======  ====  =  ====  =  ====  ==  ==  ======
//  =====  =======  ====  =  ===   =  ===   ==  ==  ======
//  =====  =======  ====  =      ===      =====    =======
//  =====  =======  ====  =  ====  =  ====  ====  ========
//  =====  =======  ====  =  ====  =  ====  ====  ========
//  ======  ===  =   ==   =  ====  =  ====  ====  ========
//  =======     ===      ==  ====  =  ====  ====  ========
//  ======================================================

/**
 * curry :: ((a,b) → c) → (a → b → c)
 *
 * 2-args currying
 *
 * @param {*} b
 * @param {*} c
 * @returns {Function}
 */
export const curry = (f) => (x, y) => f(x)(y)

//  ===================================================================================================
//  =======     ====    ===  =====  =      ===    =  =======  ====  ====        ===    ===       ======
//  ======  ===  ==  ==  ==   ===   =  ===  ===  ==   ======  ===    ======  =====  ==  ==  ====  =====
//  =====  =======  ====  =  =   =  =  ====  ==  ==    =====  ==  ==  =====  ====  ====  =  ====  =====
//  =====  =======  ====  =  == ==  =  ===  ===  ==  ==  ===  =  ====  ====  ====  ====  =  ===   =====
//  =====  =======  ====  =  =====  =      ====  ==  ===  ==  =  ====  ====  ====  ====  =      =======
//  =====  =======  ====  =  =====  =  ===  ===  ==  ====  =  =        ====  ====  ====  =  ====  =====
//  =====  =======  ====  =  =====  =  ====  ==  ==  =====    =  ====  ====  ====  ====  =  ====  =====
//  ======  ===  ==  ==  ==  =====  =  ===  ===  ==  ======   =  ====  ====  =====  ==  ==  ====  =====
//  =======     ====    ===  =====  =      ===    =  =======  =  ====  ====  ======    ===  ====  =====
//  ===================================================================================================

const COMBINATOR = {
	/**
	 * I :: a → a
	 *
	 * identity
	 *
	 * @param {*} a
	 * @returns {*} a
	 */
	I: x => x,
	/**
	 * K :: a → b → a
	 *
	 * constant
	 *
	 * @param {*} a
	 * @returns {*} a
	 */
	K: x => () => x,
	/**
	* A :: (a → b) → a → b
	*
	* apply
	* @param {Function} f
	* @returns {Function} f
	*/
	A: (f) => (x) => f(x),
	/**
	* U :: (a → a) → a
	*
	* universal
	* @param {Function} f
	* @returns {Function} f
	*/
	U: (f) => f(f),
	/**
	* Y :: (a → a) → a
	*
	* fixed-point
	* @param {Function} f
	* @returns {Function} f
	*/
	Y: (f) => COMBINATOR.U((g) => f((x) => g(g)(x))),
	/**
	* C :: (a → b → c) → b → a → c
	*
	* flip
	* @param {Function} f
	* @returns {Function} f
	*/
	C: (f) => (x) => (y) => f(y)(x),
	/**
	* S :: (a → b → c) → (a → b) → a → c
	*
	* substitution
	* @param {Function} f
	* @returns {Function} f
	*/
	S: (f) => (g) => (x) => f(x)(g(x)),
	/**
	* SI :: (a → b) → (a → b → c) → a → c
	*
	* inverted S
	* @param {Function} f
	* @returns {Function} f
	*/
	SI: (f) => (g) => (x) => g(x)(f(x)),
	/**
	* SA :: (a → b) → (a → b → c) → a → c
	*
	* async S
	* @param {Function} f
	* @returns {Function} f
	*/
	SA: (f) => (g) => async (x) => f(x)(await g(x)),
	/**
	* SIA :: (a → b) → (a → b → c) → a → c
	*
	* async SI
	* @param {Function} f
	* @returns {Function} f
	*/
	SIA: (f) => (g) => async (x) => g(x)(await f(x))
}

export const identity = COMBINATOR.I
export const constant = COMBINATOR.K
export const universal = COMBINATOR.U
export const fix = COMBINATOR.Y
export const flip = COMBINATOR.C
export const substitution = COMBINATOR.S
export const substitutionI = COMBINATOR.SI
export const substitutionAsync = COMBINATOR.SA
export const substitutionIAsync = COMBINATOR.SIA

export const overstep = (f) => (x) => {
	f(x)
	return x
}
export const functionSum = (f) => (x) => (y) => x + f(y)

//  ========================================================================================
//  =======     ====    ===  =======  =       ==    =        =    ===    ===  =======  =====
//  ======  ===  ==  ==  ==   ======  =  ====  ==  =====  =====  ===  ==  ==   ======  =====
//  =====  =======  ====  =    =====  =  ====  ==  =====  =====  ==  ====  =    =====  =====
//  =====  =======  ====  =  ==  ===  =  ====  ==  =====  =====  ==  ====  =  ==  ===  =====
//  =====  =======  ====  =  ===  ==  =  ====  ==  =====  =====  ==  ====  =  ===  ==  =====
//  =====  =======  ====  =  ====  =  =  ====  ==  =====  =====  ==  ====  =  ====  =  =====
//  =====  =======  ====  =  =====    =  ====  ==  =====  =====  ==  ====  =  =====    =====
//  ======  ===  ==  ==  ==  ======   =  ====  ==  =====  =====  ===  ==  ==  ======   =====
//  =======     ====    ===  =======  =       ==    ====  ====    ===    ===  =======  =====
//  ========================================================================================

export const ifThen = (predicate) => ([onTrue, onFalse]) => (x) => predicate(x)
	? onTrue(x)
	: (onFalse)
		? onFalse(x)
		: x

export const switchCase = (condition) => (cases) => (x) => {
	const caseF = cases[condition(x)]
	return caseF ? caseF(x) : cases.default ? cases.default(x) : undefined
}

export const switchType = switchCase(typeOf)

//  ===================================================================
//  =====  =======  =  ====  =  =====  =      ===        =       ======
//  =====   ======  =  ====  =   ===   =  ===  ==  =======  ====  =====
//  =====    =====  =  ====  =  =   =  =  ====  =  =======  ====  =====
//  =====  ==  ===  =  ====  =  == ==  =  ===  ==  =======  ===   =====
//  =====  ===  ==  =  ====  =  =====  =      ===      ===      =======
//  =====  ====  =  =  ====  =  =====  =  ===  ==  =======  ====  =====
//  =====  =====    =  ====  =  =====  =  ====  =  =======  ====  =====
//  =====  ======   =   ==   =  =====  =  ===  ==  =======  ====  =====
//  =====  =======  ==      ==  =====  =      ===        =  ====  =====
//  ===================================================================

export const sum = (x) => (y) => x + y
export const gt = (x) => (y) => x < y

//  ==============================================================
//  ======      ==        =       ==    =  =======  ==      ======
//  =====  ====  ====  ====  ====  ==  ==   ======  =   ==   =====
//  =====  ====  ====  ====  ====  ==  ==    =====  =  ====  =====
//  ======  =========  ====  ===   ==  ==  ==  ===  =  ===========
//  ========  =======  ====      ====  ==  ===  ==  =  ===========
//  ==========  =====  ====  ====  ==  ==  ====  =  =  ===   =====
//  =====  ====  ====  ====  ====  ==  ==  =====    =  ====  =====
//  =====  ====  ====  ====  ====  ==  ==  ======   =   ==   =====
//  ======      =====  ====  ====  =    =  =======  ==      ======
//  ==============================================================

export const stringTest = (re) => (str) => re.test(str)
export const stringReplace = (re) => (to) => (str) => str.replace(re, to)
export const stringMatch = (re) => (str) => str.match(re) || []
export const stringCut = (re) => stringReplace(re)('')
export const stringSplit = (re) => (str) => str.split(re)
export const stringTrim = (str) => str.trim()

//  ======================================================
//  ========  ====       ==       =====  ====  ====  =====
//  =======    ===  ====  =  ====  ===    ===   ==   =====
//  ======  ==  ==  ====  =  ====  ==  ==  ===  ==  ======
//  =====  ====  =  ===   =  ===   =  ====  ==  ==  ======
//  =====  ====  =      ===      ===  ====  ===    =======
//  =====        =  ====  =  ====  =        ====  ========
//  =====  ====  =  ====  =  ====  =  ====  ====  ========
//  =====  ====  =  ====  =  ====  =  ====  ====  ========
//  =====  ====  =  ====  =  ====  =  ====  ====  ========
//  ======================================================

export const getArray = (x) => (typeOf(x) === 'array' ? x : x ? [x] : [])
export const getArrayLength = (xs) => xs.length
export const map = (f) => (xs) => xs.map(f)
export const filter = (f) => (xs) => xs.filter(f)
export const slice = (from, to) => (xs) => xs.slice(from, to)
export const join = (delim) => (xs) => xs.join(delim)
export const removeEmpties = filter((x) => !!x)
export const removeUndefineds = filter((x) => x !== undefined)
export const concat = (array) => (x) => array.concat(x)
export const reduce = (f) => (init) => (xs) => xs.reduce(
	curry(f),
	switchType({
		object: constant({}),
		array: constant([]),
		default: identity
	})(init)
)
export const reduceDirty = (f) => (init) => (xs) => getArray(xs).reduce(curry(flip(f)), getArray(init))
export const flatten = reduce((acc) => pipe([ifThen(isArray)([flatten, identity]), concat(acc)]))([])
export const arrayHead = (xs) => xs[0]
export const arrayTail = ([, ...t]) => t
export const arrayLast = (xs) => xs[xs.length - 1]
export const arrayInit = slice(0, -1)
export const EMPTY_ARRAY = () => []
export const arrayHas = (x) => pipe([filter(isEqual(x)), isArrayFilled])
export const chunkArrayFrom = (start = 0) => (size) => (value) => {
	let i = start
	return reduce((acc) => (x) => {
		if (acc[i] === undefined) acc[i] = []
		const chunk = acc[i]
		chunk.push(x)
		if (chunk.length === size) i += 1
		return acc
	})([])(value)
}
export const chunkArray = chunkArrayFrom()

export const removeEmptiesByProp = property => filter((x) => !!x[property])
export const removeDuplicatedProp = property => pipe([reduce((acc) => (x) => {
	acc[x[property]] = x
	return acc
})({}), Object.values])

//  ===============================================================
//  =======    ===      =======    =        ===     ==        =====
//  ======  ==  ==  ===  =======  ==  ========  ===  ====  ========
//  =====  ====  =  ====  ======  ==  =======  ==========  ========
//  =====  ====  =  ===  =======  ==  =======  ==========  ========
//  =====  ====  =      ========  ==      ===  ==========  ========
//  =====  ====  =  ===  =======  ==  =======  ==========  ========
//  =====  ====  =  ====  =  ===  ==  =======  ==========  ========
//  ======  ==  ==  ===  ==  ===  ==  ========  ===  ====  ========
//  =======    ===      ====     ===        ===     =====  ========
//  ===============================================================

export const forIn = (f) => (o) => {
	const props = Reflect.ownKeys(o)
	for (let i = 0; i < props.length; i += 1) {
		const prop = props[i]
		f(prop)(o[prop])
	}
	return o
}
export const methodEmpty = (m) => (o) => (o[m] ? o[m]() : o)
export const method = (m) => (arg) => (o) => (o[m] ? o[m](arg) : o)
export const methodI = (m) => (o) => (arg) => (o[m] ? o[m](arg) : o)
export const methodIOverstep = (m) => (arg) => (o) => {
	method(m)(arg)(o)
	return o
}
export const apply = (m) => (o) => (args) => o[m](...args)
export const prop = (name) => (o) => o[name]
export const keys = (o) => Object.keys(o)
export const getInstance = (constructor) => (...args) => new constructor(...args)
export const getInstanceEmpty = (constructor) => new constructor()
export const switchProp = (o) => (x) => {
	const props = Reflect.ownKeys(o)
	for (let i = 0; i < props.length; i += 1) {
		const property = props[i]
		if (x[property]) return o[property](x)
		if (x.default) return x.default(x)
	}
	return undefined
}
export const NULL = () => null
export const climb = (f) => fix((fr) => ([h, ...t]) => (o) => (t.length ? fr(t)(f(h)(o)) : f(h)(o)))

//  =========================================================================
//  =======     ====    ===  =====  =       ====    ====      ==        =====
//  ======  ===  ==  ==  ==   ===   =  ====  ==  ==  ==  ====  =  ===========
//  =====  =======  ====  =  =   =  =  ====  =  ====  =  ====  =  ===========
//  =====  =======  ====  =  == ==  =  ====  =  ====  ==  ======  ===========
//  =====  =======  ====  =  =====  =       ==  ====  ====  ====      =======
//  =====  =======  ====  =  =====  =  =======  ====  ======  ==  ===========
//  =====  =======  ====  =  =====  =  =======  ====  =  ====  =  ===========
//  ======  ===  ==  ==  ==  =====  =  ========  ==  ==  ====  =  ===========
//  =======     ====    ===  =====  =  =========    ====      ==        =====
//  =========================================================================

export const pipe = reduce((acc) => (f) => (x) => f(acc(x)))(identity)
export const pipeAsync = reduce((acc) => (f) => async (x) => f(await acc(x)))(identity)

//  ==================================================
//  =====  =========    ====      ==    ===     ======
//  =====  ========  ==  ==   ==   ==  ===  ===  =====
//  =====  =======  ====  =  ====  ==  ==  ===========
//  =====  =======  ====  =  ========  ==  ===========
//  =====  =======  ====  =  ========  ==  ===========
//  =====  =======  ====  =  ===   ==  ==  ===========
//  =====  =======  ====  =  ====  ==  ==  ===========
//  =====  ========  ==  ==   ==   ==  ===  ===  =====
//  =====        ===    ====      ==    ===     ======
//  ==================================================

export const toBoolean = (x) => !!x
export const not = (x) => !x
export const and = (x) => (y) => toBoolean(x && y)
export const or = (x) => (y) => toBoolean(x || y)
export const TRUE = () => true
export const FALSE = () => false
export const andArray = reduce(and)(true)
export const orArray = reduce(or)(false)
export const isEqual = (sample) => (x) => x === sample
export const isNotEqual = pipe([isEqual, not])
export const isNumber = (x) => typeOf(x) === 'number'
export const isString = (x) => typeOf(x) === 'string'
export const isRegExp = (x) => typeOf(x) === 'regexp'
export const isFunction = (x) => typeOf(x) === 'function'
export const isIterator = (x) => typeOf(x) === 'iterator'
export const isArray = (x) => typeOf(x) === 'array'
export const isObject = (x) => typeOf(x) === 'object'
export const isBlob = (x) => {
	const type = typeOf(x)
	return type === 'blob' || type === 'file'
}
export const isGUID = stringTest(/^[a-z0-9]{8}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{4}-[a-z0-9]{12}$/)

//  =========================================================================================
//  =====        =   ==   =    ==      ==        ====  ====  =======  ===     ==        =====
//  =====  ========  ==  ===  ==  ====  ====  ======    ===   ======  ==  ===  =  ===========
//  =====  ========  ==  ===  ==  ====  ====  =====  ==  ==    =====  =  =======  ===========
//  =====  =========    ====  ===  =========  ====  ====  =  ==  ===  =  =======  ===========
//  =====      ======  =====  =====  =======  ====  ====  =  ===  ==  =  =======      =======
//  =====  =========    ====  =======  =====  ====        =  ====  =  =  =======  ===========
//  =====  ========  ==  ===  ==  ====  ====  ====  ====  =  =====    =  =======  ===========
//  =====  ========  ==  ===  ==  ====  ====  ====  ====  =  ======   ==  ===  =  ===========
//  =====        =  ====  =    ==      =====  ====  ====  =  =======  ===     ==        =====
//  =========================================================================================

export const isNull = (x) => x === null
export const isNotNull = pipe([isNull, not])
export const isUndefined = (x) => x === undefined
export const isDefined = pipe([isUndefined, not])
export const isZero = (x) => x === 0
export const isNotZero = pipe([isZero, not])
export const isNaN = (x) => typeOf(x) === 'number' && x.toString() === 'NaN'
export const isNotNaN = pipe([isNaN, not])
export const isNumberFilled = (x) => isNumber(x) && isNotZero(x) && isNotNaN(x)
export const isStringEmpty = (x) => x === ''
export const isStringFilled = pipe([isStringEmpty, not])
export const isArrayFilled = pipe([filter(isDefined), prop('length'), toBoolean])
export const isArrayEmpty = pipe([isArrayFilled, not])
export const isObjectFilled = ifThen(isObject)([pipe([keys, isArrayFilled]), FALSE])
export const isObjectEmpty = pipe([keys, isArrayEmpty])
export const isExists = (x) => isDefined(x) && isNotNull(x)
export const isNotExists = pipe([isExists, not])
export const isFilled = ifThen(isExists)([
	switchType({
		number: isNumberFilled,
		string: isStringFilled,
		array: isArrayFilled,
		object: isObjectFilled,
		null: FALSE,
		default: TRUE
	}),
	toBoolean
])
export const isNotFilled = pipe([isFilled, not])

export const hasOwnProp = method('hasOwnProperty')
export const hasProp = (name) => (o) => o[name]
export const isPropExists = (name) => pipe([prop(name), isExists])
export const isPropFilled = (name) => pipe([prop(name), isFilled])

//  ====================================
//  =====        =       ==  ====  =====
//  ========  ====  ====  =   ==   =====
//  ========  ====  ====  ==  ==  ======
//  ========  ====  ===   ==  ==  ======
//  ========  ====      =====    =======
//  ========  ====  ====  ====  ========
//  ========  ====  ====  ====  ========
//  ========  ====  ====  ====  ========
//  ========  ====  ====  ====  ========
//  ====================================

export const tryCatch = (tryer) => (catcher) => (data) => {
	try {
		return tryer(data)
	} catch (err) {
		return catcher(err)(data)
	}
}

export const throwError = (msg) => {
	throw new Error(msg)
}
export const throwCatchedError = (err) => (msg) => {
	throw new Error(`${msg}\n${err}`)
}

//  ===========================================================================
//  =======     ====    ===  =======  ==      ====    ===  =======        =====
//  ======  ===  ==  ==  ==   ======  =  ====  ==  ==  ==  =======  ===========
//  =====  =======  ====  =    =====  =  ====  =  ====  =  =======  ===========
//  =====  =======  ====  =  ==  ===  ==  ======  ====  =  =======  ===========
//  =====  =======  ====  =  ===  ==  ====  ====  ====  =  =======      =======
//  =====  =======  ====  =  ====  =  ======  ==  ====  =  =======  ===========
//  =====  =======  ====  =  =====    =  ====  =  ====  =  =======  ===========
//  ======  ===  ==  ==  ==  ======   =  ====  ==  ==  ==  =======  ===========
//  =======     ====    ===  =======  ==      ====    ===        =        =====
//  ===========================================================================

export const inspect = overstep(console.log)

export const log = (...args) => {
	console.log('------- Begin')
	args.map((el) => console.log(el))
	console.log('------- End')
	return args.length > 1 ? args : args[0]
}

export const report = (msg, opts = {}) => {
	if (!opts.silent && !opts.silentInfo) console.log(msg)
}

export const rootReport = (actionType, opts = {}) => {
	const {
		box,
		name,
		detailed
	} = opts
	const count = box.getCount()
	report(
		`${ACTION_TYPES[actionType]} ${count} ${name}${count > 1 ? 's' : ''}${detailed ? `: ${box.join()}` : ''} `,
		opts
	)
}

export const webReport = (actionType, opts = {}) => {
	const {
		box,
		name,
		detailed,
		contextUrl
	} = opts
	const count = box.getCount(actionType)
	report(
		`${ACTION_TYPES[actionType]
		} ${count} ${name}${count > 1 ? 's' : ''} at ${contextUrl || '/'
		}${detailed
			? `: ${box.join()}`
			: ''}`,
		opts
	)
}

export const listReport = (actionType, opts = {}) => {
	const {
		box,
		name,
		detailed,
		contextUrl,
		listUrl
	} = opts
	const count = box.getCount(actionType)
	report(
		`${ACTION_TYPES[actionType]
		} ${count
		} ${name}${count > 1 ? 's' : ''} in ${listUrl
		} at ${contextUrl || '/'
		}${detailed
			? `: ${box.join()}`
			: ''}`,
		opts
	)
}

//  ========================================================================
//  ======      ==       ====    ===  ====  =       ==        =       ======
//  =====   ==   =  ====  ==  ==  ==  ====  =  ====  =  =======  ====  =====
//  =====  ====  =  ====  =  ====  =  ====  =  ====  =  =======  ====  =====
//  =====  =======  ===   =  ====  =  ====  =  ====  =  =======  ===   =====
//  =====  =======      ===  ====  =  ====  =       ==      ===      =======
//  =====  ===   =  ====  =  ====  =  ====  =  =======  =======  ====  =====
//  =====  ====  =  ====  =  ====  =  ====  =  =======  =======  ====  =====
//  =====   ==   =  ====  ==  ==  ==   ==   =  =======  =======  ====  =====
//  ======      ==  ====  ===    ====      ==  =======        =  ====  =====
//  ========================================================================

const groupSimple = (by) => reduce((acc) => (el) => {
	const elValue = el[by]
	const trueValue = isExists(elValue) && elValue.get_lookupId ? elValue.get_lookupId() : elValue
	const groupValue = acc[trueValue]
	acc[trueValue] = isUndefined(groupValue)
		? [el]
		: isArray(groupValue)
			? concat(groupValue)(el)
			: [groupValue, el]
	return acc
})({})

const groupMapper = (f) => fix((fR) => (acc) => switchType({
	array: f,
	object: (el) => {
		const props = Reflect.ownKeys(el)
		for (let i = 0; i < props.length; i += 1) {
			const property = props[i]
			const childEl = el[property]
			acc[property] = isArray(childEl)
				? f(childEl)
				: fR({})(childEl)
		}
		return acc
	},
	default: identity
}))({})

// const grouper = pipe([getArray, flip(pipe([getArray, reduceDirty(pipe([groupSimple, groupMapper]))]))]);

const grouper = flip(reduceDirty(pipe([groupSimple, groupMapper])))

const mapper = (by) => (xs) => reduce((acc) => (el) => {
	const elValue = el[by]
	acc[isExists(elValue) && elValue.get_lookupId ? elValue.get_lookupId() : elValue] = el
	return acc
})({})(getArray(xs))

//  ====================================
//  =====  ====  =       ==  ===========
//  =====  ====  =  ====  =  ===========
//  =====  ====  =  ====  =  ===========
//  =====  ====  =  ===   =  ===========
//  =====  ====  =      ===  ===========
//  =====  ====  =  ====  =  ===========
//  =====  ====  =  ====  =  ===========
//  =====   ==   =  ====  =  ===========
//  ======      ==  ====  =        =====
//  ====================================

export const hasUrlTailSlash = stringTest(/\/$/)
export const hasUrlFilename = stringTest(/\.[^/]+$/)
export const removeEmptyUrls = removeEmptiesByProp('Url')
export const removeEmptyIDs = filter(pipe([prop('ID'), isNumberFilled]))
export const removeEmptyFilenames = filter((x) => x.Url && hasUrlFilename(x.Url))
export const removeDuplicatedUrls = removeDuplicatedProp('Url')
export const prependSlash = ifThen(stringTest(/^\//))([identity, sum('/')])
export const popSlash = stringCut(/\/$/)
export const pushSlash = (str) => `${str} /`
export const shiftSlash = stringCut(/^\//)
export const mergeSlashes = stringReplace(/\/\/+/g)('/')
export const urlSplit = stringSplit('/')
export const getTitleFromUrl = pipe([popSlash, urlSplit, arrayLast])
export const urlJoin = join('/')
export const getParentUrl = pipe([popSlash, urlSplit, arrayInit, urlJoin])
export const getFolderFromUrl = ifThen(stringTest(/\./))([getParentUrl, popSlash])
export const getFilenameFromUrl = ifThen(stringTest(/\./))([getTitleFromUrl, NULL])
export const isStrictUrl = (url) => isStringFilled(url) && !hasUrlTailSlash(url)

export const getListRelativeUrl = (webUrl) => (listUrl) => (element = {}) => {
	const { Url, Folder } = element
	if (Folder) {
		const folder = shiftSlash(Folder)
		return Url ? `${folder} /${getTitleFromUrl(Url)}` : folder
	}
	return Url && stringTest(/\//)(Url)
		? Url === '/'
			? '/'
			: shiftSlash(
				arrayLast(
					stringSplit('@list@')(
						stringReplace(listUrl)('@list@')(stringReplace(shiftSlash(webUrl))('@web@')(Url))
					)
				)
			)
		: Url
}

export const getWebRelativeUrl = (webUrl) => (element = {}) => {
	const { Url, Folder } = element
	if (Folder) {
		const folder = shiftSlash(Folder)
		return Url ? `${folder}/${getTitleFromUrl(Url)}` : folder
	}
	return Url && stringTest(/\//)(Url)
		? Url === '/'
			? '/'
			: shiftSlash(arrayLast(stringSplit('@web@')(stringReplace(shiftSlash(webUrl))('@web@')(Url))))
		: Url
}

export class AbstractBox {
	constructor(value, lifter, arrayValidator = identity) {
		this.isArray = isArray(value)
		this.prop = 'Url'
		this.joinProp = 'Url'
		this.value = this.isArray
			? ifThen(isArrayFilled)([
				pipe([
					map(lifter),
					arrayValidator
				]),
				constant([lifter()])
			])(value)
			: lifter(value)
	}

	reduce(f, init) {
		return this.isArray ? reduce(f)(isDefined(init) ? init : [])(this.value) : f(this.value)
	}

	some(f) {
		return this.isArray ? this.value.some(f) : f(this.value)
	}

	async	chain(f) {
		return this.isArray ? Promise.all(map(f)(this.value)) : f(this.value)
	}

	join() {
		return this.isArray
			? join(', ')(map(prop(this.joinProp))(this.value))
			: this.value[this.joinProp]
	}

	getCount() {
		return this.isArray ? this.value.filter(el => el[this.prop]).length : this.value[this.prop] ? 1 : 0
	}

	getHead() {
		return this.isArray ? this.value[0] : this.value
	}

	getHeadPropValue() {
		return this.isArray ? (this.value[0] ? this.value[0][this.prop] : undefined) : this.value[this.prop]
	}

	getIterable() {
		return this.isArray ? this.value : [this.value]
	}
}

//  =============================================================================
//  =====    =        =        =       =====  ====        ===    ===       ======
//  ======  =====  ====  =======  ====  ===    ======  =====  ==  ==  ====  =====
//  ======  =====  ====  =======  ====  ==  ==  =====  ====  ====  =  ====  =====
//  ======  =====  ====  =======  ===   =  ====  ====  ====  ====  =  ===   =====
//  ======  =====  ====      ===      ===  ====  ====  ====  ====  =      =======
//  ======  =====  ====  =======  ====  =        ====  ====  ====  =  ====  =====
//  ======  =====  ====  =======  ====  =  ====  ====  ====  ====  =  ====  =====
//  ======  =====  ====  =======  ====  =  ====  ====  =====  ==  ==  ====  =====
//  =====    ====  ====        =  ====  =  ====  ====  ======    ===  ====  =====
//  =============================================================================

export const deep1Iterator = ({
	contextUrl = '/',
	elementBox,
	bundleSize = REQUEST_BUNDLE_MAX_SIZE
}) => async (f) => {
	let totalElements = 0
	let clientContext = getClientContext(contextUrl)
	const clientContexts = [clientContext]
	const result = await elementBox.chain((element) => {
		totalElements += 1
		if (totalElements >= bundleSize) {
			clientContext = getClientContext(contextUrl)
			clientContexts.push(clientContext)
			totalElements = 0
		}
		return f({ clientContext, element })
	})
	return { clientContexts, result }
}

export const deep1IteratorREST = ({ elementBox }) => (f) => elementBox.chain(async (element) => f({ element }))

//  ===========================================================================
//  =======     ====    ===  =======  =        =        =   ==   =        =====
//  ======  ===  ==  ==  ==   ======  ====  ====  ========  ==  =====  ========
//  =====  =======  ====  =    =====  ====  ====  ========  ==  =====  ========
//  =====  =======  ====  =  ==  ===  ====  ====  =========    ======  ========
//  =====  =======  ====  =  ===  ==  ====  ====      ======  =======  ========
//  =====  =======  ====  =  ====  =  ====  ====  =========    ======  ========
//  =====  =======  ====  =  =====    ====  ====  ========  ==  =====  ========
//  ======  ===  ==  ==  ==  ======   ====  ====  ========  ==  =====  ========
//  =======     ====    ===  =======  ====  ====        =  ====  ====  ========
//  ===========================================================================

const newClientContext = getInstance(SP.ClientContext)

export const getClientContext = pipe([
	pipe([mergeSlashes, popSlash, prependSlash, newClientContext]),
	overstep(method('set_requestTimeout')(REQUEST_TIMEOUT))
])

//  ====================================================================================
//  =====       ==        ==      ==       ====    ===  =======  ==      ==        =====
//  =====  ====  =  =======  ====  =  ====  ==  ==  ==   ======  =  ====  =  ===========
//  =====  ====  =  =======  ====  =  ====  =  ====  =    =====  =  ====  =  ===========
//  =====  ===   =  ========  ======  ====  =  ====  =  ==  ===  ==  ======  ===========
//  =====      ===      ======  ====       ==  ====  =  ===  ==  ====  ====      =======
//  =====  ====  =  ============  ==  =======  ====  =  ====  =  ======  ==  ===========
//  =====  ====  =  =======  ====  =  =======  ====  =  =====    =  ====  =  ===========
//  =====  ====  =  =======  ====  =  ========  ==  ==  ======   =  ====  =  ===========
//  =====  ====  =        ==      ==  =========    ===  =======  ==      ==        =====
//  ====================================================================================

const getSPObjectValues = (asItem) => ifThen(isExists)([
	pipe([
		ifThen(constant(asItem))([methodEmpty('get_listItemAllFields')]),
		switchProp({
			get_fieldValues: methodEmpty('get_fieldValues'),
			get_objectData: pipe([methodEmpty('get_objectData'), methodEmpty('get_properties')]),
			default: tryCatch(pipe([JSON.stringify, sum('Wrong spObject: '), throwError]))(throwError)
		})
	])
])

const getRESTValues = pipe([
	ifThen(hasProp('body'))([pipe([prop('body'), ifThen(isString)([JSON.parse])]), prop('data')]),
	ifThen(hasProp('d'))([pipe([prop('d'), ifThen(hasProp('results'))([prop('results')])])])
])

export const prepareResponseJSOM = (results, opts = {}) => ifThen(isArray)([
	pipe([
		flatten,
		removeUndefineds,
		ifThen(constant(opts.expanded))([
			identity,
			pipe([
				map(getSPObjectValues(opts.asItem)),
				ifThen(constant(opts.groupBy))([
					grouper(opts.groupBy),
					ifThen(constant(opts.mapBy))([mapper(opts.mapBy)])
				])
			])
		])
	]),
	ifThen(constant(opts.expanded))([identity, getSPObjectValues(opts.asItem)])
])(results)

export const prepareResponseREST = (results, opts = {}) => ifThen(isArray)([
	pipe([
		flatten,
		pipe([
			ifThen(constant(opts.expanded))([identity]),
			map(getRESTValues),
			ifThen(constant(opts.groupBy))([
				grouper(opts.groupBy),
				ifThen(constant(opts.mapBy))([mapper(opts.mapBy)])
			])
		])
	]),
	getRESTValues
])(results)

//  =============================================
//  =====  =========    ======  ====       ======
//  =====  ========  ==  ====    ===  ====  =====
//  =====  =======  ====  ==  ==  ==  ====  =====
//  =====  =======  ====  =  ====  =  ====  =====
//  =====  =======  ====  =  ====  =  ====  =====
//  =====  =======  ====  =        =  ====  =====
//  =====  =======  ====  =  ====  =  ====  =====
//  =====  ========  ==  ==  ====  =  ====  =====
//  =====        ===    ===  ====  =       ======
//  =============================================

const getViewOption = ifThen(isObjectFilled)([
	ifThen(isPropFilled('view'))([
		ifThen(isPropFilled('groupBy'))([
			(opts) => concat(getArray(opts.view))(getArray(opts.groupBy)),
			pipe([prop('view'), getArray])
		]),
		EMPTY_ARRAY
	]),
	EMPTY_ARRAY
])

export const load = (clientContext, spObject, opts = {}) => ifThen(hasProp('getEnumerator'))([
	pipe([
		(data) => pipe([
			constant(getViewOption(opts)),
			ifThen(isArrayFilled)([(view) => [data, `Include(${view})`], constant([data])])
		])(data),
		apply('loadQuery')(clientContext)
	]),
	overstep(
		pipe([
			(data) => pipe([
				constant(getViewOption(opts)),
				ifThen(isArrayFilled)([(view) => [data, view], constant([data])])
			])(data),
			apply('load')(clientContext)
		])
	)
])(spObject)

//  =================================================================================
//  =====        =   ==   =        ===     ==  ====  =        ===    ===       ======
//  =====  ========  ==  ==  ========  ===  =  ====  ====  =====  ==  ==  ====  =====
//  =====  ========  ==  ==  =======  =======  ====  ====  ====  ====  =  ====  =====
//  =====  =========    ===  =======  =======  ====  ====  ====  ====  =  ===   =====
//  =====      ======  ====      ===  =======  ====  ====  ====  ====  =      =======
//  =====  =========    ===  =======  =======  ====  ====  ====  ====  =  ====  =====
//  =====  ========  ==  ==  =======  =======  ====  ====  ====  ====  =  ====  =====
//  =====  ========  ==  ==  ========  ===  =   ==   ====  =====  ==  ==  ====  =====
//  =====        =  ====  =        ===     ===      =====  ======    ===  ====  =====
//  =================================================================================

export const executorJSOM = async (clientContext) => new Promise((resolve, reject) => {
	clientContext.executeQueryAsync(
		resolve,
		(sender, args) => {
			const cid = args.get_errorTraceCorrelationId()
			reject(new Error(
				`\nMessage: ${args
					.get_message()
					.replace(
						/\n{1,}/g,
						' '
					)}\nValue: ${args.get_errorValue()}\nType: ${args.get_errorTypeName()}\nId: ${cid}`
			))
		}
	)
})

export const executeJSOM = async (clientContext, spObject, opts) => {
	const spObjects = load(clientContext, spObject, opts)
	await executorJSOM(clientContext)
	return spObjects
}

export const executorREST = async (contextUrl, opts = {}) => pipe([
	mergeSlashes,
	popSlash,
	prependSlash,
	getInstance(SP.RequestExecutor),
	(executor) => new Promise((resolve, reject) => executor.executeAsync({
		...opts,
		method: pipe([
			prop('method'),
			ifThen(stringTest(/post/i))([constant('POST'), constant('GET')])
		])(opts),
		success: resolve,
		error: (res) => {
			const { silent, silentErrors } = opts
			if (!silent && !silentErrors) {
				const { body } = res
				let msg = body
				if (typeOf(res.body) === 'string') {
					try {
						msg = JSON.parse(res.body).error.message.value
					} catch (err) {
						// err
					}
				}
				console.error(`\nMessage: ${res.statusText}\nCode: ${res.statusCode}\nValue: ${msg}`)
			}
			reject(res)
		}
	}))
])(contextUrl)

//  =================================================================
//  =====      ======  =====      ==        ==       =======  =======
//  =====  ===  ====    ===  ====  =  =======  =====  =====   =======
//  =====  ====  ==  ==  ==  ====  =  =======  ===========    =======
//  =====  ===  ==  ====  ==  ======  =======       =====  =  =======
//  =====      ===  ====  ====  ====      ===   ===  ===  ==  =======
//  =====  ===  ==        ======  ==  =======  =====  =  ===  =======
//  =====  ====  =  ====  =  ====  =  =======  =====  =         =====
//  =====  ===  ==  ====  =  ====  =  ========  ===   ======  =======
//  =====      ===  ====  ==      ==        ===     ========  =======
//  =================================================================

export const convertFileContent = switchType({
	arraybuffer: pipe([getInstance(Uint8Array), reduce(functionSum(String.fromCharCode))(''), btoa]),
	default: tryCatch(btoa)(() => identity)
})

export const base64ToBlob = ({ data, type }) => {
	const chunkSize = 512
	const byteCharacters = atob(data)
	const byteArrays = []
	const { length } = byteCharacters
	for (let offset = 0; offset < length; offset += chunkSize) {
		const chunk = byteCharacters.slice(offset, offset + chunkSize)
		const chunkLength = chunk.length
		const byteNumbers = []
		for (let i = 0; i < chunkLength; i += 1) byteNumbers.push(chunk.charCodeAt(i))
		byteArrays.push(new Uint8Array(byteNumbers))
	}
	return new Blob(byteArrays, {
		type: type || 'application/octet-stream'
	})
}

export const blobToDataUrl = (blob) => new Promise(resolve => {
	const reader = new FileReader()
	reader.onloadend = () => resolve(reader.result)
	reader.readAsDataURL(blob)
})

export const blobToArrayBuffer = (blob) => new Promise(resolve => {
	const reader = new FileReader()
	reader.onloadend = () => resolve(reader.result)
	reader.readAsArrayBuffer(blob)
})

export const blobToBase64 = async (blob) => (await blobToDataUrl(blob)).replace(/[^,]+base64,/, '')

export const dataUrlToBinary = (dataUrl) => {
	const BASE64_MARKER = ';base64,'
	const base64Index = dataUrl.indexOf(BASE64_MARKER) + BASE64_MARKER.length
	const raw = window.atob(dataUrl.substring(base64Index))
	const rawLength = raw.length
	const array = new Uint8Array(new ArrayBuffer(rawLength))
	for (let i = 0; i < rawLength; i += 1) {
		array[i] = raw.charCodeAt(i)
	}
	return array
}

//  =================================================================================
//  ======      ==       ====    ===      =======    =        ===     ==        =====
//  =====  ====  =  ====  ==  ==  ==  ===  =======  ==  ========  ===  ====  ========
//  =====  ====  =  ====  =  ====  =  ====  ======  ==  =======  ==========  ========
//  ======  ======  ====  =  ====  =  ===  =======  ==  =======  ==========  ========
//  ========  ====       ==  ====  =      ========  ==      ===  ==========  ========
//  ==========  ==  =======  ====  =  ===  =======  ==  =======  ==========  ========
//  =====  ====  =  =======  ====  =  ====  =  ===  ==  =======  ==========  ========
//  =====  ====  =  ========  ==  ==  ===  ==  ===  ==  ========  ===  ====  ========
//  ======      ==  =========    ===      ====     ===        ===     =====  ========
//  =================================================================================

const setItemSP = (name) => (item) => (value) => item.set_item(name, value)

export const getSPFolderByUrl = (initUrl) => ifThen(constant(initUrl))([
	climb((url) => pipe([methodEmpty('get_folders'), method('getByUrl')(url)]))(
		pipe([stringReplace(/\/\/+/)('/'), stringSplit('/')])(initUrl)
	),
	identity
])

export const setItem = (fieldsInfo) => (fields) => (spObject) => {
	const props = Reflect.ownKeys(fields)
	for (let i = 0; i < props.length; i += 1) {
		const property = props[i]
		const fieldInfo = fieldsInfo[property]
		const fieldValues = fields[property]
		if (fieldInfo) {
			const set = setItemSP(fieldInfo.InternalName)(spObject)
			const setLookupAndUser = (f) => (constructor) => pipe([f(constructor), set])
			switch (fieldInfo.TypeAsString) {
				case 'Lookup':
					setLookupAndUser(setLookup)(SP.FieldLookupValue)(fieldValues)
					break
				case 'LookupMulti':
					setLookupAndUser(setLookupMulti)(SP.FieldLookupValue)(getArray(fieldValues))
					break
				case 'User':
					setLookupAndUser(setLookup)(SP.FieldUserValue)(fieldValues)
					break
				case 'UserMulti':
					setLookupAndUser(setLookupMulti)(SP.FieldUserValue)(getArray(fieldValues))
					break
				case 'TaxonomyFieldTypeMulti':
					switchType({
						object: pipe([
							prop('$2_1'),
							pipe([
								reduce((acc) => (el) => acc.concat(`-1;#${el.get_label()}|${el.get_termGuid()}`))([]),
								join(';#'),
								set
							])
						]),
						array: pipe([join(';'), setItemSP('TaxKeywordTaxHTField')(spObject)]),
						string: setItemSP('TaxKeywordTaxHTField')(spObject)
					})(fieldValues)
					break
				default:
					set(fieldValues)
			}
		} else {
			setItemSP(property)(spObject)(fieldValues)
		}
	}
	spObject.update()
	return spObject
}

export const setLookupMulti = (constructor) => pipe([
	reduce((acc) => ifThen(isExists)([pipe([setLookup(constructor), concat(acc)])]))([]),
	ifThen(isArrayFilled)([identity, NULL])
])

const setLookupValue = (constructor) => (value) => pipe([
	getInstanceEmpty,
	overstep(method('set_lookupId')(value))
])(constructor)

export const setLookup = (constructor) => pipe([
	ifThen(isNull)([
		NULL,
		ifThen(hasProp('get_lookupId'))([
			pipe([methodEmpty('get_lookupId'), setLookupValue(constructor)]),
			pipe([parseInt, ifThen(constant(and(isNumber)(gt(0))))([setLookupValue(constructor), NULL])])
		])
	])
])

export const setFields = (source) => (target) => {
	const props = Reflect.ownKeys(source)
	for (let i = 0; i < props.length; i += 1) {
		const property = props[i]
		const value = source[property]
		if (value !== undefined && target[property]) target[property](value)
	}
	return target
}

export const getContext = methodEmpty('get_context')

export const getWeb = methodEmpty('get_web')

//  =============================================
//  =====        =        ==      ==        =====
//  ========  ====  =======  ====  ====  ========
//  ========  ====  =======  ====  ====  ========
//  ========  ====  ========  =========  ========
//  ========  ====      ======  =======  ========
//  ========  ====  ============  =====  ========
//  ========  ====  =======  ====  ====  ========
//  ========  ====  =======  ====  ====  ========
//  ========  ====        ==      =====  ========
//  =============================================

export const assert = (msg) => (bool) => console.assert(bool === true, msg)

export const assertString = (msg) => (sample) => (str) => assert(`${msg}: ${str}`)(isEqual(sample)(str))

export const assertProp = (property) => (o) => assert(`object has no property "${property}"`)(hasOwnProp(property)(o))

export const assertProps = (props) => (o) => {
	for (let i = 0; i < props.length; i += 1) {
		const property = props[i]
		assertProp(property)(o)
	}
	return o
}

export const assertObject = (props) => (name) => async (promise) => {
	const el = await promise
	assert(`${name} is not an object`)(isObject(el))
	assert(`${name} is empty object`)(isObjectFilled(el))
	assertProps(props)(el)
	return el
}

export const assertCollection = (props) => (name) => async (promise) => {
	const el = await promise
	assert(`${name} collection is not an array`)(isArray(el))
	assert(`${name} collection is an empty array`)(isArrayFilled(el))
	map(pipe([isObjectFilled, assert(`${name} collection element is empty`)]))(el)
	assertProps(props)(el[0])
	return el
}

export const testIsOk = (name) => () => console.log(`${name} is OK`)

//  ============================================================
//  ========    ====        ==  ====  ==        ==       =======
//  =======  ==  ======  =====  ====  ==  ========  ====  ======
//  ======  ====  =====  =====  ====  ==  ========  ====  ======
//  ======  ====  =====  =====  ====  ==  ========  ===   ======
//  ======  ====  =====  =====        ==      ====      ========
//  ======  ====  =====  =====  ====  ==  ========  ====  ======
//  ======  ====  =====  =====  ====  ==  ========  ====  ======
//  =======  ==  ======  =====  ====  ==  ========  ====  ======
//  ========    =======  =====  ====  ==        ==  ====  ======
//  ============================================================

export const getRequestDigest = async (contextUrl) => axios({
	url: `${prependSlash(contextUrl)}/_api/contextinfo`,
	headers: {
		Accept: 'application/json; odata=verbose'
	},
	method: 'POST'
}).then(res => res.data.d.GetContextWebInformation.FormDigestValue)
