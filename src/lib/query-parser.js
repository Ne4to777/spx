import {
	chunkArray,
	stringMatch,
	arrayTail,
	arrayHead,
	isObject,
	getArray,
	reduce,
	map,
	pipe,
	stringReplace,
	stringTrim,
	stringSplit,
	typeOf,
	stringTest
} from './utility'

const IN_CHUNK_SIZE = 500
const IN_CUSTOM_DELIMETER = '_DELIMITER_'
const GROUP_REGEXP_STR = '\\s(&&||||and|or)\\s'
const GROUP_REGEXP = new RegExp(GROUP_REGEXP_STR, 'i')
const GROUP_BRACES_REGEXP = new RegExp(`\\((.*${GROUP_REGEXP_STR}.*)\\)`, 'i')
const TIME_STAMP_ISO_REGEXP = /^\d\d\d\d-\d\d-\d\dT\d\d:\d\d:\d\d(\.\d\d\d)?Z$/
const COLUMN_TYPES = {
	AppAuthor: 'lookup',
	AppEditor: 'lookup',
	Attachments: 'attachments',
	Author: 'lookup',
	BaseName: 'computed',
	ContentType: 'computed',
	ContentTypeId: 'contentTypeId',
	Created: 'dateTime',
	Created_x0020_Date: 'lookup',
	DocIcon: 'computed',
	Edit: 'computed',
	Editor: 'lookup',
	EncodedAbsUrl: 'computed',
	FSObjType: 'number',
	FileDirRef: 'lookup',
	FileLeafRef: 'file',
	FileRef: 'file',
	File_x0020_Type: 'text',
	FolderChildCount: 'number',
	GUID: 'guid',
	HTML_x0020_Type: 'computed',
	ID: 'counter',
	InstanceID: 'integer',
	ItemChildCount: 'number',
	Last_x0020_Modified: 'lookup',
	LinkFilename: 'computed',
	LinkFilename2: 'computed',
	LinkFilenameNoMenu: 'computed',
	LinkTitle: 'computed',
	LinkTitle2: 'computed',
	LinkTitleNoMenu: 'computed',
	MetaInfo: 'lookup',
	Modified: 'dateTime',
	Order: 'number',
	PermMask: 'computed',
	ProgId: 'lookup',
	ScopeId: 'lookup',
	SelectTitle: 'computed',
	ServerUrl: 'computed',
	SortBehavior: 'lookup',
	SyncClientId: 'lookup',
	Title: 'text',
	UniqueId: 'lookup',
	WorkflowInstanceID: 'guid',
	WorkflowVersion: 'integer',
	_CopySource: 'text',
	_EditMenuTableEnd: 'computed',
	_EditMenuTableStart: 'computed',
	_EditMenuTableStart2: 'computed',
	_HasCopyDestinations: 'Boolean',
	_IsCurrentVersion: 'Boolean',
	_Level: 'integer',
	_ModerationComments: 'note',
	_ModerationStatus: 'modStat',
	_UIVersion: 'integer',
	_UIVersionString: 'text'
}

const COLUMN_TYPES_REGEXP = /(Text|Note|Number|Integer|Counter|Boolean|Lookup|User|DateTime|Date|Time|Computed|Currency|ModStat|Guid|File|Attachments|LookupMulti|UserMulti)\s/i
const OPERATORS_REGEXP = /\s(eq|neq|geq|leq|gt|lt|isnull|isnotnull|contains|beginsWith|includes|notincludes|search|in|membership)(\s|$)/i

const COLUMN_TYPES_MAPPED = {
	text: 'Text',
	note: 'Text',
	number: 'Number',
	integer: 'Text',
	counter: 'Text',
	boolean: 'Boolean',
	lookup: 'Lookup',
	user: 'Text',
	datetime: 'Text',
	date: 'Text',
	time: 'Text',
	currency: 'Text',
	computed: 'Text',
	lookupmulti: 'Lookup',
	usermulti: 'Text',
	modstat: 'ModStat',
	guid: 'Guid',
	file: 'File',
	attachments: 'Attachments'
}
const OPERATORS_MAPPED = {
	eq: 'Eq',
	neq: 'Neq',
	geq: 'Geq',
	leq: 'Leq',
	gt: 'Gt',
	lt: 'Lt',
	isnull: 'IsNull',
	isnotnull: 'IsNotNull',
	beginswith: 'BeginsWith',
	contains: 'Contains',
	includes: 'Includes',
	notincludes: 'NotIncludes',
	in: 'In',
	membership: 'Membership',
	search: 'Search'
}

const MEMBERSHIP_VALUES = {
	user: 'SPWeb.Users', // field with users includes current user
	users: 'SPWeb.AllUsers', // field with any user(s)
	group: 'CurrentUserGroups', // field with groups includes current user
	groups: 'SPWeb.Groups' // field with any group(s)
	// SPGroup - field with groups includes current group ???
}

const GROUP_OPERATORS_MAPPED = {
	'&&': 'And',
	'||': 'Or'
}
const normalizeOperator = operator => GROUP_OPERATORS_MAPPED[operator] || (/^and$/i.test(operator) ? 'And' : 'Or')

const getStringFromOuterBraces = str => {
	const founds = str.match(
		new RegExp(`^[^\\(]*\\((.*${GROUP_REGEXP.test(str) ? `${GROUP_REGEXP_STR}.*` : ''})\\)[^\\)]*$`, 'i')
	)
	return founds && founds[1] ? founds[1] : str
}

const trimBraces = (str = '') => {
	if (str[0] === '(' && str.substr(-1) === ')') {
		const openBracesCount = stringMatch(/\(/g)(str).length
		if (openBracesCount !== stringMatch(/\)/g)(str).length) throw new Error('query has wrong braces')
		return openBracesCount >= stringMatch(new RegExp(GROUP_REGEXP_STR, 'ig'))(str).length
			? getStringFromOuterBraces(str)
			: str
	}
	return str
}

const convertExpression = str => {
	let operator
	let operatorMatch
	let fieldOption = ''
	let type
	let valueOpts = ''
	const operatorMatches = str.match(OPERATORS_REGEXP)
	if (operatorMatches) {
		[operatorMatch] = operatorMatches
		operator = operatorMatches[1].toLowerCase()
	} else {
		throw new SyntaxError(`Wrong operator in "${str}"`)
	}
	const operatorNorm = OPERATORS_MAPPED[operator]
	const strSplits = str.split(new RegExp(operatorMatch, 'i'))
	let name = strSplits[0].replace(/\s\s+/g, ' ').trim()
	if (COLUMN_TYPES_REGEXP.test(name)) {
		const columnSplits = name.split(' ')
		type = columnSplits.shift().toLowerCase()
		name = columnSplits.join(' ')
	} else {
		type = COLUMN_TYPES[name] || 'text'
	}

	let value = strSplits[1]

	switch (type) {
		case 'datetime':
		case 'time':
		case 'date':
			valueOpts = 'StorageTZ="True"'
			if (type !== 'date') valueOpts += ' IncludeTimeValue="True"'
			value = new Date(value).toISOString()
			break
		case 'text':
			if (TIME_STAMP_ISO_REGEXP.test(value)) valueOpts = 'StorageTZ="True"'
			break
		case 'boolean':
			value = /^(true|yes|1)$/.test(value) ? '1' : '0'
			break
		case 'lookup':
		case 'lookupmulti':
			fieldOption = ' LookupId="True"'
			break
		default: {
			// default
		}
	}
	const typeNorm = COLUMN_TYPES_MAPPED[type]
	if (type !== 'text') value = value.trim()

	switch (operator) {
		case 'in': {
			const valueStrings = []
			const valueChunks = []
			value = value.split(value.indexOf(IN_CUSTOM_DELIMETER) >= 0 ? IN_CUSTOM_DELIMETER : ',')
			for (let i = 0; i < value.length; i += 1) {
				const valueItem = value[i]
				valueStrings.push(`<Value Type="${typeNorm}"${valueOpts}>${valueItem}</Value>`)
			}
			if (value.length > IN_CHUNK_SIZE) {
				const chunks = chunkArray(IN_CHUNK_SIZE)(valueStrings)
				for (let i = chunks.length - 1; i >= 0; i -= 1) {
					valueChunks.push(
						`<In><FieldRef Name="${name}"${fieldOption}/><Values>${chunks[i].join('')}</Values></In>`
					)
				}
				let itemsStr = '<IsNull><FieldRef Name="ID"/></IsNull>'
				for (let i = 0; i < valueChunks.length; i += 1) {
					const valueChunk = valueChunks[i]
					itemsStr = `<Or>${valueChunk}${itemsStr}</Or>`
				}
				return itemsStr
			}
			return `<In><FieldRef Name="${name}"${fieldOption}/><Values>${valueStrings.join('')}</Values></In>`
		}
		case 'search':
			return pipe([
				stringReplace(/\s+/g)(' '),
				stringTrim,
				stringSplit(' '),
				map(el => `${type} ${name} contains ${el}`),
				concatQueries(),
				convertGroupR
			])(value)
		case 'membership': {
			const valueType = MEMBERSHIP_VALUES[value.toLowerCase()]
			return `<${operatorNorm} Type="${valueType}"><FieldRef Name="${name}"/></${operatorNorm}>`
		}
		default: {
			// default
		}
	}
	const valueOrNull = /null/.test(operator) ? '' : `<Value Type="${typeNorm}"${valueOpts}>${value}</Value>`
	return `<${operatorNorm}><FieldRef Name="${name}"${fieldOption}/>${valueOrNull}</${operatorNorm}>`
}

const convertGroupR = str => {
	let firstExpr; let secondExpr; let
		operator
	let chars = ''
	let isBracesBefore = false
	let isBracesAfter = false
	let operatorIndex = 0
	if (GROUP_BRACES_REGEXP.test(str)) {
		let groupIsOpen = 0
		for (let i = 0; i < str.length; i += 1) {
			const char = str[i]
			if (char === '(') {
				if (!operatorIndex) isBracesBefore = true
				groupIsOpen += 1
			} else if (char === ')') {
				if (operatorIndex) isBracesAfter = true
				groupIsOpen -= 1
			} else if (!groupIsOpen) {
				chars = chars.concat(char)
				if (!operatorIndex) operatorIndex = i
			}
		}
	}
	if (chars && isBracesBefore && isBracesAfter) {
		operator = chars.trim()
		firstExpr = trimBraces(str.slice(0, operatorIndex))
		secondExpr = trimBraces(str.slice(operatorIndex + chars.length))
	} else {
		const matches = trimBraces(str).match(GROUP_BRACES_REGEXP)
		if (matches) {
			const dummyStr = '_dummy_'
			const dummyRE = new RegExp(dummyStr)
			const groupStr = matches[1]
			const dummySplits = str.replace(groupStr, dummyStr).split(GROUP_REGEXP)
			const dummySplit0 = dummySplits[0]
			if (dummySplits.length === 3) {
				[, operator] = dummySplits
				const dummySplit2 = dummySplits[2]
				if (dummyRE.test(dummySplit0)) {
					firstExpr = getStringFromOuterBraces(dummySplit0.replace(dummyStr, groupStr))
					secondExpr = dummySplit2
				} else if (dummyRE.test(dummySplit2)) {
					firstExpr = dummySplit0
					secondExpr = getStringFromOuterBraces(dummySplit2.replace(dummyStr, groupStr))
				}
			} else {
				return convertExpression(dummySplit0)
			}
		} else {
			const splits = str.split(GROUP_REGEXP);
			[firstExpr] = splits
			if (splits.length === 3) {
				[, , secondExpr] = splits;
				[, operator] = splits
			} else {
				return convertExpression(firstExpr)
			}
		}
	}
	const operatorNorm = normalizeOperator(operator)
	return `<${operatorNorm}>${convertGroupR(firstExpr)}${convertGroupR(secondExpr)}</${operatorNorm}>`
}

export const getCamlQuery = str => {
	switch (typeOf(str)) {
		case 'number':
			return getCamlQuery(`ID eq ${str}`)
		case 'object':
			return getCamlView(str)
		case 'string':
			if (str) {
				if (str === (+str).toString()) return getCamlQuery(`ID eq ${str}`)
				if (stringTest(/\s/)(str) && stringSplit(/\s/)(str).length > 1) {
					const sanitaizedStr = trimBraces(str.replace(/\s{2,}/g, ' ').replace(/^\s|\s$/g, ''))
					if (stringSplit(/\s/)(sanitaizedStr).length > 1) {
						if (/^<|>$/.test(sanitaizedStr)) return str
						return GROUP_REGEXP.test(sanitaizedStr)
							? convertGroupR(sanitaizedStr)
							: convertExpression(sanitaizedStr)
					}
				}
			}
			return ''
		default:
			return ''
	}
}

export const getCamlView = (str = {}) => {
	let orderBys
	let orderByStr = ''
	let scopeStr = ''
	const {
		OrderBy, Scope, Limit
	} = str
	let {
		Query = ''
	} = str
	if (!isObject(str)) Query = str
	if (OrderBy) {
		orderBys = getArray(OrderBy)
		for (let i = 0; i < orderBys.length; i += 1) {
			const item = orderBys[i]
			orderByStr += `<FieldRef Name="${item.split('>')[0]}"${/>/.test(item) ? ' Ascending="FALSE"' : ''}/>`
		}
		orderByStr = `<OrderBy>${orderByStr}</OrderBy>`
	}
	if (Scope) {
		scopeStr = /allItems/i.test(Scope)
			? ' Scope="Recursive"'
			: /^items$/i.test(Scope)
				? ' Scope="FilesOnly"'
				: /^all$/i.test(Scope)
					? ' Scope="RecursiveAll"'
					: ''
	}
	const whereStr = Query ? `<Where>${getCamlQuery(Query)}</Where>` : ''
	const queryStr = whereStr || orderByStr ? `<Query>${whereStr}${orderByStr}</Query>` : ''
	const limitStr = `<RowLimit>${Limit || 160000}</RowLimit>`
	return `<View${scopeStr}>${queryStr}${limitStr}</View>`
}

export const craftQuery = ({
	joiner = '||', operator = values === undefined ? 'isnotnull' : 'eq', columns, values
}) => {
	const normalizedJoiner = normalizeOperator(joiner)
	const columnsArray = getArray(columns)
	const valuesArray = getArray(values)
	let query = normalizedJoiner === 'And' ? `ID ${OPERATORS_MAPPED.isnotnull}` : `ID ${OPERATORS_MAPPED.isnull}`
	if (valuesArray.length) {
		for (let i = columnsArray.length - 1; i >= 0; i -= 1) {
			for (let j = valuesArray.length - 1; j >= 0; j -= 1) {
				query = `(${columnsArray[i]} ${`${operator} ${valuesArray[j]}`} ${normalizedJoiner} ${query})`
			}
		}
	} else {
		for (let i = columnsArray.length - 1; i >= 0; i -= 1) {
			query = `(${columnsArray[i]} ${operator} ${normalizedJoiner} ${query})`
		}
	}
	return trimBraces(query)
}

const concat2Queries = (joiner = '||') => query1 => query2 => {
	const normalizedJoiner = normalizeOperator(joiner)
	const dummy = normalizedJoiner === 'And' ? `ID ${OPERATORS_MAPPED.isnotnull}` : `ID ${OPERATORS_MAPPED.isnull}`
	if (query1) {
		const trimmed1 = trimBraces(query1)
		if (query2) {
			const trimmed2 = trimBraces(query2)
			return GROUP_REGEXP.test(trimmed1)
				? GROUP_REGEXP.test(trimmed2)
					? `(${trimmed1}) ${normalizedJoiner} (${trimmed2})`
					: `(${trimmed1}) ${normalizedJoiner} ${trimmed2}`
				: GROUP_REGEXP.test(trimmed2)
					? `${trimmed1} ${normalizedJoiner} (${trimmed2})`
					: `${trimmed1} ${normalizedJoiner} ${trimmed2}`
		}
		return GROUP_REGEXP.test(trimmed1)
			? `(${trimmed1}) ${normalizedJoiner} ${dummy}`
			: `${trimmed1} ${normalizedJoiner} ${dummy}`
	}
	if (query2) {
		const trimmed2 = trimBraces(query2)
		return GROUP_REGEXP.test(trimmed2)
			? `${dummy} ${normalizedJoiner} (${trimmed2})`
			: `${dummy} ${normalizedJoiner} ${trimmed2}`
	}
	return `(${dummy}) ${normalizedJoiner} ${dummy}`
}

export const concatQueries = (joiner = '||') => (queries = []) => reduce(
	concat2Queries(joiner)
)(arrayHead(queries))(arrayTail(queries))

export const camlLog = str => console.log(str && DOMParser ? new DOMParser().parseFromString(str, 'text/xml') : str)
