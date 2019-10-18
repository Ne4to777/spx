/* eslint no-unused-vars:0 */
/* eslint max-len:0 */

import { getCamlQuery } from '../lib/query-parser'
import { assertString, testWrapper, identity } from '../lib/utility'

const queryAssert = sample => value => assertString('query is')(sample)(getCamlQuery(value))
const emptyAssert = value => queryAssert('')(getCamlQuery(value))

export default {
	run: () => testWrapper('query parse')([
		() => emptyAssert(),
		() => emptyAssert(null),
		() => emptyAssert(true),
		() => emptyAssert(false),
		() => emptyAssert(''),
		() => emptyAssert('a '),
		() => emptyAssert(' a'),
		() => emptyAssert(' a '),
		() => emptyAssert(identity),
		() => emptyAssert([]),
		() => emptyAssert({}),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery(1)),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery('1')),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery('ID eq 1')),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery('ID Eq 1')),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery('Counter ID eq 1')),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery('(ID eq 1)')),
		() => queryAssert('<Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq>')(getCamlQuery(' ( ID eq 1 ) ')),
		() => queryAssert(
			'<Or><Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq><Eq><FieldRef Name="Title"/><Value Type="Text">hi</Value></Eq></Or>'
		)(getCamlQuery('ID eq 1 or Title eq hi')),
		() => queryAssert(
			'<Or><Eq><FieldRef Name="ID"/><Value Type="Text">1</Value></Eq><Or><Eq><FieldRef Name="Title"/><Value Type="Text">hi</Value></Eq><Eq><FieldRef Name="user" LookupId="True"/><Value Type="Lookup">2</Value></Eq></Or></Or>'
		)(getCamlQuery('ID eq 1 or (Title eq hi or lookup user eq 2)'))
	])
}
