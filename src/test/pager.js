/* eslint no-unused-vars:0 */
import web from '../modules/web'
import {
	assertCollection,
	testWrapper,
} from '../lib/utility'

const PROPS = [
	'AppAuthor',
	'AppEditor',
	'Attachments',
	'Author',
	'ContentTypeId',
	'Created',
	'Created_x0020_Date',
	'Editor',
	'FSObjType',
	'FileDirRef',
	'FileLeafRef',
	'FileRef',
	'File_x0020_Type',
	'FolderChildCount',
	'GUID',
	'ID',
	'InstanceID',
	'ItemChildCount',
	'Last_x0020_Modified',
	'MetaInfo',
	'Modified',
	'Order',
	'ProgId',
	'ScopeId',
	'SortBehavior',
	'SyncClientId',
	'Title',
	'UniqueId',
	'WorkflowInstanceID',
	'WorkflowVersion',
	'owshiddenversion',
	'_CopySource',
	'_HasCopyDestinations',
	'_IsCurrentVersion',
	'_Level',
	'_ModerationComments',
	'_ModerationStatus',
	'_UIVersion',
	'_UIVersionString'
]

const assertCollectionProps = assertCollection(PROPS)

const pager = web('test/spx').list('Pager').pager()

export default {
	get: () => testWrapper('pager GET')([
		() => assertCollectionProps('pager next')(pager.next()),
		() => assertCollectionProps('pager next')(pager.next()),
		() => assertCollectionProps('pager previous')(pager.previous()),
	]),
	all() {
		testWrapper('pager ALL')([
			this.get,
		])
	}
}
