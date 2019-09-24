import {
	getClientContext,
	getInstance,
	load,
	executorJSOM,
	setFields,
	switchCase,
	typeOf
} from '../lib/utility'

const QUERY_TEMPLATES = [
	'{searchboxquery}',
	'-IsDocument:true',
	'-Site:http://mysites.aura.dme.aero.corp',
	'-Site:http://wiki.aura.dme.aero.corp',
	'-contentclass:STS_Site',
	'-contentclass:STS_Web',
	'-contentclass:STS_Document',
	'-contentclass:STS_ListItem_DocumentLibrary',
	'-contentclass:STS_ListItem_PublishingPages',
	'-contentclass:STS_ListItem_DiscussionBoard',
	'-contentclass:STS_ListItem_PictureLibrary',
	'-contentclass:STS_ListItem_Events',
	// '-contentclass:STS_ListItem_GenericList',
	'-contentclass:STS_List_*',
	// '-contentclass:STS_ListItem_*',
	'-contentclass:urn:content-class:SPSPeople*',
	'-Site:http://aura.dme.aero.corp/wikilibrary',
	'-Site:http://aura.dme.aero.corp/app',
	'-Site:http://aura.dme.aero.corp/Tikhonchuk*',
	'-Site:http://aura.dme.aero.corp/SitePages*',
	'-Site:http://aura.dme.aero.corp/Pages',
	'-Site:http://aura.dme.aero.corp/PublishingImages*',
	'-Site:http://aura.dme.aero.corp/Survey',
	'-Site:http://aura.dme.aero.corp/News',
	'-Site:http://aura.dme.aero.corp/social',
	'-Site:http://aura.dme.aero.corp/Intellect',
	'-Site:http://aura.dme.aero.corp/board',
	'-Site:http://aura.dme.aero.corp/buro',
	'-Site:http://aura.dme.aero.corp/marketingpresentations',
	'-Site:http://aura.dme.aero.corp/crowd',
	'-Site:http://aura.dme.aero.corp/System',
	'-Site:http://aura.dme.aero.corp/Forum',
	'-Site:http://aura.dme.aero.corp/Lenta',
	'-Site: http://aura.dme.aero.corp/lib',
]


const lifter = switchCase(typeOf)({
	object: query => Object.assign({}, query),
	string: (query = '') => ({
		Query: query,
	}),
	default: () => ({
		Query: ''
	})
})

class Search {
	constructor(parent, query) {
		this.name = 'search'
		this.parent = parent
		this.element = lifter(query)
		this.contextUrl = parent.box.getHeadPropValue()
	}

	async get(opts) {
		const clientContext = getClientContext(this.contextUrl || '/')
		const { element } = this
		const searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext)
		const keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext)

		setFields({
			set_queryText: element.Query.toLowerCase(),
			set_clientType: element.ClientType || 'AllResultsQuery',
			// set_queryTemplate: QUERY_TEMPLATES.join(' '),
			set_queryTemplate: element.QueryTemplate,
			set_refiners: element.Refiners,
			set_rowsPerPage: element.RowsPerPage || 10,
			set_totalRowsExactMinimum: element.TotalRowsExactMinimum || 11,
			set_blockDedupeMode: element.BlockDedupeMode,
			set_bypassResultTypes: element.BypassResultTypes,
			set_collapseSpecification: element.CollapseSpecification,
			set_culture: element.Culture || 1033,
			set_desiredSnippetLength: element.DesiredSnippetLength,
			set_enableInterleaving: element.EnableInterleaving,
			set_enableNicknames: element.EnableNicknames,
			set_enableOrderHitHighlightedProperty: element.EnableOrderHitHighlightedProperty,
			set_enablePhonetic: element.EnablePhonetic,
			set_enableQueryRules: element.EnableQueryRules,
			set_enableSorting: element.EnableSorting,
			set_enableStemming: element.EnableStemming,
			set_generateBlockRankLog: element.GenerateBlockRankLog,
			set_hiddenConstrains: element.HiddenConstrains,
			set_hitHighlightedMultivaluePropertyLimit: element.HitHighlightedMultivaluePropertyLimit,
			set_impressionID: element.ImpressionID,
			set_maxSnippetLength: element.MaxSnippetLength,
			set_personalizationData: element.PersonalizationData,
			set_processBestBets: element.ProcessBestBets,
			set_processPersonalFavorites: element.ProcessPersonalFavorites,
			set_queryTag: element.QueryTag,
			set_rankingModelId: element.RankingModelId,
			set_refinementFilters: element.RefinementFilters,
			set_reorderingRules: element.ReorderingRules,
			set_resultsUrl: element.ResultsUrl,
			set_rowLimit: element.RowLimit || 10,
			set_startRow: element.StartRow,
			set_showPeopleNameSuggestions: element.ShowPeopleNameSuggestions,
			set_summaryLength: element.SummaryLength,
			set_timeZoneId: element.TimeZoneId || 51,
			set_timeout: element.Timeout,
			set_trimDuplicates: element.TrimDuplicates,
			set_trimDuplicatesIncludeId: element.TrimDuplicatesIncludeId,
			set_uiLanguage: element.UiLanguage || 1033
		})(keywordQuery)

		load(clientContext, keywordQuery, opts)
		const result = searchExecutor.executeQuery(keywordQuery)
		await executorJSOM(clientContext)

		return result.get_value().ResultTables[0].ResultRows
	}
}

export default getInstance(Search)
