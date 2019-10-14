import {
	getClientContext,
	getInstance,
	load,
	executorJSOM,
	setFields,
	switchType
} from '../lib/utility'

const lifter = switchType({
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
		this.rowsPerPage = this.element.RowsPerPage || 10
		if (this.element.StartRow === undefined) this.element.StartRow = 0
	}

	async get(opts) {
		const clientContext = getClientContext(this.contextUrl || '/')
		const { element } = this
		const searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext)
		const keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext)

		setFields({
			set_queryText: element.Query,
			set_clientType: element.ClientType || 'AllResultsQuery',
			set_queryTemplate: element.QueryTemplate ? element.QueryTemplate.join(' ') : undefined,
			set_refiners: element.Refiners,
			set_rowsPerPage: element.RowsPerPage || this.rowsPerPage,
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
			set_showPeopleNameSuggestions: element.ShowPeopleNameSuggestions || false,
			set_summaryLength: element.SummaryLength,
			set_timeZoneId: element.TimeZoneId,
			set_timeout: element.Timeout,
			set_trimDuplicates: element.TrimDuplicates,
			set_trimDuplicatesIncludeId: element.TrimDuplicatesIncludeId,
			set_uiLanguage: element.UiLanguage || 1033
		})(keywordQuery)

		load(clientContext, keywordQuery, opts)
		const result = searchExecutor.executeQuery(keywordQuery)
		await executorJSOM(clientContext)
		return result.get_value()
	}

	async next(opts) {
		const res = await this.get(opts)
		this.element.StartRow += this.rowsPerPage
		return res
	}

	async previous(opts) {
		const res = await this.get(opts)
		this.element.StartRow -= this.rowsPerPage
		return res
	}

	async move(n = 1, opts) {
		this.element.StartRow = n * this.rowsPerPage
		return this.get(opts)
	}
}

export default getInstance(Search)
