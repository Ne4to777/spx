import { getClientContext, getInstance, Box, load, executorJSOM, setFields } from './../utility';

// const QUERY_TEMPLATES = [
//   '{searchboxquery}',
//   '-IsDocument:true',
//   '-Site:http://mysites.aura.dme.aero.corp',
//   '-Site:http://wiki.aura.dme.aero.corp',
//   '-contentclass:STS_Site',
//   '-contentclass:STS_Web',
//   '-contentclass:STS_Document',
//   '-contentclass:STS_ListItem_DocumentLibrary',
//   '-contentclass:STS_ListItem_PublishingPages',
//   '-contentclass:STS_ListItem_DiscussionBoard',
//   '-contentclass:STS_ListItem_PictureLibrary',
//   '-contentclass:STS_ListItem_Events',
//   // '-contentclass:STS_ListItem_GenericList',
//   '-contentclass:STS_List_*',
//   // '-contentclass:STS_ListItem_*',
//   '-contentclass:urn:content-class:SPSPeople*',
//   '-Site:http://aura.dme.aero.corp/wikilibrary',
//   '-Site:http://aura.dme.aero.corp/app',
//   '-Site:http://aura.dme.aero.corp/Tikhonchuk*',
//   '-Site:http://aura.dme.aero.corp/SitePages*',
//   '-Site:http://aura.dme.aero.corp/Pages',
//   '-Site:http://aura.dme.aero.corp/PublishingImages*',
//   '-Site:http://aura.dme.aero.corp/Survey',
//   '-Site:http://aura.dme.aero.corp/News',
//   '-Site:http://aura.dme.aero.corp/social',
//   '-Site:http://aura.dme.aero.corp/Intellect',
//   '-Site:http://aura.dme.aero.corp/board',
//   '-Site:http://aura.dme.aero.corp/buro',
//   '-Site:http://aura.dme.aero.corp/marketingpresentations',
//   '-Site:http://aura.dme.aero.corp/crowd',
//   '-Site:http://aura.dme.aero.corp/System',
//   '-Site:http://aura.dme.aero.corp/Forum',
//   '-Site:http://aura.dme.aero.corp/Lenta',
//   '-Site:http://aura.dme.aero.corp/lib',
// ];
export default (parent, elements) => (opts = {}) => {
  const box = getInstance(Box)(elements);
  return parent.box.chainAsync(async context => {
    const clientContext = getClientContext(context.Url);
    const keywordQuery = new Microsoft.SharePoint.Client.Search.Query.KeywordQuery(clientContext);
    const searchExecutor = new Microsoft.SharePoint.Client.Search.Query.SearchExecutor(clientContext);
    const result = box.chain(element => {
      const sortList = keywordQuery.get_sortList();
      sortList.add('[formula:rank]', 1);
      setFields({
        set_queryText: element.Url.toLowerCase(),
        set_clientType: element.ClientType,
        set_queryTemplate: QUERY_TEMPLATES.join(' '),
        set_refiners: element.Refiners,
        // set_rowsPerPage: element.RowsPerPage || 10,
        // set_totalRowsExactMinimum: element.TotalRowsExactMinimum || 11,
        set_blockDedupeMode: element.BlockDedupeMode,
        set_bypassResultTypes: element.BypassResultTypes,
        set_collapseSpecification: element.CollapseSpecification,
        set_culture: element.Culture,
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
        set_rowLimit: element.RowLimit,
        set_startRow: element.StartRow,
        set_showPeopleNameSuggestions: element.ShowPeopleNameSuggestions,
        set_summaryLength: element.SummaryLength,
        set_timeZoneId: element.TimeZoneId,
        set_timeout: element.Timeout,
        set_trimDuplicates: element.TrimDuplicates,
        set_trimDuplicatesIncludeId: element.TrimDuplicatesIncludeId,
        set_uiLanguage: element.UiLanguage
      })(keywordQuery)
      load(clientContext)(keywordQuery)(opts);
      return searchExecutor.executeQuery(keywordQuery);
    })
    await executorJSOM(clientContext)(opts);
    console.log(result);
    console.log(keywordQuery.get_objectData().get_properties());
    const resultTables = result.get_value().ResultTables;
    console.log(resultTables);
    if (resultTables.length) {
      return resultTables[0].ResultRows
    }
  })
}