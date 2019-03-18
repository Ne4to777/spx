import site from './../modules/site';
import { assertObject, assertCollection, testIsOk } from './../lib/utility';

const PROPS = [
  'AllowCreateDeclarativeWorkflow',
  'AllowDesigner',
  'AllowMasterPageEditing',
  'AllowRevertFromTemplate',
  'AllowSaveDeclarativeWorkflowAsTemplate',
  'AllowSavePublishDeclarativeWorkflow',
  'AllowSelfServiceUpgrade',
  'AllowSelfServiceUpgradeEvaluation',
  'AuditLogTrimmingRetention',
  'CompatibilityLevel',
  'Id',
  'LockIssue',
  'MaxItemsPerThrottledOperation',
  'PrimaryUri',
  'ReadOnly',
  'RequiredDesignerVersion',
  'ServerRelativeUrl',
  'ShareByLinkEnabled',
  'ShowUrlStructure',
  'TrimAuditLog',
  'UIVersionConfigurationEnabled',
  'UpgradeReminderDate',
  'Upgrading',
  'Url'
];

const WEB_TEMPLATE_PROPS = [
  'Description',
  'DisplayCategory',
  'Id',
  'ImageUrl',
  'IsHidden',
  'IsRootWebOnly',
  'IsSubWebOnly',
  'Lcid',
  'Name',
  'Title',
]

const LIST_TEMPLATE_PROPS = [
  'AllowsFolderCreation',
  'BaseType',
  'Description',
  'FeatureId',
  'ImageUrl',
  'InternalName',
  'IsCustomTemplate',
  'ListTemplateTypeKind',
  'Name',
  'OnQuickLaunch',
  'Unique'
]

export default _ => Promise.all([
  assertObject(PROPS)('site')(site.get()),
  assertCollection(WEB_TEMPLATE_PROPS)('web template')(site.getWebTemplates()),
  assertCollection(LIST_TEMPLATE_PROPS)('custom lists template')(site.getCustomListTemplates()),
]).then(testIsOk('site'))
