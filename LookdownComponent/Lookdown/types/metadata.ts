export interface IEntityDefinition {
  LogicalName: string;
  DisplayName: {
    UserLocalizedLabel: {
      Label: string;
    } | null;
  };
  DisplayNameLocalized: string;
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string;
  EntitySetName: string;
  DisplayCollectionName: {
    UserLocalizedLabel: {
      Label: string;
    } | null;
  };
  DisplayCollectionNameLocalized: string;
  IsQuickCreateEnabled?: boolean;
  IconVectorName?: string;
  IconSmallName?: string;
  IconMediumName?: string;
  PrimaryImageAttribute?: string;
  RecordImageUrlTemplate?: string;
}

interface IViewRow {
  Name: string;
  Id: string;
  Cells: IViewCell[];
  MultiObjectIdField: string;
  LayoutStyle: string;
}

interface IViewCell {
  Name: string;
  Width: number;
  RelatedEntityName: string;
  DisableMetaDataBinding: boolean;
  LabelId: string;
  IsHidden: boolean;
  DisableSorting: boolean;
  AddedBy: string;
  Desc: string;
  CellType: string;
  ImageProviderWebresource: string;
  ImageProviderFunctionName: string;
}

export interface IViewLayout {
  Name: string;
  Object: number;
  Rows: IViewRow[];
  CustomControlDescriptions: [];
  Jump: string;
  Select: true;
  Icon: true;
  Preview: true;
  IconRenderer: "";
}

export interface IViewDefinition {
  savedqueryid: string;
  name: string;
  fetchxml: string;
  layoutjson: IViewLayout;
  querytype: number;
}

export interface IMetadata {
  lookupEntity: IEntityDefinition;
  lookupView: IViewDefinition;
}
