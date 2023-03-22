type RelationshipType = "OneToManyRelationship" | "ManyToManyRelationship";

export interface IRelationshipDefinition {
  SchemaName?: string;
  RelationshipType?: RelationshipType;
}

export interface IManyToManyRelationship {
  SchemaName: string;
  Entity1LogicalName: string;
  Entity2LogicalName: string;
  Entity1IntersectAttribute: string;
  Entity2IntersectAttribute: string;
  RelationshipType: string;
  IntersectEntityName: string;
}

export interface IOneToManyRelationship extends IRelationshipDefinition {
  ReferencedAttribute?: string;
  ReferencedAttributeName?: string;
  ReferencedEntity?: string;
  ReferencedEntityName?: string;
  ReferencingAttribute?: string;
  ReferencingAttributeName?: string;
  ReferencingEntity?: string;
  ReferencingEntityName?: string;
  ReferencedEntityNavigationPropertyName?: string;
  ReferencingEntityNavigationPropertyName?: string;
}

export interface IEntityDefinition {
  LogicalName: string;
  DisplayName: {
    UserLocalizedLabel: {
      Label: string;
    };
  };
  PrimaryIdAttribute: string;
  PrimaryNameAttribute: string;
  EntitySetName: string;
  DisplayCollectionName: {
    UserLocalizedLabel: {
      Label: string;
    };
  };
  IsQuickCreateEnabled: boolean;
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
  nnRelationship: IManyToManyRelationship;
  currentEntity: IEntityDefinition;
  intersectEntity: IEntityDefinition;
  associatedEntity: IEntityDefinition;
  associatedView: IViewDefinition;
}
