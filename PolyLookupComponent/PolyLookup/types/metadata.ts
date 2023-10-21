type RelationshipType = "OneToManyRelationship" | "ManyToManyRelationship";

export interface IRelationshipDefinition {
  SchemaName: string;
  RelationshipType: RelationshipType;
}

export interface IManyToManyRelationship extends IRelationshipDefinition {
  Entity1LogicalName: string;
  Entity2LogicalName: string;
  Entity1IntersectAttribute: string;
  Entity2IntersectAttribute: string;
  IntersectEntityName: string;
}

export interface IOneToManyRelationship extends IRelationshipDefinition {
  ReferencedAttribute: string;
  ReferencedEntity: string;
  ReferencingAttribute: string;
  ReferencingEntity: string;
  ReferencedEntityNavigationPropertyName: string;
  ReferencingEntityNavigationPropertyName: string;
}

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
  relationship1: IManyToManyRelationship | IOneToManyRelationship;
  relationship2: IOneToManyRelationship | undefined;
  currentEntity: IEntityDefinition;
  intersectEntity: IEntityDefinition;
  associatedEntity: IEntityDefinition;
  associatedView: IViewDefinition;
  currentIntesectAttribute: string;
  associatedIntesectAttribute: string;
  currentEntityNavigationPropertyName?: string;
  associatedEntityNavigationPropertyName?: string;
}

export const isOneToMany = (r: IRelationshipDefinition | undefined): r is IOneToManyRelationship => {
  return r?.RelationshipType === "OneToManyRelationship";
};

export const isManyToMany = (r: IRelationshipDefinition | undefined): r is IManyToManyRelationship => {
  return r?.RelationshipType === "ManyToManyRelationship";
};
