import { LanguagePack } from "./languagePack";

export interface PolyLookupProps {
  currentTable: string;
  currentRecordId: string;
  relationshipName: string;
  relationship2Name?: string;
  relationshipType: RelationshipTypeEnum;
  clientUrl: string;
  lookupView?: string;
  itemLimit?: number;
  pageSize?: number;
  disabled?: boolean;
  formType?: XrmEnum.FormType;
  outputSelectedItems?: boolean;
  showIcon?: ShowIconOptions;
  tagAction?: TagAction;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  fluentDesign?: ComponentFramework.FluentDesignState;
  onChange?: (selectedItems: EntityReference[] | undefined) => void;
  onQuickCreate?: (
    entityName: string | undefined,
    primaryAttribute: string | undefined,
    value: string | undefined,
    useQuickCreateForm: boolean | undefined
  ) => Promise<string | undefined>;
}

export enum RelationshipTypeEnum {
  ManyToMany,
  Custom,
  Connection,
}

export enum TagAction {
  None = 0,
  OpenInline = 1,
  OpenDialog = 2,
  OpenInlineIntersect = 3,
  OpenDialogIntersect = 4,
}

export interface EntityOption {
  id: string; // if relationship type is Many to Many, use associated id, otherwise use intersect id
  intersectId: string;
  associatedId: string;
  associatedName: string;
  optionText: string;
  selectedOptionText: string;
  iconSrc?: string;
  group?: string;
  entity: ComponentFramework.WebApi.Entity;
}

export interface EntityReference {
  etn: string;
  id: string;
  name: string;
}

export enum ShowIconOptions {
  None = 0,
  EntityIcon = 1,
  RecordImage = 2,
}
