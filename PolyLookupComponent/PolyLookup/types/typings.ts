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
  tagAction?: TagAction;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  fluentDesign?: ComponentFramework.FluentDesignState;
  onChange?: (selectedItems: ComponentFramework.EntityReference[] | undefined) => void;
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
