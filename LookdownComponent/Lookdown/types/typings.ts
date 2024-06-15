import { LanguagePack } from "./languagePack";

export interface LookdownControlProps {
  lookupViewId?: string | null;
  lookupEntity?: string | null;
  selectedId?: string;
  selectedText?: string;
  customFilter?: string | null;
  groupBy?: string | null;
  optionTemplate?: string | null; // example of optionTemplate: {{ fullname }} - {{ emailaddress1 }}
  selectedItemTemplate?: string | null; // example of selectedItemTemplate: {{ fullname }} - {{ emailaddress1 }}
  showIcon?: ShowIconOptions;
  iconSize?: IconSizes;
  openRecordMode?: OpenRecordMode;
  allowQuickCreate?: boolean;
  allowLookupPanel?: boolean;
  disabled?: boolean;
  defaultLanguagePack: LanguagePack;
  languagePackPath?: string;
  fluentDesign?: ComponentFramework.FluentDesignState;
  onChange?: (selectedItem: ComponentFramework.LookupValue | null) => void;
}

export enum ShowIconOptions {
  None = 0,
  EntityIcon = 1,
  RecordImage = 2,
}

export enum IconSizes {
  Normal = 0,
  Large = 1,
}

export enum OpenRecordMode {
  None = 0,
  Inline = 1,
  Dialog = 2,
}

export interface EntityOption {
  id: string;
  primaryName: string;
  optionText: string;
  selectedOptionText: string;
  iconSrc?: string;
  iconSize?: number;
  group?: string;
}
