import React from "react";
import { Root, createRoot } from "react-dom/client";
import PolyLookupControl from "./components/PolyLookupControl";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IExtendedContext } from "./types/extendedContext";
import { LanguagePack } from "./types/languagePack";
import { PolyLookupProps, RelationshipTypeEnum, TagAction } from "./types/typings";
export class PolyLookup implements ComponentFramework.StandardControl<IInputs, IOutputs> {
  private container: HTMLDivElement;
  private root: Root;
  private notifyOutputChanged: () => void;
  private output: string | undefined;
  private outputSelectedItems: string | undefined;
  private context: IExtendedContext;
  private languagePack: LanguagePack;

  /**
   * Empty constructor.
   */
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  constructor() {
    // Empty constructor
  }

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: IExtendedContext,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    this.container = container;
    this.languagePack = {
      AddNewLabel: context.resources.getString("AddNewLabel"),
      ControlIsNotAvailableMessage: context.resources.getString("ControlIsNotAvailableMessage"),
      CreateFormNotSupportedMessage: context.resources.getString("CreateFormNotSupportedMessage"),
      EmptyListDefaultMessage: context.resources.getString("EmptyListDefaultMessage"),
      EmptyListMessage: context.resources.getString("EmptyListMessage"),
      LoadingMessage: context.resources.getString("LoadingMessage"),
      PlaceholderDefault: context.resources.getString("PlaceholderDefault"),
      Placeholder: context.resources.getString("Placeholder"),
      RelationshipNotSupportedMessage: context.resources.getString("RelationshipNotSupportedMessage"),
      SuggestionListHeaderDefaultLabel: context.resources.getString("SuggestionListHeaderDefaultLabel"),
      SuggestionListHeaderLabel: context.resources.getString("SuggestionListHeaderLabel"),
      LoadMoreLabel: context.resources.getString("LoadMoreLabel"),
      NoMoreRecordsMessage: context.resources.getString("NoMoreRecordsMessage"),
      SuggestionListFullMessage: context.resources.getString("SuggestionListFullMessage"),
      GenericErrorMessage: context.resources.getString("GenericErrorMessage"),
      InvalidLookupViewMessage: context.resources.getString("InvalidLookupViewMessage"),
    };
    this.root = createRoot(this.container, {
      identifierPrefix: "DCEPCF-PolyLookup",
    });
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   */
  public updateView(context: IExtendedContext): void {
    this.context = context;
    this.render();
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
      boundField: this.output,
      outputField: this.outputSelectedItems,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
    this.root.unmount();
  }

  public render(): void {
    let clientUrl = "";

    try {
      clientUrl = this.context.page.getClientUrl();
    } catch {
      // ignore error
    }

    const props: PolyLookupProps = {
      currentTable: this.context.page.entityTypeName,
      currentRecordId: this.context.page.entityId,
      relationshipName: this.context.parameters.relationship?.raw ?? "",
      relationship2Name: this.context.parameters.relationship2?.raw ?? undefined,
      relationshipType: Number.parseInt(this.context.parameters.relationshipType?.raw) as RelationshipTypeEnum,
      clientUrl: clientUrl,
      lookupView: this.context.parameters.lookupView?.raw ?? undefined,
      itemLimit: this.context.parameters.itemLimit?.raw ?? undefined,
      pageSize: this.context.userSettings?.pagingLimit ?? undefined,
      disabled: this.context.mode.isControlDisabled,
      formType: typeof Xrm === "undefined" ? undefined : Xrm.Page.ui.getFormType(),
      outputSelectedItems: !!this.context.parameters.outputField?.attributes?.LogicalName,
      tagAction: this.context.parameters.tagAction?.raw
        ? (Number.parseInt(this.context.parameters.tagAction.raw) as TagAction)
        : undefined,
      defaultLanguagePack: this.languagePack,
      languagePackPath: this.context.parameters.languagePackPath?.raw ?? undefined,
      fluentDesign: this.context.fluentDesignLanguage,
      onChange:
        this.context.parameters.outputSelected?.raw === "1" ||
        this.context.parameters.outputField?.attributes?.LogicalName
          ? this.onLookupChange
          : undefined,
      onQuickCreate: this.context.parameters.allowQuickCreate?.raw === "1" ? this.onQuickCreate : undefined,
    };
    this.root.render(React.createElement(PolyLookupControl, props));
  }

  public onLookupChange = (selectedItems: ComponentFramework.EntityReference[] | undefined) => {
    if (this.context.parameters.outputSelected.raw === "1") {
      this.output = selectedItems?.map((item) => item.name).join(", ");
    }

    if (this.context.parameters.outputField.attributes) {
      if (!selectedItems?.length) {
        this.outputSelectedItems = "";
      } else {
        this.outputSelectedItems = JSON.stringify(selectedItems);
      }
    }
    this.notifyOutputChanged();
  };

  public onQuickCreate = async (
    entityName: string | undefined,
    primaryAttribute: string | undefined,
    value: string | undefined,
    useQuickCreateForm: boolean | undefined
  ) => {
    if (entityName && primaryAttribute) {
      let result: ComponentFramework.NavigationApi.OpenFormSuccessResponse;
      if (useQuickCreateForm) {
        result = await this.context.navigation.openForm(
          {
            entityName: entityName,
            useQuickCreateForm: true,
          },
          {
            [primaryAttribute]: value ?? "",
          }
        );
      } else {
        result = await this.context.navigation.navigateTo(
          {
            pageType: "entityrecord",
            entityName: entityName,
            data: {
              [primaryAttribute]: value ?? "",
            },
          },
          {
            target: 2,
            height: { value: 80, unit: "%" },
            width: { value: 70, unit: "%" },
            position: 1,
          }
        );
      }

      if (result?.savedEntityReference?.length) {
        return result.savedEntityReference[0].id;
      }
    }

    return undefined;
  };
}
