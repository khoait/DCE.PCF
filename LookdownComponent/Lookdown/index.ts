import { createElement } from "react";
import LookdownControl from "./components/LookdownControl";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { IExtendedContext } from "./types/extendedContext";
import { LanguagePack } from "./types/languagePack";
import { IconSizes, LookdownControlProps, OpenRecordMode, ShowIconOptions } from "./types/typings";

export class Lookdown implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
  private notifyOutputChanged: () => void;
  private output: ComponentFramework.LookupValue | null;
  private context: IExtendedContext;
  private languagePack: LanguagePack;

  /**
   * Empty constructor.
   */
  // eslint-disable-next-line @typescript-eslint/no-empty-function
  constructor() {}

  /**
   * Used to initialize the control instance. Controls can kick off remote server calls and other initialization actions here.
   * Data-set values are not initialized here, use updateView.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to property names defined in the manifest, as well as utility functions.
   * @param notifyOutputChanged A callback method to alert the framework that the control has new outputs ready to be retrieved asynchronously.
   * @param state A piece of data that persists in one session for a single user. Can be set at any point in a controls life cycle by calling 'setControlState' in the Mode interface.
   */
  public init(context: IExtendedContext, notifyOutputChanged: () => void, state: ComponentFramework.Dictionary): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    this.languagePack = {
      BlankValueLabel: context.resources.getString("BlankValueLabel"),
      EmptyListMessage: context.resources.getString("EmptyListMessage"),
      OpenRecordLabel: context.resources.getString("OpenRecordLabel"),
      QuickCreateLabel: context.resources.getString("QuickCreateLabel"),
      LookupPanelLabel: context.resources.getString("LookupPanelLabel"),
      LoadDataErrorMessage: context.resources.getString("LoadDataErrorMessage"),
    };
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   * @returns ReactElement root react element for the control
   */
  public updateView(context: IExtendedContext): React.ReactElement {
    this.context = context;

    const isAuthoringMode = typeof context.parameters.lookupField?.getViewId === "function" ? false : true;

    const props: LookdownControlProps = {
      lookupViewId: isAuthoringMode
        ? undefined
        : context.parameters.lookupField?.getViewId() ??
          context.utils._customControlProperties.descriptor.Parameters.DefaultViewId,
      lookupEntity: isAuthoringMode ? undefined : context.parameters.lookupField.getTargetEntityType(),
      selectedId: context.parameters.lookupField?.raw?.at(0)?.id,
      selectedText: context.parameters.lookupField?.raw?.at(0)?.name,
      customFilter: context.parameters.customFilter?.raw,
      groupBy: context.parameters.groupByField?.raw,
      optionTemplate: context.parameters.optionTemplate?.raw,
      selectedItemTemplate: context.parameters.selectedItemTemplate?.raw,
      showIcon: context.parameters.showIcon?.raw
        ? (Number.parseInt(context.parameters.showIcon.raw) as ShowIconOptions)
        : undefined,
      iconSize: context.parameters.iconSize?.raw
        ? (Number.parseInt(context.parameters.iconSize.raw) as IconSizes)
        : undefined,
      openRecordMode: context.parameters.commandOpenRecord?.raw
        ? (Number.parseInt(context.parameters.commandOpenRecord.raw) as OpenRecordMode)
        : undefined,
      allowQuickCreate: context.parameters.commandQuickCreate?.raw === "1",
      allowLookupPanel: context.parameters.commandQuickCreate?.raw === "1",
      disabled: context.mode.isControlDisabled,
      defaultLanguagePack: this.languagePack,
      languagePackPath: context.parameters.languagePackPath?.raw ?? undefined,
      fluentDesign: this.context.fluentDesignLanguage,
      onChange: (value) => {
        this.output = value;
        this.notifyOutputChanged();
      },
    };

    return createElement(LookdownControl, props);
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
      lookupField: this.output ? [this.output] : undefined,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }
}
