import * as React from "react";
import LookdownControl, { IconSizes, OpenRecordMode, ShowIconOptions } from "./components/LookdownControl";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { LanguagePack } from "./types/languagePack";

export class Lookdown implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
  private notifyOutputChanged: () => void;
  private output: ComponentFramework.LookupValue | null;
  private context: ComponentFramework.Context<IInputs>;
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
  public init(
    context: ComponentFramework.Context<IInputs>,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary
  ): void {
    this.notifyOutputChanged = notifyOutputChanged;
    this.context = context;
    this.languagePack = {
      BlankValueLabel: context.resources.getString("BlankValueLabel"),
      EmptyListMessage: context.resources.getString("EmptyListMessage"),
      OpenRecordLabel: context.resources.getString("OpenRecordLabel"),
      QuickCreateLabel: context.resources.getString("QuickCreateLabel"),
      LookupPanelLabel: context.resources.getString("LookupPanelLabel"),
    };
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   * @returns ReactElement root react element for the control
   */
  public updateView(context: ComponentFramework.Context<IInputs>): React.ReactElement {
    this.context = context;

    const isAuthoringMode = typeof context.parameters.lookupField?.getViewId === "function" ? false : true;

    return React.createElement(LookdownControl, {
      lookupViewId: isAuthoringMode ? undefined : context.parameters.lookupField.getViewId(),
      lookupEntity: isAuthoringMode ? undefined : context.parameters.lookupField.getTargetEntityType(),
      selectedId: context.parameters.lookupField?.raw?.at(0)?.id,
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
      onChange: (value) => {
        this.output = value;
        this.notifyOutputChanged();
      },
    });
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
