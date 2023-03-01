import { IInputs, IOutputs } from "./generated/ManifestTypes";
import * as React from "react";
import PolyLookupControl, { PolyLookupProps, RelationshipTypeEnum } from "./components/PolyLookupControl";

interface IExtendedUserSettings extends ComponentFramework.UserSettings {
  pagingLimit: number;
}

interface IExtendedContext extends ComponentFramework.Context<IInputs> {
  page: {
    appId: string;
    entityTypeName: string;
    entityId: string;
    isPageReadOnly: boolean;
    getClientUrl: () => string;
  };
  userSettings: IExtendedUserSettings;
}

export class PolyLookup implements ComponentFramework.ReactControl<IInputs, IOutputs> {
  private theComponent: ComponentFramework.ReactControl<IInputs, IOutputs>;
  private notifyOutputChanged: () => void;
  output: string | undefined;

  /**
   * Empty constructor.
   */
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
  }

  /**
   * Called when any value in the property bag has changed. This includes field values, data-sets, global values such as container height and width, offline status, control metadata values such as label, visible, etc.
   * @param context The entire property bag available to control via Context Object; It contains values as set up by the customizer mapped to names defined in the manifest, as well as utility functions
   * @returns ReactElement root react element for the control
   */
  public updateView(context: IExtendedContext): React.ReactElement {
    const props: PolyLookupProps = {
      currentTable: context.page.entityTypeName,
      currentRecordId: context.page.entityId,
      relationshipName: context.parameters.relationship.raw ?? "",
      relationshipType:
        RelationshipTypeEnum[context.parameters.relationshipType.raw as keyof typeof RelationshipTypeEnum],
      clientUrl: context.page.getClientUrl(),
      sortBy: context.parameters.sortAttribute.raw ?? undefined,
      pageSize: context.userSettings.pagingLimit ?? undefined,
      disabled: context.mode.isControlDisabled,
      onChange: context.parameters.outputSelected.raw === "1" ? this.onLookupChange : undefined,
    };
    return React.createElement(PolyLookupControl, props);
  }

  /**
   * It is called by the framework prior to a control receiving new data.
   * @returns an object based on nomenclature defined in manifest, expecting object[s] for property marked as “bound” or “output”
   */
  public getOutputs(): IOutputs {
    return {
      boundField: this.output,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    // Add code to cleanup control if necessary
  }

  public onLookupChange = (value: string | undefined) => {
    this.output = value;
    this.notifyOutputChanged();
  };
}
