import React from "react";
import { createRoot, Root } from "react-dom/client";
import TimePickerControl from "./components/TimePickerControl";
import TimePickerControlNewLook from "./components/TimePickerControlNewLook";
import { IInputs, IOutputs } from "./generated/ManifestTypes";
import { convertDate, getUtcDate } from "./services/dateService";
import { IExtendedContext } from "./types/extendedContext";
import { TimePickerControlProps } from "./types/typings";

export class TimePicker
  implements ComponentFramework.StandardControl<IInputs, IOutputs>
{
  private context: IExtendedContext;
  private container: HTMLDivElement;
  private notifyOutputChanged: () => void;
  private root: Root;
  private outputValue: Date | null;
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
   * @param container If a control is marked control-type='standard', it will receive an empty div element within which it can render its content.
   */
  public init(
    context: IExtendedContext,
    notifyOutputChanged: () => void,
    state: ComponentFramework.Dictionary,
    container: HTMLDivElement
  ): void {
    this.context = context;
    this.container = container;
    this.notifyOutputChanged = notifyOutputChanged;
    this.root = createRoot(this.container);
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
      boundField: this.outputValue ?? undefined,
    };
  }

  /**
   * Called when the control is to be removed from the DOM tree. Controls should use this call for cleanup.
   * i.e. cancelling any pending remote calls, removing listeners, etc.
   */
  public destroy(): void {
    this.root.unmount();
  }

  public render(): void {
    const isNewLook = !!this.context.fluentDesignLanguage;

    const boundValue = this.convertToLocalDate(
      this.context.parameters.boundField
    );
    const dateAnchor = this.convertToLocalDate(
      this.context.parameters.dateAnchor
    );

    const props: TimePickerControlProps = {
      inputValue: boundValue,
      dateAnchor: dateAnchor,
      disabled: this.context.mode.isControlDisabled,
      placeholder: this.context.parameters.placeholder.raw ?? undefined,
      increment: this.context.parameters.increment.raw ?? undefined,
      hourCycle12: this.context.parameters.hourCycle12.raw === "1",
      freeform: this.context.parameters.freeform.raw === "1",
      startHour: this.context.parameters.startHour.raw ?? undefined,
      endHour: this.context.parameters.endHour.raw ?? undefined,
      onTimeChange: (time) => {
        this.outputValue = time;
        this.notifyOutputChanged();
      },
    };

    const key = boundValue ? boundValue.getTime() : 0;

    this.root.render(
      isNewLook
        ? React.createElement(TimePickerControlNewLook, {
            key,
            theme: this.context.fluentDesignLanguage?.tokenTheme,
            ...props,
          })
        : React.createElement(TimePickerControl, {
            key,
            ...props,
          })
    );
  }

  public convertToLocalDate(
    dateProperty: ComponentFramework.PropertyTypes.DateTimeProperty
  ) {
    if (dateProperty.attributes?.Behavior === 1) {
      // user local
      if (!dateProperty.raw) return null;

      return convertDate(
        dateProperty.raw,
        this.context.userSettings.getTimeZoneOffsetMinutes(dateProperty.raw)
      );
    } else {
      return getUtcDate(dateProperty.raw);
    }
  }
}
