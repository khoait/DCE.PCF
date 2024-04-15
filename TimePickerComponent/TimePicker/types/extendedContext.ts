import { IInputs } from "../generated/ManifestTypes";

interface IExtendedUserSettings extends ComponentFramework.UserSettings {
  timeZoneUtcOffsetMinutes: number;
}

export interface IExtendedContext extends ComponentFramework.Context<IInputs> {
  userSettings: IExtendedUserSettings;
}
