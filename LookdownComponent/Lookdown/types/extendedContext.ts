import { IInputs } from "../generated/ManifestTypes";

interface IExtendedUtility extends ComponentFramework.Utility {
  _customControlProperties: {
    descriptor: {
      Parameters: {
        DefaultViewId: string;
      };
    };
  };
}

export interface IExtendedContext extends ComponentFramework.Context<IInputs> {
  utils: IExtendedUtility;
}
