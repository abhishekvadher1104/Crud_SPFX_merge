import { SPFI, spfi } from "@pnp/sp";
import { SPFx } from "@pnp/sp/presets/all";
import { WebPartContext } from "@microsoft/sp-webpart-base";

let _sp: SPFI | undefined;

/**
 * Get the SPFI instance
 */
export const getSP = (): SPFI => {
  if (!_sp) {
    throw new Error("PnpJS not initialized. Call initSPFx(context) first.");
  }
  return _sp;
};

/**
 * Initialize PnP JS with SPFx context
 */
export const initSPFx = (context: WebPartContext): void => {
  _sp = spfi().using(SPFx(context));
};
