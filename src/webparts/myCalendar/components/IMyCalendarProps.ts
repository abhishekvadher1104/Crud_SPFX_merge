import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMyCalendarProps {
  context: WebPartContext;
  description: string; // ADD THIS LINE
}
