
import { WebPartContext } from "@microsoft/sp-webpart-base";

export interface IMqOrgchartHierarchyProps {
  description: string;
  isDarkTheme: boolean;
  context: WebPartContext,
  environmentMessage: string;
  hasTeamsContext: boolean;
  userDisplayName: string;
}

export interface IZoomPanFunctions {
  zoomIn: () => void;
  zoomOut: () => void;
  resetTransform: () => void;
}