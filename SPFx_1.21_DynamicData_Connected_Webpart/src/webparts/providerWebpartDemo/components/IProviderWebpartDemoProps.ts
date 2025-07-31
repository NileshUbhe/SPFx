import { WebPartContext } from "@microsoft/sp-webpart-base";
import { DepartmentSelectedCallback } from './DepartmentSelectedCallBack'


export interface IProviderWebPartDemoProps {
  description: string;
  context: WebPartContext;
  siteUrl: string;
  onDepartmentSelected: DepartmentSelectedCallback;
}
