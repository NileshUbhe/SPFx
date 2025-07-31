
import { IEmployee } from "./IEmployee";



export interface IConsumerWebPartDemoState {
    status: string;
    EmployeeListItems: IEmployee[];
    EmployeeListItem: IEmployee;  
    DeptTitleId: string ;  
  }