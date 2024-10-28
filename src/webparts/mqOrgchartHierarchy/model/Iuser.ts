export interface IUser {
    name: string;
    email:string;
    projectRole: string;
    superiorRole:string;
    dateFrom: Date;
    dateTo: Date;
    userimage: string;
    children?: IUser[];
    predecessor?: IUser[];
  }
  