import { IUser } from "./Iuser";

export interface IUsercollection extends IUser
{
    name: string;
    projectRole: string;
    superiorRole:string;
    dateFrom: Date;
    dateTo: Date;
    userimage: string;
    children?: IUser[];
    
}