import { IUser } from "../../model/Iuser";

export default interface IUserProps {
  MQUserData : IUser;
}

export interface IUserPredecessor
{
  MQPredecessorData : IUser[];
  showPredecessor:boolean;
}