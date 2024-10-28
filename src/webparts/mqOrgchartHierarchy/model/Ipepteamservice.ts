
import { IUser } from "./Iuser";

export interface Ipepteamservice {
    getAllTeamUsers: (pepweburl : string) =>  Promise<IUser[]>;
}