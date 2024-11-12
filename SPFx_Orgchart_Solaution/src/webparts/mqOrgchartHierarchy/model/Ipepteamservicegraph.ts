
import { IUserPresence } from "../components/user/IUserPresense";

export interface Ipepteamservicegraph {
    getUserId: (useremail : string) =>   Promise<string>;
    getPresence: (userId: string) => Promise<IUserPresence>;
}