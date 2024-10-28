
import { Ipepteamservicegraph } from "../model/Ipepteamservicegraph";
//import { WebPartContext } from "@microsoft/sp-webpart-base";
import { GraphFI, graphfi } from "@pnp/graph";
import "@pnp/graph/users";
import { IUserPresence } from "../components/user/IUserPresense";

export class Pepteamservicegraph implements Ipepteamservicegraph {

  private _graph : GraphFI ;
  constructor() {
     //this._graph =  graphfi(...).using(SPFx(this.context));
     this._graph = graphfi().using();
  }

  public async getUserId(useremail?: string): Promise<string> {
    return new Promise<string>(async (resolve, reject) => {
      try {
        if (useremail)
        {
        let userid: string = await this._graph.users.getById(useremail)();
        resolve(userid);
        }
        else
        {
          resolve("empty");
        }
      }
      catch (error) {
        reject(error); // Reject the promise with the error
      }
    });
} 
  
public async getPresence(userId: string): Promise<IUserPresence> {
  let MQUserpresenseInstance : IUserPresence;

  return new Promise<IUserPresence>(async (resolve, reject) => {
    try {
      const MQUserpresenseresponse = await  this._graph.users.getById(userId).presence();
      MQUserpresenseInstance = {
        Activity :  MQUserpresenseresponse.activity?.toString(),
        Availability: MQUserpresenseresponse.availability?.toString()
      }
      resolve(MQUserpresenseInstance);
    }
    catch (error) {
      reject(error); // Reject the promise with the error
    }
  });
}

}
