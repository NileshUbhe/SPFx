import "@pnp/graph/users";
import "@pnp/graph/cloud-communications";
import { WebPartContext } from "@microsoft/sp-webpart-base";
import { IUserPresence } from "./IUserPresense";
import * as React from "react";
import { Pepteamservicegraph } from "../../services/Pepteamservicegraph";

export interface IUserPresenceProps {
    context: WebPartContext;
    Availability?: string;
    Activity?:string;
    mqUserEmail?:string;
  }

  export default class MQUserPresense extends React.Component<IUserPresence,IUserPresenceProps, {}> 
  {
  private peptmsrvinst: Pepteamservicegraph; 
   constructor(props: IUserPresence)
    {
      super(props);
      this.state={
         mqUserEmail  : this.props.MqUserEmail,
         context : this.context
      };
      //this.peptmsrvinst = new Pepteamservicegraph(this.context);
      this.peptmsrvinst = new Pepteamservicegraph();
    }
    public componentDidMount(): void {
      this.getUserPresenseafterpromise();
    }

    private getUserPresenseafterpromise():void
    {
      this.peptmsrvinst
      .getUserId(this.props.MqUserEmail)
      .then((usersid: string) => {
        this.peptmsrvinst.getPresence(usersid)  
        .then((userpresense: IUserPresence) => {
          this.setState({Activity : userpresense.Activity, Availability: userpresense.Availability });
        });
      });
    }

   
    public render(): React.ReactElement<IUserPresence> {
        return (
          <div>
              <span> Availability{this.props.MqUserEmail}{this.state.Availability} </span>
          </div>
        );
      }
  }